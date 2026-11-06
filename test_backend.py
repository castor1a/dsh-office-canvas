"""Round-trip and adversarial tests for the Office canvas backend."""
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import warnings
import zipfile

from docx import Document
from pptx import Presentation
from pptx.util import Inches
from lxml import etree
import backend


class BackendTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(dir=Path(__file__).parent)
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def request(self, path, action='read', **kwargs):
        return backend.handle({'action': action, 'workspaceRoot': str(self.root), 'path': str(path), **kwargs})

    def fixture(self, kind):
        path = self.root / ('source.' + kind)
        if kind == 'docx':
            doc = Document()
            p = doc.add_paragraph()
            p.add_run('Original ').bold = True
            p.add_run('mixed text').italic = True
            doc.add_table(rows=1, cols=1).cell(0, 0).text = 'Table text'
            doc.save(path)
        else:
            deck = Presentation()
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
            p = box.text_frame.paragraphs[0]
            p.add_run().text = 'Original '
            p.add_run().text = 'mixed text'
            deck.save(path)
        with zipfile.ZipFile(path, 'a') as archive:
            archive.writestr('custom/preserve.bin', b'untouched\x00\xffpayload')
        return path

    def assert_unchanged_parts(self, before, after, changed):
        with zipfile.ZipFile(before) as a, zipfile.ZipFile(after) as b:
            self.assertEqual(a.namelist(), b.namelist())
            self.assertEqual(len(b.namelist()), len(set(b.namelist())))
            for name in a.namelist():
                if name not in changed:
                    self.assertEqual(a.read(name), b.read(name), name)
            self.assertEqual(b.read('custom/preserve.bin'), b'untouched\x00\xffpayload')

    def test_docx_targeted_run_format_table_and_preservation(self):
        path = self.fixture('docx')
        model = self.request(path)
        self.assertEqual([p['text'] for p in model['paragraphs']], ['Original mixed text', 'Table text'])
        self.assertTrue(model['paragraphs'][0]['runs'][0]['bold'])
        saved = self.request(path, 'save', expectedHash=model['hash'], operations=[
            {'type': 'text', 'id': 'r:0:1', 'text': 'updated\ttext\nnext'},
            {'type': 'format', 'id': 'r:0:1', 'bold': True, 'color': '#aabbcc', 'size': 22},
            {'type': 'format', 'id': 'p:0', 'align': 'center'},
            {'type': 'text', 'id': 'p:1', 'text': 'Edited cell'},
        ])
        reread = self.request(saved['path'])
        run = reread['paragraphs'][0]['runs'][1]
        self.assertEqual(run['text'], 'updated\ttext\nnext')
        self.assertEqual((run['bold'], run['italic'], run['size'], run['color']), (True, True, 22, 'AABBCC'))
        self.assertEqual(reread['paragraphs'][0]['align'], 'center')
        doc = Document(saved['path'])
        self.assertEqual(doc.tables[0].cell(0, 0).text, 'Edited cell')
        self.assert_unchanged_parts(path, saved['path'], {'word/document.xml'})
        self.assertEqual(self.request(path)['hash'], model['hash'])
        second = self.request(path, 'save', expectedHash=model['hash'], operations=[])
        self.assertTrue(second['path'].endswith('-edited-2.docx'))

    def test_pptx_run_format_and_move_preserves_other_parts(self):
        path = self.fixture('pptx')
        model = self.request(path)
        shape = model['slides'][0]['elements'][0]
        run_id = shape['paragraphs'][0]['runs'][0]['id']
        saved = self.request(path, 'save', expectedHash=model['hash'], operations=[
            {'type': 'text', 'id': run_id, 'text': 'Edited'},
            {'type': 'format', 'id': run_id, 'bold': True, 'underline': True, 'size': 26, 'color': 'FF0000'},
            {'type': 'move', 'id': shape['id'], 'x': 2000000, 'y': 1000000, 'width': 3000000, 'height': 900000},
        ])
        reopened = Presentation(saved['path'])
        changed = reopened.slides[0].shapes[0]
        self.assertEqual(changed.text, 'Editedmixed text')
        self.assertEqual((changed.left, changed.top, changed.width), (2000000, 1000000, 3000000))
        self.assertEqual(changed.text_frame.paragraphs[0].runs[0].font.size.pt, 26)
        self.assertTrue(changed.text_frame.paragraphs[0].runs[0].font.underline)
        self.assert_unchanged_parts(path, saved['path'], {'ppt/slides/slide1.xml'})

    def test_stale_hash_and_existing_output_require_hash(self):
        path = self.fixture('docx')
        model = self.request(path)
        with self.assertRaisesRegex(backend.BackendError, 'Source changed'):
            self.request(path, 'save', expectedHash='stale', operations=[])
        with self.assertRaises(backend.BackendError) as caught:
            self.request(path, 'save', expectedHash=model['hash'], outputPath=str(path), operations=[])
        self.assertEqual(caught.exception.code, 'output-conflict')
        saved = self.request(path, 'save', expectedHash=model['hash'], outputPath=str(path), expectedOutputHash=model['hash'], operations=[{'type': 'text', 'id': 'r:0:0', 'text': 'Replacement'}])
        self.assertEqual(saved['path'], str(path))
        self.assertNotEqual(saved['hash'], model['hash'])

    def test_traversal_symlink_and_bad_zip(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as outside:
            external = Path(outside) / 'outside.docx'
            Document().save(external)
            with self.assertRaises(backend.BackendError) as caught:
                self.request(external)
            self.assertEqual(caught.exception.code, 'outside-workspace')
            link = self.root / 'link.docx'
            try:
                link.symlink_to(external)
            except OSError:
                pass
            else:
                with self.assertRaises(backend.BackendError):
                    self.request(link)
        for name in ('../escape', '/absolute', 'C:/evil', 'bad\\name'):
            data = io.BytesIO()
            with zipfile.ZipFile(data, 'w') as archive:
                archive.writestr('[Content_Types].xml', '<Types/>')
                # ZipInfo's constructor normalizes backslashes on Windows; preserve
                # the requested adversarial member name in the actual ZIP bytes.
                info = zipfile.ZipInfo()
                info.filename = name
                archive.writestr(info, b'bad')
            with self.assertRaises(backend.BackendError):
                backend.Package(data.getvalue())

    def test_duplicate_members_bomb_entities_and_unknown_targets(self):
        data = io.BytesIO()
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            with zipfile.ZipFile(data, 'w') as archive:
                archive.writestr('[Content_Types].xml', '<Types/>')
                archive.writestr('same', b'1')
                archive.writestr('same', b'2')
        with self.assertRaises(backend.BackendError):
            backend.Package(data.getvalue())
        data = io.BytesIO()
        with zipfile.ZipFile(data, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr('[Content_Types].xml', '<Types/>')
            archive.writestr('bomb', b'0' * (2 * 1024 * 1024))
        with self.assertRaises(backend.BackendError) as caught:
            backend.Package(data.getvalue())
        self.assertEqual(caught.exception.code, 'zip-limit')
        path = self.fixture('docx')
        model = self.request(path)
        with self.assertRaises(backend.BackendError):
            self.request(path, 'save', expectedHash=model['hash'], operations=[{'type': 'text', 'id': 'missing', 'text': 'x'}])
        self.assertFalse((self.root / 'source-edited.docx').exists())
        data = io.BytesIO()
        with zipfile.ZipFile(data, 'w') as archive:
            archive.writestr('[Content_Types].xml', '<Types/>')
            archive.writestr('word/document.xml', '<!DOCTYPE test [<!ENTITY xxe SYSTEM "file:///secret">]><test>&xxe;</test>')
        pkg = backend.Package(data.getvalue())
        try:
            with self.assertRaises(backend.BackendError):
                pkg.xml('word/document.xml')
        finally:
            pkg.zip.close()

    def test_cli_json_envelope(self):
        path = self.fixture('docx')
        # Files as stdio avoid named-pipe use under the Windows sandbox.
        request_path = self.root / 'request.json'
        response_path = self.root / 'response.json'
        request_path.write_text(json.dumps({'action': 'read', 'path': str(path), 'workspaceRoot': str(self.root)}), encoding='utf-8')
        with request_path.open('rb') as source, response_path.open('wb') as output:
            subprocess.run([sys.executable, str(Path(backend.__file__))], stdin=source, stdout=output, check=True)
        response = json.loads(response_path.read_text())
        self.assertTrue(response['ok'])
        self.assertEqual(response['value']['kind'], 'docx')


class MemoryTests(unittest.TestCase):
    """Core serialization checks independent of filesystem permission provisioning."""
    def package(self, parts):
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, 'w') as archive:
            archive.writestr('[Content_Types].xml', '<Types/>')
            for name, value in parts.items():
                archive.writestr(name, value)
        pkg = backend.Package(stream.getvalue())
        self.addCleanup(pkg.zip.close)
        return pkg

    def test_run_property_schema_order(self):
        word = etree.fromstring(f'''<w:r xmlns:w="{backend.NS['w']}"><w:rPr>
          <w:rStyle w:val="Strong"/><w:rFonts w:ascii="Arial"/>
          <w:lang w:val="en-US"/><w:rPrChange w:id="1"/>
          </w:rPr><w:t>text</w:t></w:r>''')
        backend.format_node(word, 'run', 'w', {'bold': True, 'italic': True,
                            'underline': True, 'color': '123456', 'size': 18})
        self.assertEqual([etree.QName(x).localname for x in word[0]],
                         ['rStyle', 'rFonts', 'b', 'i', 'color', 'sz', 'u', 'lang', 'rPrChange'])
        drawing = etree.fromstring(f'''<a:r xmlns:a="{backend.NS['a']}"><a:rPr>
          <a:ln/><a:gradFill/><a:effectLst/><a:latin typeface="Arial"/>
          <a:hlinkClick/><a:extLst/></a:rPr><a:t>text</a:t></a:r>''')
        backend.format_node(drawing, 'run', 'a', {'color': '123456', 'underline': False})
        self.assertEqual([etree.QName(x).localname for x in drawing[0]],
                         ['ln', 'solidFill', 'effectLst', 'latin', 'hlinkClick', 'extLst'])
        self.assertEqual(drawing[0].get('u'), 'none')
        self.assertFalse(backend.run_model(drawing, 'a', 'r')['underline'])

    def test_docx_drawing_textbox_owns_its_text(self):
        pkg = self.package({'word/document.xml': f'''<w:document xmlns:w="{backend.NS['w']}">
          <w:body><w:p><w:r><w:t>Outer</w:t><w:drawing><w:txbxContent>
          <w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:txbxContent></w:drawing>
          <w:t> tail</w:t></w:r></w:p></w:body></w:document>'''})
        model, targets = backend.build_model(pkg, Path('test.docx'))
        self.assertEqual([p['text'] for p in model['paragraphs']], ['Outer tail', 'Inner'])
        self.assertEqual(model['paragraphs'][0]['runs'][0]['text'], 'Outer tail')
        backend.apply_operations(pkg, targets, [{'type': 'text', 'id': 'p:0', 'text': 'Changed'}])
        model, _ = backend.build_model(pkg, Path('test.docx'))
        self.assertEqual([p['text'] for p in model['paragraphs']], ['Changed', 'Inner'])
        self.assertIsNotNone(pkg.xml('word/document.xml').find('.//w:drawing', backend.NS))

    def test_pptx_multi_paragraph_whole_shape_refused_without_mutation(self):
        deck = Presentation()
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
        box.text = 'First\nSecond\n'
        stream = io.BytesIO()
        deck.save(stream)
        pkg = backend.Package(stream.getvalue())
        self.addCleanup(pkg.zip.close)
        model, targets = backend.build_model(pkg, Path('test.pptx'))
        shape = model['slides'][0]['elements'][0]
        before = etree.tostring(pkg.xml('ppt/slides/slide1.xml'))
        with self.assertRaisesRegex(backend.BackendError, 'edit individual paragraphs or runs') as caught:
            backend.apply_operations(pkg, targets, [{'type': 'text', 'id': shape['id'], 'text': 'Replacement'}])
        self.assertEqual(caught.exception.code, 'unsupported-operation')
        self.assertEqual(etree.tostring(pkg.xml('ppt/slides/slide1.xml')), before)
        self.assertFalse(pkg.changed)
        backend.apply_operations(pkg, targets, [
            {'type': 'text', 'id': shape['paragraphs'][0]['id'], 'text': 'Updated'},
            {'type': 'text', 'id': shape['paragraphs'][1]['runs'][0]['id'], 'text': 'Kept'},
        ])
        output = io.BytesIO()
        pkg.save(output)
        self.assertEqual(Presentation(output).slides[0].shapes[0].text, 'Updated\nKept\n')

    def test_pptx_group_descendants_are_not_flattened(self):
        deck = Presentation()
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        group = slide.shapes.add_group_shape()
        nested = group.shapes.add_group_shape()
        box = nested.shapes.add_textbox(Inches(1), Inches(2), Inches(3), Inches(1))
        box.text = 'Grouped'
        group.left, group.top = Inches(4), Inches(5)
        stream = io.BytesIO()
        deck.save(stream)
        pkg = backend.Package(stream.getvalue())
        self.addCleanup(pkg.zip.close)
        model, targets = backend.build_model(pkg, Path('test.pptx'))
        elements = model['slides'][0]['elements']
        self.assertEqual(len(elements), 1)
        self.assertEqual((elements[0]['kind'], elements[0]['editable']), ('unsupported', False))
        self.assertEqual((elements[0]['x'], elements[0]['y']), (Inches(4), Inches(5)))
        self.assertIn('Grouped shapes', elements[0]['reason'])
        self.assertFalse(targets)

    def test_pptx_placeholder_layout_and_master_geometry(self):
        deck = Presentation()
        slide = deck.slides.add_slide(deck.slide_layouts[0])
        slide.shapes.title.text = 'Inherited title'
        stream = io.BytesIO()
        deck.save(stream)
        pkg = backend.Package(stream.getvalue())
        self.addCleanup(pkg.zip.close)
        model, targets = backend.build_model(pkg, Path('test.pptx'))
        title = model['slides'][0]['elements'][0]
        shape = slide.shapes.title
        self.assertEqual((title['x'], title['y'], title['width'], title['height']),
                         (shape.left, shape.top, shape.width, shape.height))
        self.assertTrue(title['editable'])
        self.assertFalse(title['movable'])
        self.assertIn(title['id'], targets)
        # Remove layout geometry and prove the base title placeholder resolves on master.
        part = backend.related_part(pkg, 'ppt/slides/slide1.xml', 'slideLayout')
        layout_title = pkg.xml(part).find('p:cSld/p:spTree/p:sp', backend.NS)
        transform = backend.shape_transform(layout_title)
        transform.getparent().remove(transform)
        master = pkg.xml(backend.related_part(pkg, part, 'slideMaster'))
        master_title = next(s for s in master.findall('p:cSld/p:spTree/p:sp', backend.NS)
                            if backend.placeholder(s).get('type') == 'title')
        expected = backend.shape_geometry(pkg, 'unused', master_title)
        model, _ = backend.build_model(pkg, Path('test.pptx'))
        title = model['slides'][0]['elements'][0]
        self.assertEqual({key: title[key] for key in expected}, expected)
        self.assertTrue(title['renderable'])

    def test_pptx_missing_geometry_is_explicitly_uneditable(self):
        deck = Presentation()
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
        box.text = 'Unresolved'
        stream = io.BytesIO()
        deck.save(stream)
        pkg = backend.Package(stream.getvalue())
        self.addCleanup(pkg.zip.close)
        xml = pkg.xml('ppt/slides/slide1.xml')
        transform = xml.find('.//p:sp/p:spPr/a:xfrm', backend.NS)
        transform.getparent().remove(transform)
        model, targets = backend.build_model(pkg, Path('test.pptx'))
        shape = model['slides'][0]['elements'][0]
        self.assertFalse(shape['editable'])
        self.assertFalse(shape['renderable'])
        self.assertIsNone(shape['width'])
        self.assertIn('geometry', shape['reason'])
        self.assertEqual(shape['text'], 'Unresolved')
        self.assertFalse(targets)

    def test_unsafe_zip_members_are_rejected_in_memory(self):
        for name in ('../escape', '/absolute', 'C:/evil', 'bad\\name'):
            with self.subTest(name=name):
                stream = io.BytesIO()
                with zipfile.ZipFile(stream, 'w') as archive:
                    archive.writestr('[Content_Types].xml', '<Types/>')
                    info = zipfile.ZipInfo()
                    info.filename = name
                    archive.writestr(info, b'bad')
                with self.assertRaises(backend.BackendError):
                    pkg = backend.Package(stream.getvalue())
                    pkg.zip.close()

    def test_docx_memory_roundtrip(self):
        stream = io.BytesIO()
        doc = Document()
        p = doc.add_paragraph()
        p.add_run('Before').bold = True
        p.add_run(' untouched').italic = True
        doc.save(stream)
        pkg = backend.Package(stream.getvalue())
        try:
            model, targets = backend.build_model(pkg, Path('test.docx'))
            backend.apply_operations(pkg, targets, [
                {'type': 'text', 'id': 'r:0:0', 'text': 'After\tchange\nline'},
                {'type': 'format', 'id': 'r:0:0', 'color': '#112233', 'size': 18},
            ])
            output = io.BytesIO()
            pkg.save(output)
            reopened = Document(output)
            self.assertEqual(reopened.paragraphs[0].text, 'After\tchange\nline untouched')
            self.assertTrue(reopened.paragraphs[0].runs[0].bold)
            self.assertTrue(reopened.paragraphs[0].runs[1].italic)
            self.assertEqual(reopened.paragraphs[0].runs[0].font.size.pt, 18)
            with zipfile.ZipFile(io.BytesIO(stream.getvalue())) as a, zipfile.ZipFile(io.BytesIO(output.getvalue())) as b:
                self.assertEqual(a.namelist(), b.namelist())
                for name in a.namelist():
                    if name != 'word/document.xml':
                        self.assertEqual(a.read(name), b.read(name))
        finally:
            pkg.zip.close()

    def test_pptx_memory_roundtrip(self):
        stream = io.BytesIO()
        deck = Presentation()
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
        box.text = 'Before'
        deck.save(stream)
        pkg = backend.Package(stream.getvalue())
        try:
            model, targets = backend.build_model(pkg, Path('test.pptx'))
            shape = model['slides'][0]['elements'][0]
            run_id = shape['paragraphs'][0]['runs'][0]['id']
            backend.apply_operations(pkg, targets, [
                {'type': 'text', 'id': run_id, 'text': 'After'},
                {'type': 'format', 'id': run_id, 'bold': True, 'underline': True, 'size': 26, 'color': 'FF0000'},
                {'type': 'move', 'id': shape['id'], 'x': 2000000, 'width': 3000000},
            ])
            output = io.BytesIO()
            pkg.save(output)
            reopened = Presentation(output)
            changed = reopened.slides[0].shapes[0]
            self.assertEqual(changed.text, 'After')
            self.assertEqual(changed.left, 2000000)
            self.assertEqual(changed.width, 3000000)
            self.assertEqual(changed.text_frame.paragraphs[0].runs[0].font.size.pt, 26)
            self.assertTrue(changed.text_frame.paragraphs[0].runs[0].font.underline)
            saved_pkg = backend.Package(output.getvalue())
            self.addCleanup(saved_pkg.zip.close)
            _, saved_targets = backend.build_model(saved_pkg, Path('test.pptx'))
            backend.apply_operations(saved_pkg, saved_targets, [
                {'type': 'format', 'id': run_id, 'underline': False},
            ])
            output = io.BytesIO()
            saved_pkg.save(output)
            self.assertFalse(Presentation(output).slides[0].shapes[0]
                             .text_frame.paragraphs[0].runs[0].font.underline)
        finally:
            pkg.zip.close()


if __name__ == '__main__':
    unittest.main()
