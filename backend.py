"""Bounded OOXML canvas backend. JSON on stdin, one JSON result on stdout.

Positions are OOXML EMUs; font sizes are points. Only changed XML parts are
serialized. All other ZIP member contents and metadata are copied verbatim.
"""
from __future__ import annotations

import base64
import hashlib
import json
import math
import mimetypes
import os
from pathlib import Path, PurePosixPath
import posixpath
import re
import sys
import tempfile
import zipfile
from lxml import etree as E

MAX_FILE = 64 * 1024 * 1024
MAX_TOTAL = 256 * 1024 * 1024
MAX_REQUEST = 4 * 1024 * 1024
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


class BackendError(ValueError):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


def fail(code, message):
    raise BackendError(code, message)


def q(prefix, local):
    return '{' + NS[prefix] + '}' + local


def digest(data):
    return hashlib.sha256(data).hexdigest()


def file_digest(path):
    hasher = hashlib.sha256()
    with path.open('rb') as source:
        while block := source.read(1024 * 1024):
            hasher.update(block)
    return hasher.hexdigest()


def scoped_path(value, root, must_exist=True):
    if not isinstance(value, str) or not value or '\x00' in value:
        fail('invalid-path', 'A nonempty file path is required')
    path = Path(value)
    if not path.is_absolute():
        path = root / path
    path = path.resolve(strict=must_exist)
    if not path.is_relative_to(root):
        fail('outside-workspace', 'File path must resolve inside workspaceRoot')
    if path.suffix.lower() not in ('.docx', '.pptx'):
        fail('unsupported-format', 'Only .docx and .pptx files are supported')
    if must_exist and not path.is_file():
        fail('invalid-path', 'Path is not a regular file')
    return path


class Package:
    def __init__(self, data):
        import io
        self.raw = data
        self.changed = set()
        self.xmls = {}
        try:
            self.zip = zipfile.ZipFile(io.BytesIO(data))
            self.infos = self.zip.infolist()
            if len(self.infos) > 10000:
                fail('zip-limit', 'Too many ZIP members')
            self.names = set()
            total = 0
            for info in self.infos:
                name = info.filename
                if name in self.names:
                    fail('invalid-zip', 'Duplicate ZIP member name')
                self.names.add(name)
                # ZipInfo normalizes Windows separators when reading the name.
                if any('\\' in value or value.startswith('/') or '..' in PurePosixPath(value).parts
                       or ':' in value for value in (name, info.orig_filename)):
                    fail('invalid-zip', 'Unsafe ZIP member name')
                if info.flag_bits & 1:
                    fail('invalid-zip', 'Encrypted ZIP members are unsupported')
                if (info.external_attr >> 16) & 0o170000 == 0o120000:
                    fail('invalid-zip', 'ZIP symbolic links are unsupported')
                total += info.file_size
                if info.file_size > MAX_FILE or total > MAX_TOTAL or info.file_size > max(info.compress_size, 1) * 1000:
                    fail('zip-limit', 'ZIP expansion limit exceeded')
            if '[Content_Types].xml' not in self.names:
                fail('invalid-zip', 'Missing OOXML content types')
        except (zipfile.BadZipFile, NotImplementedError) as exc:
            fail('invalid-zip', str(exc))

    def read(self, name):
        if name not in self.names:
            fail('invalid-package', 'Missing package part: ' + name)
        try:
            return self.zip.read(name)
        except (zipfile.BadZipFile, RuntimeError) as exc:
            fail('invalid-zip', str(exc))

    def xml(self, name):
        if name not in self.xmls:
            try:
                data = self.read(name)
                parser = E.XMLParser(resolve_entities=False, no_network=True, load_dtd=False, huge_tree=False)
                tree = E.fromstring(data, parser)
                if tree.getroottree().docinfo.doctype:
                    fail('invalid-xml', 'XML document types are unsupported')
                self.xmls[name] = tree
            except E.XMLSyntaxError as exc:
                fail('invalid-xml', str(exc))
        return self.xmls[name]

    def relationships(self, part):
        name = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
        if name not in self.names:
            return {}
        result = {}
        for rel in self.xml(name):
            if rel.get('TargetMode') == 'External':
                continue
            target = rel.get('Target', '')
            resolved = posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join(posixpath.dirname(part), target))
            if resolved.startswith('../') or '\\' in resolved:
                fail('invalid-package', 'Relationship escapes package')
            result[rel.get('Id')] = resolved
        return result

    def save(self, filename):
        with zipfile.ZipFile(filename, 'w') as out:
            out.comment = self.zip.comment
            for info in self.infos:
                content = E.tostring(self.xmls[info.filename], xml_declaration=True, encoding='UTF-8', standalone=True) if info.filename in self.changed else self.read(info.filename)
                out.writestr(info, content)


# CT_RPr and CT_TextCharacterProperties are ordered sequences, not bags.
PROPERTY_ORDER = {
    q('w', 'rPr'): ('rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
                    'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
                    'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
                    'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
                    'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
                    'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'),
    q('a', 'rPr'): ('ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill',
                    'grpFill', 'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn',
                    'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym', 'hlinkClick',
                    'hlinkMouseOver', 'rtl', 'extLst'),
}


def child(parent, prefix, name):
    found = parent.find(q(prefix, name))
    if found is None:
        found = E.Element(q(prefix, name))
        order = PROPERTY_ORDER.get(parent.tag)
        if order and name in order:
            rank = order.index(name)
            index = next((i for i, item in enumerate(parent)
                          if E.QName(item).namespace == NS[prefix]
                          and E.QName(item).localname in order
                          and order.index(E.QName(item).localname) > rank), len(parent))
            parent.insert(index, found)
        else:
            parent.insert(0, found)
    return found


def text_of(node, prefix):
    out = []
    # Text belongs to this run only; drawings can contain independent textbox runs.
    for item in node:
        if item.tag == q(prefix, 't'):
            out.append(item.text or '')
        elif item.tag == q(prefix, 'tab'):
            out.append('\t')
        elif item.tag in (q(prefix, 'br'), q(prefix, 'cr')):
            out.append('\n')
    return ''.join(out)


def run_model(run, prefix, ident):
    pr = run.find(q(prefix, 'rPr'))
    def prop(name):
        return pr.find(q(prefix, name)) if pr is not None else None
    def toggle(name):
        item = prop(name)
        return None if item is None else item.get(q('w', 'val'), '1') not in ('0', 'false', 'off')
    result = {'id': ident, 'text': text_of(run, prefix), 'bold': None, 'italic': None,
              'underline': None, 'size': None, 'color': None, 'font': None}
    if prefix == 'w':
        result.update(bold=toggle('b'), italic=toggle('i'))
        u, size, color, font = prop('u'), prop('sz'), prop('color'), prop('rFonts')
        result['underline'] = None if u is None else u.get(q('w', 'val'), 'single') != 'none'
        result['size'] = float(size.get(q('w', 'val'))) / 2 if size is not None else None
        result['color'] = color.get(q('w', 'val')) if color is not None else None
        result['font'] = font.get(q('w', 'ascii')) if font is not None else None
    elif pr is not None:
        result.update(bold=None if pr.get('b') is None else pr.get('b') == '1',
                      italic=None if pr.get('i') is None else pr.get('i') == '1',
                      underline=None if pr.get('u') is None else pr.get('u') != 'none',
                      size=float(pr.get('sz')) / 100 if pr.get('sz') else None)
        color = pr.find('a:solidFill/a:srgbClr', NS)
        font = pr.find('a:latin', NS)
        result['color'] = color.get('val') if color is not None else None
        result['font'] = font.get('typeface') if font is not None else None
    return result


def paragraphs(nodes, prefix, id_prefix, part, targets):
    result = []
    for n, para in enumerate(nodes):
        pid = f'p:{id_prefix}{n}'
        targets[pid] = (part, para, 'paragraph', prefix)
        # Exclude nested paragraphs (e.g. a Word textbox) from this paragraph's runs.
        runs = [r for r in para.iter(q(prefix, 'r')) if next((a for a in r.iterancestors() if a.tag == q(prefix, 'p')), None) is para]
        models = []
        for m, run in enumerate(runs):
            rid = f'r:{id_prefix}{n}:{m}'
            targets[rid] = (part, run, 'run', prefix)
            models.append(run_model(run, prefix, rid))
        pp = para.find(q(prefix, 'pPr'))
        align = None
        style = None
        if pp is not None:
            if prefix == 'w':
                jc, st = pp.find(q('w', 'jc')), pp.find(q('w', 'pStyle'))
                align = jc.get(q('w', 'val')) if jc is not None else None
                style = st.get(q('w', 'val')) if st is not None else None
            else:
                align = pp.get('algn')
        result.append({'id': pid, 'text': ''.join(r['text'] for r in models), 'runs': models, 'style': style, 'align': align})
    return result


def solid_fill(node):
    color = node.find('a:solidFill/a:srgbClr', NS) if node is not None else None
    return '#' + color.get('val') if color is not None else None


def related_part(pkg, part, relation):
    name = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
    if name not in pkg.names:
        return None
    rels = pkg.relationships(part)
    return next((rels.get(rel.get('Id')) for rel in pkg.xml(name)
                 if rel.get('Type', '').endswith('/' + relation)
                 and rel.get('TargetMode') != 'External'), None)


def shape_transform(shape):
    for path in ('p:spPr/a:xfrm', 'p:xfrm', 'p:grpSpPr/a:xfrm'):
        transform = shape.find(path, NS)
        if transform is not None:
            return transform
    return None


def placeholder(shape):
    return shape.find('p:nvSpPr/p:nvPr/p:ph', NS)


def inherited_transforms(pkg, part, shape):
    """Resolve slide placeholder -> layout by idx -> master by base type."""
    transforms = [shape_transform(shape)]
    ph = placeholder(shape)
    if ph is None:
        return transforms
    layout_part = related_part(pkg, part, 'slideLayout')
    if not layout_part:
        return transforms
    layout = pkg.xml(layout_part)
    matches = [candidate for candidate in layout.findall('p:cSld/p:spTree/p:sp', NS)
               if placeholder(candidate) is not None
               and placeholder(candidate).get('idx', '0') == ph.get('idx', '0')]
    if len(matches) != 1:
        return transforms
    layout_shape = matches[0]
    transforms.append(shape_transform(layout_shape))
    master_part = related_part(pkg, layout_part, 'slideMaster')
    if master_part:
        ph_type = placeholder(layout_shape).get('type', 'obj')
        base_type = ('title' if ph_type == 'ctrTitle' else 'body'
                     if ph_type in ('obj', 'subTitle', 'chart', 'tbl', 'clipArt', 'dgm', 'media', 'pic')
                     else ph_type)
        matches = [candidate for candidate in pkg.xml(master_part).findall('p:cSld/p:spTree/p:sp', NS)
                   if placeholder(candidate) is not None
                   and placeholder(candidate).get('type', 'obj') == base_type]
        if len(matches) == 1:
            transforms.append(shape_transform(matches[0]))
    return transforms


def shape_geometry(pkg, part, shape):
    transforms = [xp for xp in inherited_transforms(pkg, part, shape) if xp is not None]
    geometry = {}
    for key, tag, attr in (('x', 'off', 'x'), ('y', 'off', 'y'),
                           ('width', 'ext', 'cx'), ('height', 'ext', 'cy')):
        geometry[key] = next((int(node.get(attr)) for xp in transforms
                              for node in [xp.find(q('a', tag))]
                              if node is not None and node.get(attr) is not None), None)
    geometry['rotation'] = next((int(xp.get('rot')) / 60000 for xp in transforms
                                 if xp.get('rot') is not None), 0)
    return geometry


def build_model(pkg, path):
    targets = {}
    model = {'path': str(path), 'hash': digest(pkg.raw)}
    if path.suffix.lower() == '.docx':
        part = 'word/document.xml'
        root = pkg.xml(part)
        body = root.find('w:body', NS)
        if body is None:
            fail('invalid-package', 'Missing Word document body')
        model.update(kind='docx', paragraphs=paragraphs(body.iter(q('w', 'p')), 'w', '', part, targets))
    else:
        root = pkg.xml('ppt/presentation.xml')
        size = root.find('p:sldSz', NS)
        model.update(kind='pptx', width=int(size.get('cx')) if size is not None else 9144000,
                     height=int(size.get('cy')) if size is not None else 6858000, slides=[])
        rels = pkg.relationships('ppt/presentation.xml')
        for s, slide_id in enumerate(root.findall('p:sldIdLst/p:sldId', NS)):
            part = rels.get(slide_id.get(q('r', 'id')))
            if not part:
                fail('invalid-package', 'Slide relationship is missing')
            slide = pkg.xml(part)
            sr = pkg.relationships(part)
            bg = slide.find('p:cSld/p:bg/p:bgPr', NS)
            row = {'id': f'slide:{s}', 'index': s, 'part': part, 'background': solid_fill(bg), 'elements': []}
            tree = slide.find('p:cSld/p:spTree', NS)
            if tree is None:
                fail('invalid-package', 'Missing slide shape tree')
            # Group descendants use local coordinate systems; expose only the group.
            for shape in tree:
                tag = E.QName(shape).localname
                if tag not in ('sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'):
                    continue
                nv = shape.find('.//p:cNvPr', NS)
                sid = nv.get('id') if nv is not None else None
                if sid is None:
                    continue
                ident = f'shape:{s}:{sid}'
                if ident in targets:
                    fail('invalid-package', 'Duplicate slide shape ID')
                xp = shape_transform(shape)
                geometry = shape_geometry(pkg, part, shape)
                tx = shape.find('p:txBody', NS)
                kind = 'image' if tag == 'pic' else 'text' if tx is not None else 'shape'
                if tag == 'grpSp':
                    kind = 'unsupported'
                if tag == 'graphicFrame':
                    kind = 'table' if shape.find('.//a:tbl', NS) is not None else 'chart' if any(E.QName(x).localname == 'chart' for x in shape.iter()) else 'unsupported'
                renderable = (all(geometry[key] is not None for key in ('x', 'y', 'width', 'height'))
                              and geometry['width'] > 0 and geometry['height'] > 0)
                editable = renderable and kind not in ('unsupported', 'table', 'chart')
                item = {'id': ident, 'shapeId': sid, 'kind': kind, **geometry,
                        'editable': editable, 'renderable': renderable,
                        'movable': editable and xp is not None
                                   and xp.find('a:off', NS) is not None
                                   and xp.find('a:ext', NS) is not None,
                        'fill': solid_fill(shape.find('p:spPr', NS)), 'paragraphs': []}
                if not renderable:
                    item['reason'] = 'Shape geometry is missing, unresolved, or has nonpositive size'
                elif not editable:
                    item['reason'] = 'Grouped shapes are unsupported' if tag == 'grpSp' else 'This shape type is unsupported'
                if editable:
                    targets[ident] = (part, shape, 'shape', 'a')
                if tx is not None:
                    item['paragraphs'] = paragraphs(tx.findall('a:p', NS), 'a', f'{s}:{sid}:', part, targets if editable else {})
                item['text'] = '\n'.join(p['text'] for p in item['paragraphs'])
                if tag == 'pic':
                    blip = shape.find('.//a:blip', NS)
                    image = sr.get(blip.get(q('r', 'embed'))) if blip is not None else None
                    if image and image in pkg.names:
                        image_data = pkg.read(image)
                        mime = mimetypes.guess_type(image)[0]
                        if len(image_data) <= 8 * 1024 * 1024 and mime in ('image/png', 'image/jpeg', 'image/gif', 'image/webp'):
                            item['image'] = 'data:' + mime + ';base64,' + base64.b64encode(image_data).decode('ascii')
                row['elements'].append(item)
            model['slides'].append(row)
    return model, targets


def set_run_text(run, prefix, text):
    # Keep properties, hyperlinks' parents, drawings and unsupported extension XML.
    for item in list(run):
        if E.QName(item).localname in ('t', 'tab', 'br', 'cr'):
            run.remove(item)
    if prefix == 'a':
        node = E.SubElement(run, q('a', 't'))
        node.set(XML_SPACE, 'preserve')
        node.text = text
        return
    for index, line in enumerate(text.split('\n')):
        if index:
            E.SubElement(run, q(prefix, 'br'))
        chunks = line.split('\t')
        for n, chunk in enumerate(chunks):
            if n:
                E.SubElement(run, q(prefix, 'tab')) if prefix == 'w' else E.SubElement(run, q(prefix, 't')).__setattr__('text', '\t')
            if chunk or len(chunks) == 1:
                node = E.SubElement(run, q(prefix, 't'))
                node.set(XML_SPACE, 'preserve')
                node.text = chunk


def set_text(node, kind, prefix, text):
    if not isinstance(text, str) or len(text) > 1000000 or any(ord(c) < 32 and c not in '\n\r\t' for c in text):
        fail('invalid-operation', 'Text must be valid XML text up to one million characters')
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    if kind == 'run':
        set_run_text(node, prefix, text)
        return
    if kind == 'shape':
        tx = node.find('p:txBody', NS)
        if tx is None:
            fail('unsupported-operation', 'This shape has no editable text body')
        ps = tx.findall('a:p', NS)
        if len(ps) > 1:
            fail('unsupported-operation', 'Whole-shape text replacement requires a single paragraph; edit individual paragraphs or runs')
    else:
        ps = [node]
    if not ps:
        fail('unsupported-operation', 'No editable paragraph')
    if any(E.QName(x).localname in ('fldChar', 'instrText', 'fld') for p in ps for x in p.iter()):
        fail('unsupported-operation', 'Paragraphs containing fields require run-level edits')
    runs = [r for p in ps for r in p.iter(q(prefix, 'r')) if next((a for a in r.iterancestors() if a.tag == q(prefix, 'p')), None) is p]
    if not runs:
        run = E.SubElement(ps[0], q(prefix, 'r'))
        runs = [run]
    set_run_text(runs[0], prefix, text)
    for run in runs[1:]:
        set_run_text(run, prefix, '')


def format_node(node, kind, prefix, op):
    runs = [node] if kind == 'run' else list(node.iter(q(prefix, 'r')))
    for key in ('bold', 'italic', 'underline'):
        if key in op and not isinstance(op[key], bool):
            fail('invalid-operation', key + ' must be boolean')
    if 'size' in op and (isinstance(op['size'], bool) or not isinstance(op['size'], (int, float)) or not math.isfinite(op['size']) or not 1 <= op['size'] <= 400):
        fail('invalid-operation', 'Font size must be between 1 and 400 points')
    if 'color' in op and (not isinstance(op['color'], str) or not re.fullmatch(r'#?[0-9a-fA-F]{6}', op['color'])):
        fail('invalid-operation', 'Color must have six hexadecimal digits')
    for run in runs:
        pr = child(run, prefix, 'rPr')
        for key, attr in (('bold', 'b'), ('italic', 'i'), ('underline', 'u')):
            if key not in op:
                continue
            val = ('single' if op[key] else 'none') if key == 'underline' else ('1' if op[key] else '0')
            if prefix == 'w':
                child(pr, 'w', attr).set(q('w', 'val'), val)
            else:
                pr.set(attr, ('sng' if op[key] else 'none') if key == 'underline' else val)
        if 'size' in op:
            val = str(round(op['size'] * (2 if prefix == 'w' else 100)))
            child(pr, 'w', 'sz').set(q('w', 'val'), val) if prefix == 'w' else pr.set('sz', val)
        if 'color' in op:
            val = op['color'].lstrip('#').upper()
            if prefix == 'w':
                child(pr, 'w', 'color').set(q('w', 'val'), val)
            else:
                for old in list(pr):
                    if E.QName(old).localname in ('solidFill', 'noFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'):
                        pr.remove(old)
                E.SubElement(child(pr, 'a', 'solidFill'), q('a', 'srgbClr'), val=val)
    if 'align' in op:
        mapping = {'left': 'l', 'center': 'ctr', 'right': 'r', 'justify': 'just'}
        if not isinstance(op['align'], str) or op['align'] not in mapping:
            fail('invalid-operation', 'Unsupported paragraph alignment')
        ps = [node] if kind == 'paragraph' else [next((a for a in node.iterancestors() if a.tag == q(prefix, 'p')), None)] if kind == 'run' else list(node.iter(q(prefix, 'p')))
        for para in ps:
            if para is not None:
                pp = child(para, prefix, 'pPr')
                child(pp, 'w', 'jc').set(q('w', 'val'), op['align']) if prefix == 'w' else pp.set('algn', mapping[op['align']])


def apply_operations(pkg, targets, operations):
    if not isinstance(operations, list) or len(operations) > 5000:
        fail('invalid-operation', 'operations must be an array of at most 5000 edits')
    for op in operations:
        if not isinstance(op, dict) or not isinstance(op.get('id'), str) or op['id'] not in targets:
            fail('invalid-operation', 'Unknown edit target')
        part, node, kind, prefix = targets[op['id']]
        if op.get('type') == 'text':
            set_text(node, kind, prefix, op.get('text'))
        elif op.get('type') == 'format':
            format_node(node, kind, prefix, op)
        elif op.get('type') == 'move' and kind == 'shape':
            if E.QName(node).localname == 'grpSp':
                fail('unsupported-operation', 'Group transforms are unsupported')
            xp = node.find('p:spPr/a:xfrm', NS)
            if xp is None:
                xp = node.find('p:xfrm', NS)
            if xp is None:
                fail('unsupported-operation', 'Inherited shape positions cannot be moved')
            for key, tag, attr in (('x', 'off', 'x'), ('y', 'off', 'y'), ('width', 'ext', 'cx'), ('height', 'ext', 'cy')):
                if key in op:
                    value = op[key]
                    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or abs(value) > 2147483647 or (key in ('width', 'height') and value <= 0):
                        fail('invalid-operation', 'Invalid shape dimension')
                    child(xp, 'a', tag).set(attr, str(round(value)))
        else:
            fail('unsupported-operation', 'Unsupported operation or target')
        pkg.changed.add(part)


def handle(request):
    if not isinstance(request, dict):
        fail('invalid-request', 'Request must be an object')
    root_value = request.get('workspaceRoot')
    if not isinstance(root_value, str) or not Path(root_value).is_absolute():
        fail('invalid-path', 'workspaceRoot must be an absolute directory')
    root = Path(root_value).resolve(strict=True)
    if not root.is_dir():
        fail('invalid-path', 'workspaceRoot is not a directory')
    path = scoped_path(request.get('path'), root)
    if path.stat().st_size > MAX_FILE:
        fail('file-limit', 'File exceeds 64 MiB')
    with path.open('rb') as source:
        data = source.read(MAX_FILE + 1)
    if len(data) > MAX_FILE:
        fail('file-limit', 'File exceeds 64 MiB')
    pkg = Package(data)
    try:
        model, targets = build_model(pkg, path)
        if request.get('action') == 'read':
            return model
        if request.get('action') != 'save':
            fail('invalid-request', 'action must be read or save')
        if request.get('expectedHash') != model['hash']:
            fail('conflict', 'Source changed; reload before saving')
        apply_operations(pkg, targets, request.get('operations'))
        explicit = request.get('outputPath') is not None
        output = scoped_path(request['outputPath'], root, False) if explicit else path.with_name(path.stem + '-edited' + path.suffix)
        if output.suffix.lower() != path.suffix.lower():
            fail('unsupported-format', 'Output extension must match source')
        if not output.parent.is_dir():
            fail('invalid-path', 'Output directory does not exist')
        if explicit and output.exists() and request.get('expectedOutputHash') != file_digest(output):
            fail('output-conflict', 'Replacing an output requires its current expectedOutputHash')
        fd, temporary = tempfile.mkstemp(prefix='.office-canvas-', suffix=path.suffix, dir=output.parent)
        os.close(fd)
        try:
            pkg.save(temporary)
            if file_digest(path) != model['hash']:
                fail('conflict', 'Source changed during save')
            if explicit and output.exists():
                if request.get('expectedOutputHash') != file_digest(output):
                    fail('output-conflict', 'Output changed during save')
                os.replace(temporary, output)
            else:
                number = 1
                while True:
                    try:
                        # Hard-link publication is atomic and refuses to replace an existing file.
                        os.link(temporary, output)
                        break
                    except FileExistsError:
                        if explicit:
                            fail('output-conflict', 'Output already exists')
                        number += 1
                        output = path.with_name(f'{path.stem}-edited-{number}{path.suffix}')
            return {'path': str(output), 'hash': file_digest(output), 'kind': model['kind'], 'operationsApplied': len(request['operations'])}
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)
    finally:
        pkg.zip.close()


def main():
    try:
        raw = sys.stdin.buffer.read(MAX_REQUEST + 1)
        if len(raw) > MAX_REQUEST:
            fail('request-limit', 'JSON request exceeds 4 MiB')
        request = json.loads(raw.decode('utf-8-sig'))
        response = {'ok': True, 'value': handle(request)}
    except BackendError as exc:
        response = {'ok': False, 'error': {'code': exc.code, 'message': str(exc)}}
    except (ValueError, TypeError, KeyError, OSError, OverflowError) as exc:
        response = {'ok': False, 'error': {'code': 'invalid-request', 'message': str(exc)}}
    sys.stdout.write(json.dumps(response, ensure_ascii=True, separators=(',', ':')))


if __name__ == '__main__':
    main()
