import {readFile, realpath, mkdtemp, rm, stat} from 'node:fs/promises';
import {resolve, relative, isAbsolute, join, extname} from 'node:path';
import {pathToFileURL} from 'node:url';
import {createHash} from 'node:crypto';

let temporary;
let converter;
try {
  const chunks = [];
  let length = 0;
  for await (const chunk of process.stdin) { length += chunk.length; if (length > 1048576) throw Error('Request too large'); chunks.push(chunk); }
  const request = JSON.parse(Buffer.concat(chunks).toString('utf8'));
  const root = await realpath(request.workspaceRoot);
  const path = await realpath(resolve(root, request.path));
  const rel = relative(root, path);
  if (rel.startsWith('..') || isAbsolute(rel) || !['.docx', '.pptx'].includes(extname(path).toLowerCase())) throw Error('Only workspace DOCX/PPTX files are supported');
  const info = await stat(path);
  if (!info.isFile() || info.size > 64 * 1024 * 1024) throw Error('File exceeds 64 MiB');
  const pages = request.pages ?? [1];
  if (!Array.isArray(pages) || !pages.length || pages.length > 5 || pages.some(p => !Number.isSafeInteger(p) || p < 1)) throw Error('Invalid page selection');
  const before = createHash('sha256').update(await readFile(path)).digest('hex');
  if (request.expectedHash && request.expectedHash !== before) throw Error('File changed. Reload before previewing.');
  const {createConverter} = await import(pathToFileURL(request.libreofficeModulePath).href);
  temporary = await mkdtemp(join(root, '.office-canvas-'));
  converter = await createConverter({timeoutMs: 90000, maxOutputBytes: 20 * 1024 * 1024});
  const manifest = await converter.renderImages({inputPath: path, outputDir: join(temporary, 'pages'), pages, dpi: 110, maxPages: 5, maxPixels: 16000000, maxDimension: 1800});
  const images = [];
  for (const image of manifest.images) images.push({index: image.page ?? image.index, width: image.width, height: image.height, data: 'data:image/png;base64,' + (await readFile(image.path)).toString('base64')});
  const after = createHash('sha256').update(await readFile(path)).digest('hex');
  if (before !== after) throw Error('File changed during preview. Reload.');
  process.stdout.write(JSON.stringify({ok: true, value: {pages: images, pageCount: manifest.pageCount, hash: before, missingFonts: manifest.missingFonts}}));
} catch (error) {
  process.stdout.write(JSON.stringify({ok: false, error: {code: 'preview-failed', message: error.message}}));
} finally {
  await converter?.dispose();
  if (temporary) await rm(temporary, {recursive: true, force: true});
}
