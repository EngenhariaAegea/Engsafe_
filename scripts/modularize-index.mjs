import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const frontend = resolve(root, 'frontend');
const assets = resolve(frontend, 'assets');
const images = resolve(assets, 'images');
const indexPath = resolve(frontend, 'index.html');

await mkdir(images, { recursive: true });
let html = await readFile(indexPath, 'utf8');

if (html.includes('./assets/app.js')) {
  console.log('O index.html já está modularizado.');
  process.exit(0);
}

const styleMatch = html.match(/<style>([\s\S]*?)<\/style>/);
if (!styleMatch) throw new Error('Bloco principal de estilos não encontrado.');
await writeFile(resolve(assets, 'styles.css'), `${styleMatch[1].trim()}\n`, 'utf8');
html = html.replace(styleMatch[0], '<link rel="stylesheet" href="./assets/styles.css">');

const scriptStart = html.search(/<script>\s*\/\/ ── CONFIG/);
const scriptEnd = html.lastIndexOf('</script>');
if (scriptStart < 0 || scriptEnd < scriptStart) {
  throw new Error('Bloco principal de JavaScript não encontrado.');
}
const scriptContentStart = html.indexOf('>', scriptStart) + 1;
let app = `${html.slice(scriptContentStart, scriptEnd).trim()}\n`;
const pdfMatch = app.match(/const PDF_TEMPLATE_DATA='(data:image\/png;base64,[^']+)';/);
if (!pdfMatch) throw new Error('Template do PDF não encontrado.');

await writeFile(
  resolve(assets, 'pdf-template.js'),
  `window.ENGSAFE_PDF_TEMPLATE = ${JSON.stringify(pdfMatch[1])};\n`,
  'utf8'
);
app = app.replace(pdfMatch[0], 'const PDF_TEMPLATE_DATA=window.ENGSAFE_PDF_TEMPLATE;');
await writeFile(resolve(assets, 'app.js'), app, 'utf8');

html = `${html.slice(0, scriptStart)}<script src="./assets/pdf-template.js"></script>\n<script src="./assets/app.js"></script>${html.slice(scriptEnd + '</script>'.length)}`;

let imageCount = 0;
const imageWrites = [];
html = html.replace(
  /<img([^>]*?)class="([^"]+)"([^>]*?)src="data:image\/(png|jpeg|webp);base64,([^"]+)"([^>]*)>/g,
  (tag, before, className, middle, extension, base64, after) => {
    imageCount += 1;
    const safeClass = className.trim().split(/\s+/)[0].replace(/[^a-z0-9-]/gi, '-');
    const ext = extension === 'jpeg' ? 'jpg' : extension;
    const filename = `${String(imageCount).padStart(2, '0')}-${safeClass}.${ext}`;
    imageWrites.push(writeFile(resolve(images, filename), Buffer.from(base64, 'base64')));
    return `<img${before}class="${className}"${middle}src="./assets/images/${filename}"${after}>`;
  }
);

await Promise.all(imageWrites);
await writeFile(indexPath, html, 'utf8');
console.log(`Index modularizado com ${imageCount} imagem(ns) externa(s).`);
