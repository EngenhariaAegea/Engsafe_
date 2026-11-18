import { access, readFile } from 'node:fs/promises';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const frontend = resolve(root, 'frontend');
const html = await readFile(resolve(frontend, 'index.html'), 'utf8');
const required = [
  './assets/styles.css',
  './config.js',
  './assets/pdf-template.js',
  './assets/app.js'
];

for (const reference of required) {
  if (!html.includes(reference)) throw new Error(`Referência ausente no index.html: ${reference}`);
  await access(resolve(frontend, reference));
}

const scripts = [
  ['config.js', await readFile(resolve(frontend, 'config.js'), 'utf8')],
  ['assets/pdf-template.js', await readFile(resolve(frontend, 'assets/pdf-template.js'), 'utf8')],
  ['assets/app.js', await readFile(resolve(frontend, 'assets/app.js'), 'utf8')]
];

for (const [name, source] of scripts) {
  try {
    new Function(source);
  } catch (error) {
    throw new Error(`${name}: ${error.message}`);
  }
}

const order = required.slice(1).map((reference) => html.indexOf(reference));
if (order.some((position, index) => index > 0 && position <= order[index - 1])) {
  throw new Error('Ordem dos scripts inválida no index.html.');
}

if (/<style(?:\s|>)/i.test(html) || /<script>\s*\/\/ ── CONFIG/i.test(html)) {
  throw new Error('O index.html ainda contém estilos ou lógica principal inline.');
}

if (/data:image\//i.test(html)) throw new Error('O index.html ainda contém imagens embutidas.');

console.log('Estrutura modular e JavaScript validados.');
