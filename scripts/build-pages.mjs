import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const source = resolve(root, 'frontend');
const output = resolve(root, 'dist');

// Variáveis: primeiro o ambiente (GitHub Actions / Cloudflare), depois o .env local (ignorado pelo git).
const envFile = resolve(root, '.env');
if (existsSync(envFile)) {
  for (const line of readFileSync(envFile, 'utf8').split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*?)\s*$/);
    if (m && process.env[m[1]] === undefined) process.env[m[1]] = m[2].replace(/^(['"])(.*)\1$/, '$2');
  }
}

const url = String(process.env.SUPABASE_URL || '').replace(/\/+$/, '');
const key = String(process.env.SUPABASE_PUBLISHABLE_KEY || '');

if (!/^https:\/\/[a-z0-9]+\.supabase\.co$/.test(url)) {
  throw new Error('SUPABASE_URL ausente ou inválida (esperado https://<ref>.supabase.co).');
}
if (/^sb_secret_/.test(key) || /service_role/.test(key)) {
  throw new Error('SUPABASE_PUBLISHABLE_KEY contém uma chave SECRETA. Nunca use sb_secret_/service_role no front-end. Rotacione essa chave.');
}
if (!/^sb_publishable_[A-Za-z0-9_-]+$/.test(key)) {
  throw new Error('SUPABASE_PUBLISHABLE_KEY ausente ou inválida (esperado sb_publishable_...).');
}

await rm(output, { recursive: true, force: true });
await mkdir(output, { recursive: true });
await cp(source, output, { recursive: true, force: true });

const fill = async (rel, pairs) => {
  const file = resolve(output, rel);
  let text = await readFile(file, 'utf8');
  for (const [marker, value] of pairs) text = text.split(marker).join(value);
  if (/__SUPABASE_[A-Z_]+__/.test(text)) throw new Error(`Marcador não substituído em ${rel}.`);
  await writeFile(file, text);
};
await fill('config.js', [['__SUPABASE_URL__', url], ['__SUPABASE_PUBLISHABLE_KEY__', key]]);
await fill('_headers', [['__SUPABASE_URL__', url], ['__SUPABASE_WSS__', url.replace(/^https:/, 'wss:')]]);

const html = await readFile(resolve(output, 'index.html'), 'utf8');
const app = await readFile(resolve(output, 'assets/app.js'), 'utf8');
const productionSource = `${html}\n${app}`;
if (!app.includes('OUTLOOK_COMPOSE_URL')) throw new Error('Fluxo do Outlook não encontrado em assets/app.js.');
if (/RESEND_API_KEY|EMAIL_WORKER_URL|api\.resend\.com|\/api\/notificacoes/.test(productionSource)) {
  throw new Error('O frontend ainda contém configuração de envio automático.');
}
if (/sb_(publishable|secret)_[A-Za-z0-9]|https:\/\/[a-z0-9]+\.supabase\.co/.test(app)) {
  throw new Error('assets/app.js não deve conter URL ou chave do Supabase fixas; use config.js.');
}
console.log(`Cloudflare Pages preparado. Supabase: ${url}`);
