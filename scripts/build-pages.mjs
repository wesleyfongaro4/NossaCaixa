/*
 * build-pages.mjs — gera docs/index.html (versão estática p/ GitHub Pages)
 * a partir de Index.html (template do GAS).
 *
 * Uso:
 *   node scripts/build-pages.mjs              → gera com NC_API_URL = '' (o app pede a URL uma vez, salva no localStorage)
 *   node scripts/build-pages.mjs <url /exec>  → gera com a URL fixada (cuidado: fica visível no repo público)
 */
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const src = readFileSync(root + 'Index.html', 'utf8');
const dst = root + 'docs/index.html';

const SCRIPTLET = 'let NC_API_URL = ("<?= ncApiUrl_() ?>" || \'\').trim();';
const baked = process.argv[2] || '';
if (baked && !/\/exec$/.test(baked)) {
  console.error('A URL deve terminar em /exec. Nada foi alterado.');
  process.exit(1);
}

let out = src.replace(SCRIPTLET, 'let NC_API_URL = (' + JSON.stringify(baked) + ').trim();');
const fixado = out !== src;
if (!fixado) {
  console.error('Scriptlet do GAS não encontrado — nada foi alterado.');
  process.exit(1);
}

// Guarda de segurança: mesmo com URL fixada, mantém comportamento de template fora do GAS
mkdirSync(root + 'docs', { recursive: true });
writeFileSync(dst, out, 'utf8');
console.log('docs/index.html gerado' + (baked ? ' com URL fixada: ' + baked : ' sem URL fixada (o app pede no primeiro acesso)'));
