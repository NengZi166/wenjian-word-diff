import { cp, mkdir } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const source = resolve(projectRoot, 'node_modules/docxodus/dist/wasm');
const destination = resolve(projectRoot, 'public/vendor/docxodus/wasm');

await mkdir(destination, { recursive: true });
await cp(source, destination, { recursive: true, force: true });
console.log('Docxodus WASM assets are ready.');
