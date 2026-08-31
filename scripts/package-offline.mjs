import { copyFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const output = resolve(projectRoot, 'offline-dist');

await copyFile(resolve(projectRoot, 'launcher/bin/WenjianLauncher.exe'), resolve(output, '启动文鉴.exe'));
await copyFile(resolve(projectRoot, 'launcher/WenjianLauncher.exe.config'), resolve(output, '启动文鉴.exe.config'));
await copyFile(resolve(projectRoot, 'offline/离线使用说明.txt'), resolve(output, '离线使用说明.txt'));
await copyFile(resolve(projectRoot, 'LICENSE'), resolve(output, 'LICENSE'));
await copyFile(resolve(projectRoot, 'THIRD_PARTY_NOTICES.md'), resolve(output, 'THIRD_PARTY_NOTICES.md'));
console.log('Offline package is ready in offline-dist.');
