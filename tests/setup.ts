import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';

const source = readFileSync(fileURLToPath(new URL('../public/vendor/xlsx.full.min.js', import.meta.url)), 'utf8');
const sandbox: Record<string, unknown> = { console, Uint8Array, ArrayBuffer, Date, Buffer, setTimeout, clearTimeout };
vm.createContext(sandbox);
vm.runInContext(source, sandbox, { timeout: 10_000 });
Object.assign(globalThis, { XLSX: sandbox.XLSX });
