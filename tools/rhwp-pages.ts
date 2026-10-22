// rhwp가 센 쪽 수. 사용: node tools/rhwp-pages.ts <파일...>
import { readFileSync } from 'node:fs';
import * as rhwp from '@rhwp/core';
(globalThis as Record<string, unknown>).measureTextWidth = (_f: string, t: string) => [...t].length * 10;
rhwp.initSync({ module: readFileSync(new URL('../node_modules/@rhwp/core/rhwp_bg.wasm', import.meta.url)) });
for (const f of process.argv.slice(2)) console.log(`${f}\t${new rhwp.HwpDocument(new Uint8Array(readFileSync(f))).pageCount()}`);
