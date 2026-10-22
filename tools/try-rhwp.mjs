import { readFileSync } from 'node:fs';
import * as rhwp from '../node_modules/@rhwp/core/rhwp.js';
globalThis.measureTextWidth = (font, text) => [...text].length * 10;
rhwp.initSync({ module: readFileSync(new URL('../node_modules/@rhwp/core/rhwp_bg.wasm', import.meta.url)) });
console.log('version', rhwp.version());
const d = new rhwp.HwpDocument(new Uint8Array(readFileSync(new URL('../fixtures/hwpx/pps-manual-14-7.hwpx', import.meta.url))));
console.log('pages', d.pageCount());
const svg = d.renderPageSvg(0); console.log(svg.length, svg.slice(0,100));
