// 웹앱 빌드: src/app/main.ts → dist/app.js, web/* 복사, 샘플 문서 복사.
import { build } from 'esbuild';
import { mkdirSync, copyFileSync, rmSync, cpSync, readdirSync, readFileSync, writeFileSync, statSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = new URL('../', import.meta.url);
const p = (rel) => fileURLToPath(new URL(rel, root));
rmSync(p('dist'), { recursive: true, force: true });
mkdirSync(p('dist/samples'), { recursive: true });
await build({
  entryPoints: [p('src/app/main.ts')],
  outdir: p('dist'), entryNames: 'app',
  bundle: true, minify: true, format: 'esm', splitting: true, sourcemap: false,
  target: ['es2020', 'chrome100', 'safari15'],
  logLevel: 'info',
});
copyFileSync(p('node_modules/@rhwp/core/rhwp_bg.wasm'), p('dist/rhwp_bg.wasm'));
copyFileSync(p('node_modules/pdfjs-dist/build/pdf.worker.min.mjs'), p('dist/pdf.worker.min.mjs'));
cpSync(p('node_modules/pdfjs-dist/cmaps'), p('dist/cmaps'), { recursive: true });
cpSync(p('node_modules/pdfjs-dist/standard_fonts'), p('dist/standard_fonts'), { recursive: true });
for (const f of ['index.html', 'app.css']) copyFileSync(p(`web/${f}`), p(`dist/${f}`));
for (const f of ['hwpx/pps-manual-14-7.hwpx', 'hwpx/pipc-privacy-20231127.hwpx', 'docx/basic.docx', 'docx/table.docx', 'xlsx/types.xlsx', 'pptx/basic.pptx', 'pdf/classic.pdf']) {
  copyFileSync(p(`fixtures/${f}`), p(`dist/samples/${f.split('/').pop()}`));
}
// 오프라인 지원: 앱 파일 목록과 버전(내용 해시)을 서비스 워커에 넣는다. 큰 보기 엔진은 처음 쓸 때 캐시에 둔다.
for (const f of ['manifest.webmanifest', 'icon-180.png', 'icon-192.png', 'icon-512.png']) copyFileSync(p(`web/${f}`), p(`dist/${f}`));
const walk = (d) => readdirSync(d).flatMap(n => { const f = join(d, n); return statSync(f).isDirectory() ? walk(f) : [f]; });
const dist = p('dist');
const shell = walk(dist).map(f => relative(dist, f).split(sep).join('/')).filter(f => !/^(cmaps|standard_fonts)\//.test(f) && !/\.(wasm)$/.test(f) && !/^pdf\.worker/.test(f) && f !== 'sw.js').sort();
const hash = createHash('sha256');
for (const f of shell) hash.update(f).update(readFileSync(join(dist, f)));
const sw = readFileSync(p('web/sw.js'), 'utf8').replace('__VERSION__', `office-${hash.digest('hex').slice(0, 12)}`).replace('__SHELL__', JSON.stringify(shell));
writeFileSync(join(dist, 'sw.js'), sw);
console.log(`서비스 워커: 앱 파일 ${shell.length}개 미리 저장`);
console.log('빌드 완료: dist/');
