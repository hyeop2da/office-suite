// 두 HWPX/OOXML 패키지를 비교한다. 사용: node tools/compare.ts <원본> <수정본> [--json]
import { readFileSync } from 'node:fs';
import { diffPackages } from '../src/xmldiff.ts';

const [a, b, flag] = process.argv.slice(2);
if (!a || !b) { console.error('사용: node tools/compare.ts <원본> <수정본> [--json]'); process.exit(2); }
const d = diffPackages(new Uint8Array(readFileSync(a)), new Uint8Array(readFileSync(b)));
if (flag === '--json') { console.log(JSON.stringify(d, null, 2)); process.exit(0); }

const z = d.zip;
console.log(`ZIP  항목 ${z.names.a.length} → ${z.names.b.length}, 순서 변경: ${z.orderChanged ? '예' : '아니오'}`);
console.log(`     바이트(압축 포함) 동일 ${z.rawIdentical.length}, 내용 동일 ${z.identical.length}, 내용 변경 ${z.changed.length}, 추가 ${z.added.length}, 삭제 ${z.removed.length}`);
if (z.changed.length) console.log(`     변경: ${z.changed.join(', ')}`);
if (z.added.length) console.log(`     추가: ${z.added.join(', ')}`);
if (z.methodChanged.length) console.log(`     압축 방식만 변경: ${z.methodChanged.join(', ')}`);
for (const [name, x] of Object.entries(d.xml)) {
  console.log(`\n[${name}]`);
  const show = (title: string, o: Record<string, unknown>) => { const k = Object.keys(o); if (k.length) console.log(`  ${title}: ${k.map(key => `${key}=${JSON.stringify(o[key])}`).join(', ')}`); };
  show('요소 수 변화', x.tagCount);
  show('속성 손실', x.attrLost);
  show('속성 추가', x.attrAdded);
  show('속성 값 변경', x.attrChanged);
  show('빈 요소 수', x.emptyCount);
}
