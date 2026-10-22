// HWPX 시료별 기능 목록. 사용: node tools/inventory.ts [--md]
import { readFileSync, readdirSync } from 'node:fs';
import { inventory } from '../src/inventory.ts';

const groups: [string, string][] = [['한/글 작성 공개 문서', 'fixtures/hwpx'], ['rhwp가 HWP에서 변환', 'fixtures/hwpx-generated']];
const rows: { group: string; id: string; inv: ReturnType<typeof inventory> }[] = [];
for (const [group, dir] of groups) {
  for (const f of readdirSync(new URL(`../${dir}/`, import.meta.url)).filter(x => x.endsWith('.hwpx'))) {
    rows.push({ group, id: f.replace(/\.hwpx$/, ''), inv: inventory(new Uint8Array(readFileSync(new URL(`../${dir}/${f}`, import.meta.url)))) });
  }
}
if (process.argv.includes('--md')) {
  const keys = Object.keys(rows[0].inv.counts);
  console.log(`| 기능 | ${rows.map(r => r.id).join(' | ')} |`);
  console.log(`|---|${rows.map(() => '---:').join('|')}|`);
  console.log(`| 구역 수 | ${rows.map(r => r.inv.sections).join(' | ')} |`);
  console.log(`| BinData 파일 | ${rows.map(r => r.inv.binData).join(' | ')} |`);
  for (const k of keys) console.log(`| ${k} | ${rows.map(r => r.inv.counts[k]).join(' | ')} |`);
  console.log(`| 필드 종류 | ${rows.map(r => Object.entries(r.inv.fieldTypes).map(([t, n]) => `${t}×${n}`).join(' ') || '-').join(' | ')} |`);
} else console.log(JSON.stringify(rows, null, 2));
