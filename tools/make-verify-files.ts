// 한/글 직접 열기 검증용 복사본을 verify-out/에 만든다(원본 fixtures는 건드리지 않는다).
// 사용: node tools/make-verify-files.ts
import { readFileSync, writeFileSync, mkdirSync, readdirSync, existsSync } from 'node:fs';
import { HwpxDocument } from '../src/hwpx.ts';

const out = new URL('../verify-out/', import.meta.url);
mkdirSync(out, { recursive: true });
const ids = readdirSync(new URL('../fixtures/hwpx/', import.meta.url)).filter(f => f.endsWith('.hwpx')).map(f => f.replace('.hwpx', ''));
const load = (id: string) => new Uint8Array(readFileSync(new URL(`../fixtures/hwpx/${id}.hwpx`, import.meta.url)));
// 한/글이 열어 둔 파일은 잠겨 있으므로, 내용이 같으면 다시 쓰지 않는다.
const put = (name: string, b: Uint8Array) => {
  const u = new URL(name, out);
  if (existsSync(u) && Buffer.compare(readFileSync(u), Buffer.from(b)) === 0) return;
  writeFileSync(u, b);
};
const MARK = 'EDIT2026';

const manifest: string[] = [];
for (const id of ids) {
  const orig = load(id);
  put(`E1-${id}-original.hwpx`, orig);
  const same = HwpxDocument.open(orig).save();
  put(`E1-${id}-ours-unmodified.hwpx`, same);
  const d = HwpxDocument.open(orig);
  d.prependText(MARK, { layout: 'keep' });
  put(`E2-${id}-marker.hwpx`, d.save()); // 이전 기본 정책(keep)
  const d2 = HwpxDocument.open(orig);
  d2.prependText(MARK); // 현재 기본 정책(drop)
  put(`E5-${id}-marker-drop.hwpx`, d2.save());
  manifest.push(`${id}: unmodified save identical=${Buffer.compare(Buffer.from(same), Buffer.from(orig)) === 0}`);
}

// E3: 줄이 늘어나는 긴 글자 삽입. 같은 편집을 keep/drop으로 각각.
const LONG = ' 줄배치시험 가나다라마바사 아자차카타파하 가나다라마바사 아자차카타파하 가나다라마바사 아자차카타파하 가나다라마바사 아자차카타파하 끝';
for (const layout of ['keep', 'drop'] as const) {
  const d = HwpxDocument.open(load('pps-guide-14-1'));
  const p = d.paragraphs().find(x => x.context === 'body' && [...x.text].length >= 20)!;
  d.setParagraphText(p.section, p.index, p.text + LONG, { layout });
  put(`E3-pps-guide-long-${layout}.hwpx`, d.save());
  manifest.push(`E3 ${layout}: 편집 문단 "${p.text.slice(0, 30)}…" (${[...p.text].length}자) + ${[...LONG].length}자`);
}
// E4: 좁은 표 셀에 긴 글자(줄 수가 늘어남), 기본 정책(drop)
{
  const d = HwpxDocument.open(load('pps-spec-14-13'));
  const c = d.paragraphs().find(x => x.context === 'cell' && [...x.text].length >= 4)!;
  d.setParagraphText(c.section, c.index, c.text + LONG);
  put('E4-pps-spec-cell-long-drop.hwpx', d.save());
  manifest.push(`E4 cell: "${c.text.slice(0, 20)}" + ${[...LONG].length}자`);
}
writeFileSync(new URL('MANIFEST.txt', out), manifest.join('\n') + '\n');
console.log(manifest.join('\n'));
