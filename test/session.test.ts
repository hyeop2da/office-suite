import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { EditSession } from '../src/session.ts';
import { UnsupportedFormatError } from '../src/engine.ts';

const fx = (p: string) => new Uint8Array(readFileSync(new URL(`../fixtures/${p}`, import.meta.url)));

for (const [file, name] of [['hwpx/pps-manual-14-7.hwpx', 'a.hwpx'], ['hwpx/pipc-privacy-20231127.hwpx', 'b.hwpx'], ['docx/table.docx', 'c.docx'], ['docx/basic.docx', 'd.docx']] as const) {
  test(`${name}: 수정 목록 → 저장 → 재열기 검증 통과`, () => {
    const s = new EditSession(fx(file), name);
    const targets = s.paragraphs.filter(p => !p.partlyLocked && p.text.length >= 2).slice(0, 3);
    assert.ok(targets.length >= 1);
    targets.forEach((p, i) => s.setText(p.id, `${p.text}[${i}]`));
    assert.equal(s.changeCount, targets.length);
    const r = s.save();
    assert.deepEqual(r.problems, []);
    assert.ok(r.verified);
    const again = new EditSession(r.bytes, name);
    for (const [i, p] of targets.entries()) assert.equal(again.paragraphs.find(x => x.id === p.id)!.text, `${p.text}[${i}]`);
    // 수정하지 않은 문단은 그대로
    const touched = new Set(targets.map(t => t.id));
    for (const p of s.paragraphs) if (!touched.has(p.id)) assert.equal(again.paragraphs.find(x => x.id === p.id)!.text, p.text);
    // 원본 바이트는 세션이 바꾸지 않는다
    assert.deepEqual(s.original, fx(file));
  });
}

test('원래 글자로 돌리면 수정 목록에서 빠지고 되돌리기가 동작한다', () => {
  const s = new EditSession(fx('docx/basic.docx'), 'x.docx');
  const p = s.paragraphs[0];
  s.setText(p.id, 'x');
  assert.equal(s.changeCount, 1);
  s.setText(p.id, p.text);
  assert.equal(s.changeCount, 0);
  s.setText(p.id, 'y'); s.revert(p.id);
  assert.equal(s.changeCount, 0);
  s.setText(p.id, 'z'); s.revertAll();
  assert.equal(s.changeCount, 0);
  const noChange = s.save();
  assert.equal(Buffer.compare(Buffer.from(noChange.bytes), Buffer.from(fx('docx/basic.docx'))), 0, '수정이 없으면 원본과 같다');
});

test('편집 불가 글자가 들어가면 저장이 거부된다', () => {
  const s = new EditSession(fx('docx/basic.docx'), 'x.docx');
  s.setText(s.paragraphs[0].id, '줄\n바꿈');
  assert.throws(() => s.save(), /줄바꿈/);
});

test('지원하지 않는 형식은 열기에서 UnsupportedFormatError', () => {
  assert.throws(() => new EditSession(new TextEncoder().encode('%PDF-1.7 ...'), 'a.pdf'), UnsupportedFormatError);
});
