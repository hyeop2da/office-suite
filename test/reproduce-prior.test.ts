// 이전 분석(HWPX구조분석.md §3)의 수치를 새 비교 도구(src/xmldiff.ts)로 독립 재현한다.
// 이 시험은 "rhwp 0.8.7이 저장한 파일 ↔ 한/글 원본"의 파일 구조 차이만 다룬다. 한/글에서 열어 본 결과가 아니다.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { diffPackages } from '../src/xmldiff.ts';
import { HwpxDocument } from '../src/hwpx.ts';

const load = (p: string) => new Uint8Array(readFileSync(new URL(`../fixtures/${p}`, import.meta.url)));

// 이전 보고서 §3.2, §3.3, §3.4의 값
const PRIOR: Record<string, { pId: number; t: [number, number]; empty: [number, number]; order: boolean; extra?: (x: ReturnType<typeof diffPackages>) => void }> = {
  'nrich-application': { pId: 81, t: [56, 82], empty: [2, 28], order: true, extra: d => assert.equal(d.xml['Contents/header.xml'].attrLost['hh:paraPr@textDir'], 16) },
  'pipc-privacy-20231127': {
    pId: 58, t: [102, 110], empty: [11, 18], order: true,
    extra: d => {
      const s = d.xml['Contents/section0.xml'];
      assert.equal(s.attrLost['hp:subList@metatag'], 12);
      assert.equal(s.attrLost['hp:fieldBegin@zorder'], 1);
      assert.deepEqual(s.tagCount['hp:metaTag'], [1, 0]);
      assert.equal(d.xml['Contents/header.xml'].attrChanged['hh:strikeout@shape'], 2);
    },
  },
  'pps-guide-14-1': { pId: 100, t: [212, 238], empty: [1, 27], order: true },
  'pps-manual-14-7': { pId: 100, t: [65, 101], empty: [4, 40], order: false, extra: d => assert.deepEqual(d.zip.added, ['Preview/PrvImage.png']) },
  'pps-spec-14-13': { pId: 301, t: [165, 362], empty: [6, 203], order: false, extra: d => assert.deepEqual(d.zip.added, ['Preview/PrvImage.png']) },
};

for (const [id, exp] of Object.entries(PRIOR)) {
  test(`이전 보고서 재현(rhwp 저장본): ${id}`, () => {
    const d = diffPackages(load(`hwpx/${id}.hwpx`), load(`hwpx-rhwp-saved/${id}.hwpx`));
    const s = d.xml['Contents/section0.xml'];
    assert.equal(s.attrChanged['hp:p@id'], exp.pId, 'hp:p@id가 바뀐 문단 수');
    assert.deepEqual(s.tagCount['hp:t'], exp.t, 'hp:t 요소 수');
    assert.deepEqual(s.emptyCount['hp:t'], exp.empty, '빈 hp:t 수');
    // 이전 보고서는 "5쌍 모두 순서가 바뀌었다"고 썼으나, 공통 항목의 상대 순서가 바뀐 것은 3쌍이다.
    // 나머지 2쌍(manual, spec)은 PrvImage.png가 끼어들어 순서가 달라 보였을 뿐이다.
    assert.equal(d.zip.orderChanged, exp.order, 'ZIP 공통 항목 순서 변경 여부');
    assert.ok(d.zip.methodChanged.includes('version.xml'), 'version.xml 압축 방식 변경');
    assert.equal(d.zip.rawIdentical.length, 0, 'rhwp 저장본은 어떤 항목도 압축 바이트까지 같지 않다');
    exp.extra?.(d);
  });
}

test('같은 비교를 우리 편집 코어 출력에 적용: 구조 손실이 없다(글자 한 군데 편집)', () => {
  for (const id of Object.keys(PRIOR)) {
    const orig = load(`hwpx/${id}.hwpx`);
    const doc = HwpxDocument.open(orig);
    assert.ok(doc.prependText('〔편집〕', { layout: 'keep' }));
    const d = diffPackages(orig, doc.save());
    assert.deepEqual(d.zip.changed, ['Contents/section0.xml'], id);
    assert.equal(d.zip.orderChanged, false);
    assert.deepEqual(d.zip.methodChanged, []);
    assert.deepEqual(d.zip.added.concat(d.zip.removed), []);
    assert.equal(d.zip.rawIdentical.length, d.zip.names.a.length - 1, '나머지 항목은 압축 바이트까지 동일');
    const s = d.xml['Contents/section0.xml'];
    assert.deepEqual(s.tagCount, {}, '요소 수 변화 없음');
    assert.deepEqual(s.attrLost, {}); assert.deepEqual(s.attrAdded, {}); assert.deepEqual(s.attrChanged, {});
    assert.deepEqual(s.emptyCount['hp:t'][0], s.emptyCount['hp:t'][1]);
  }
});

test('기본 정책(drop): 구조 변화는 편집한 문단의 줄 배치 캐시 하나뿐이다', () => {
  for (const id of Object.keys(PRIOR)) {
    const orig = load(`hwpx/${id}.hwpx`);
    const doc = HwpxDocument.open(orig);
    assert.ok(doc.prependText('〔편집〕'));
    const s = diffPackages(orig, doc.save()).xml['Contents/section0.xml'];
    assert.deepEqual(Object.keys(s.tagCount).sort(), ['hp:linesegarray', 'hp:lineseg'].sort(), id);
    assert.equal(s.tagCount['hp:linesegarray'][0] - s.tagCount['hp:linesegarray'][1], 1, '문단 하나의 캐시만 삭제');
    assert.deepEqual(s.attrLost, {}); assert.deepEqual(s.attrChanged, {});
  }
});
