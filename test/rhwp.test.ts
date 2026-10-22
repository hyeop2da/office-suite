// 독립 파서(rhwp, MIT) 교차 검증: 우리가 고친 HWPX를 rhwp가 같은 쪽 수로 열고 고친 글자를 그리는지 본다.
// 한/글 2022 확인은 아니다(한글-비교-체크리스트.md 참고).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import * as rhwp from '@rhwp/core';
import { HwpxDocument } from '../src/hwpx.ts';

(globalThis as Record<string, unknown>).measureTextWidth = (_f: string, t: string) => [...t].length * 10;
rhwp.initSync({ module: readFileSync(new URL('../node_modules/@rhwp/core/rhwp_bg.wasm', import.meta.url)) });

const dir = new URL('../fixtures/hwpx/', import.meta.url);
const fixtures = readdirSync(dir).filter(f => f.endsWith('.hwpx'));
const MARK = '〔편집시험〕';

function svgText(svg: string): string {
  return [...svg.matchAll(/<text[^>]*>([^<]*)<\/text>/g)].map(m => m[1]).join('');
}
function open(bytes: Uint8Array) { return new rhwp.HwpDocument(bytes); }

for (const f of fixtures) {
  for (const layout of ['keep', 'drop'] as const) {
    test(`${f} [layout=${layout}]: rhwp가 같은 쪽 수로 열고 마커를 그린다`, () => {
      const orig = new Uint8Array(readFileSync(new URL(f, dir)));
      const base = open(orig);
      const doc = HwpxDocument.open(orig);
      assert.ok(doc.prependText(MARK, { layout }));
      const edited = open(doc.save());
      assert.equal(edited.pageCount(), base.pageCount(), '쪽 수');
      assert.ok(svgText(edited.renderPageSvg(0)).includes(MARK), '1쪽에 마커 글자가 그려짐');
      assert.ok(!svgText(base.renderPageSvg(0)).includes(MARK));
    });
  }

  test(`${f}: 무수정 저장본은 rhwp 출력이 원본과 같다`, () => {
    const orig = new Uint8Array(readFileSync(new URL(f, dir)));
    const a = open(orig), b = open(HwpxDocument.open(orig).save());
    assert.equal(b.pageCount(), a.pageCount());
    for (let p = 0; p < a.pageCount(); p++) assert.equal(b.renderPageSvg(p), a.renderPageSvg(p), `${p + 1}쪽 SVG`);
  });
}

// HWP(바이너리) 쪽 수: rhwp 값과 한/글 2022(12.0.0.4426)가 상태 표시줄에 보인 값(2026-10-22, 사람이 화면으로 확인)
const HANGUL_PAGES: Record<string, { rhwp: number; hangul: number }> = {
  'assembly-audit': { rhwp: 2, hangul: 2 },
  'assembly-briefing': { rhwp: 25, hangul: 26 }, // 불일치: 미리보기는 쪽 수가 다를 수 있다
  'assembly-nabo': { rhwp: 2, hangul: 2 },
  'seoul-judge-sample': { rhwp: 1, hangul: 1 },
  'seoul-menu-3e': { rhwp: 2, hangul: 2 },
};
for (const [id, exp] of Object.entries(HANGUL_PAGES)) {
  test(`HWP 쪽 수 ${id}: rhwp ${exp.rhwp}쪽 (한/글 ${exp.hangul}쪽${exp.rhwp === exp.hangul ? '과 일치' : '과 불일치, 기록용'})`, () => {
    const d = new rhwp.HwpDocument(new Uint8Array(readFileSync(new URL(`../fixtures/hwp/${id}.hwp`, import.meta.url))));
    assert.equal(d.pageCount(), exp.rhwp);
  });
}
