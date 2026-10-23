import { test, expect, type Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { ZipArchive } from '../src/zip.ts';
import { EditSession } from '../src/session.ts';
import { XlsxDocument } from '../src/xlsx.ts';
import { PdfDocument } from '../src/pdf.ts';
import assert from 'node:assert/strict';

const fx = (p: string) => new Uint8Array(readFileSync(new URL(`../fixtures/${p}`, import.meta.url)));

const HWPX_LABEL = 'HWPX 제품설명서 서식(표 많음)';
const DOCX_TABLE_LABEL = 'DOCX 표';

async function openSample(page: Page, label: string): Promise<void> {
  await page.goto('/');
  await page.locator('#sample').selectOption({ label });
  await expect(page.locator('#doc')).toBeVisible();
  await expect(page.locator('#paras textarea').first()).toBeVisible();
}

async function download(page: Page): Promise<Uint8Array> {
  const [dl] = await Promise.all([page.waitForEvent('download'), page.locator('#save').click()]);
  const path = await dl.path();
  expect(path).toBeTruthy();
  return new Uint8Array(readFileSync(path!));
}

test.describe('첫 화면', () => {
  test('가로 스크롤이 없고, 지원 범위와 개인정보 안내가 보이며, 콘솔 오류·외부 요청이 없다', async ({ page, baseURL }) => {
    const errors: string[] = [];
    const external: string[] = [];
    page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
    page.on('pageerror', e => errors.push(String(e)));
    page.on('request', r => { if (!r.url().startsWith(baseURL!) && !r.url().startsWith('data:') && !r.url().startsWith('blob:')) external.push(r.url()); });
    await page.goto('/');
    await expect(page.locator('#privacy')).toContainText('서버로 전송되지 않습니다');
    await expect(page.locator('#empty')).toContainText('아직 지원하지 않습니다');
    const overflow = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    expect(overflow).toBeLessThanOrEqual(0);
    expect(errors).toEqual([]);
    expect(external).toEqual([]);
  });

  test('터치 대상이 충분히 크다(버튼·선택 상자 높이 40px 이상)', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    const heights = await page.evaluate(() =>
      [...document.querySelectorAll<HTMLElement>('.toolbar .btn, .topbar .btn, .topbar select')].filter(e => e.offsetParent).map(e => e.getBoundingClientRect().height));
    expect(heights.length).toBeGreaterThan(2);
    for (const h of heights) expect(h).toBeGreaterThanOrEqual(40);
  });
});

test.describe('HWPX 흐름', () => {
  test('열기 → 편집 → 저장 → 다운로드한 파일을 독립적으로 검증', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    const first = page.locator('#paras textarea').first();
    const before = await first.inputValue();
    await first.fill(`${before} 모바일 & <시험>`);
    await expect(page.locator('#count')).toHaveText('변경 1');
    await expect(page.locator('.para.changed')).toHaveCount(1);
    const saved = await download(page);
    await expect(page.locator('#status')).toContainText('다시 열어');

    const orig = fx('hwpx/pps-manual-14-7.hwpx');
    const a = new ZipArchive(orig), b = new ZipArchive(saved);
    expect(b.entries.map(e => e.name)).toEqual(a.entries.map(e => e.name));
    const changed = a.entries.filter(e => Buffer.compare(Buffer.from(a.read(e.name)), Buffer.from(b.read(e.name))) !== 0).map(e => e.name);
    expect(changed).toEqual(['Contents/section0.xml']);
    const s = new EditSession(saved, 'x.hwpx');
    expect(s.paragraphs[0].text).toBe(`${before} 모바일 & <시험>`);
    expect(s.paragraphs.length).toBe(new EditSession(orig, 'o.hwpx').paragraphs.length);
  });

  test('수정이 없으면 저장본이 원본과 바이트까지 같다', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    const saved = await download(page);
    expect(Buffer.compare(Buffer.from(saved), Buffer.from(fx('hwpx/pps-manual-14-7.hwpx')))).toBe(0);
  });

  test('표 셀 문단을 고친다', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    const cell = page.locator('.para', { hasText: '표 셀' }).first().locator('textarea');
    const text = await cell.inputValue();
    await cell.fill(`${text}셀`);
    const saved = await download(page);
    const s = new EditSession(saved, 'x.hwpx');
    expect(s.paragraphs.some(p => p.context === 'cell' && p.text === `${text}셀`)).toBe(true);
  });

  test('탭·줄바꿈을 그대로 넣을 수 있고, 저장·재열기해도 유지된다(2026-10-23 추가)', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    const ta = page.locator('#paras textarea').nth(1);
    await ta.fill('가\n나\t다');
    await expect(ta).toHaveValue('가\n나\t다'); // 더 이상 공백으로 바뀌지 않는다
    const saved = await download(page);
    const s = new EditSession(saved, 'x.hwpx');
    expect(s.paragraphs[1].text).toBe('가\n나\t다');
  });

  test('탭·줄바꿈 단추로 커서 위치에 끼워 넣을 수 있다', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    const box = page.locator('.para').nth(1);
    const ta = box.locator('textarea');
    await ta.click();
    await page.keyboard.press('End');
    await box.getByRole('button', { name: '탭', exact: true }).click();
    await ta.pressSequentially('뒤');
    await box.getByRole('button', { name: '줄바꿈', exact: true }).click();
    await ta.pressSequentially('끝');
    await expect(ta).toHaveValue(/\t뒤\n끝$/);
  });

  test('되돌리기와 모두 되돌리기', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    const t0 = page.locator('#paras textarea').nth(0), t1 = page.locator('#paras textarea').nth(1);
    const o0 = await t0.inputValue();
    await t0.fill('변경A'); await t1.fill('변경B');
    await expect(page.locator('#count')).toHaveText('변경 2');
    await page.locator('.para').first().getByRole('button', { name: '되돌리기' }).click();
    await expect(t0).toHaveValue(o0);
    await expect(page.locator('#count')).toHaveText('변경 1');
    page.once('dialog', d => d.accept());
    await page.locator('#revert-all').click();
    await expect(page.locator('#count')).toHaveText('변경 0');
    await expect(page.locator('.para.changed')).toHaveCount(0);
  });

  test('도구 모음이 화면 안에 있고 문단 내용이 도구 모음에 가려지지 않게 여백이 있다', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    const vh = page.viewportSize()!.height;
    const box = (await page.locator('#toolbar').boundingBox())!;
    expect(box.y + box.height).toBeLessThanOrEqual(vh + 1);
    const padding = await page.evaluate(() => parseFloat(getComputedStyle(document.body).paddingBottom));
    expect(padding).toBeGreaterThanOrEqual(box.height - 2);
  });
});

test.describe('DOCX 흐름', () => {
  test('표 문서를 열고 셀을 고쳐 저장한다', async ({ page }) => {
    await openSample(page, DOCX_TABLE_LABEL);
    await expect(page.locator('#doc-format')).toHaveText('DOCX');
    const cell = page.locator('.para', { hasText: '표 셀' }).nth(1).locator('textarea');
    await cell.fill('고친 셀');
    const saved = await download(page);
    const s = new EditSession(saved, 'x.docx');
    expect(s.paragraphs.filter(p => p.context === 'cell').map(p => p.text)).toEqual(['셀 1-1', '고친 셀', '셀 2-1', '셀 2-2']);
  });
});

test.describe('임시 보존', () => {
  test('새로고침 뒤 이어서 하기로 수정이 복구된다', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    await page.locator('#paras textarea').first().fill('복구시험');
    await page.waitForTimeout(1200); // 자동 보존(0.8초 지연)
    await page.reload();
    await expect(page.locator('#draft-banner')).toBeVisible();
    await expect(page.locator('#draft-text')).toContainText('변경 1개');
    await page.locator('#draft-resume').click();
    await expect(page.locator('#paras textarea').first()).toHaveValue('복구시험');
    await expect(page.locator('#count')).toHaveText('변경 1');
  });

  test('임시 저장 지우기 뒤에는 복구 안내가 나오지 않는다', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    await page.locator('#paras textarea').first().fill('지울내용');
    await page.waitForTimeout(1200);
    await page.locator('#clear-draft').click();
    await expect(page.locator('#status')).toContainText('임시 보존을 지웠습니다'); // 삭제 완료 후에 새로고침
    await page.reload();
    await expect(page.locator('#draft-banner')).toBeHidden();
  });

  test('임시 보존을 끄면 저장된 내용이 지워지고 이후 수정은 저장되지 않는다', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    await page.locator('#paras textarea').first().fill('먼저');
    await page.waitForTimeout(1200);
    await page.locator('#autosave').uncheck();
    await page.locator('#paras textarea').first().fill('나중');
    await page.waitForTimeout(1200);
    await page.reload();
    await expect(page.locator('#draft-banner')).toBeHidden();
  });
});

test.describe('잘못된 파일', () => {
  test('내용이 없는 가짜 PDF는 열 수 없다고 알린다', async ({ page }) => {
    await page.goto('/');
    await page.locator('#file').setInputFiles({ name: 'a.pdf', mimeType: 'application/pdf', buffer: Buffer.from('%PDF-1.7 ' + ' '.repeat(60)) });
    await expect(page.locator('#status')).toContainText('열 수 없는 PDF');
    await expect(page.locator('#pdf')).toBeHidden();
  });

  test('DRM 암호화 파일은 열 수 없다고 알리고 우회하지 않는다', async ({ page }) => {
    await page.goto('/');
    const bytes = Buffer.alloc(4096);
    Buffer.from(String.fromCharCode(0x9b) + ' DRMONE  This Document is encrypted and protected by Fasoo DRM ', 'latin1').copy(bytes);
    await page.locator('#file').setInputFiles({ name: 'secured.hwpx', mimeType: 'application/octet-stream', buffer: bytes });
    await expect(page.locator('#status')).toContainText('DRM');
    await expect(page.locator('#status')).toContainText('우회하지 않습니다');
    await expect(page.locator('#doc')).toBeHidden();
  });

  test('구형 HWP는 HWPX와 다른 형식이라고 알린다', async ({ page }) => {
    await page.goto('/');
    const ole = Buffer.alloc(512); Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(ole);
    await page.locator('#file').setInputFiles({ name: 'a.hwp', mimeType: 'application/octet-stream', buffer: ole });
    await expect(page.locator('#status')).toContainText('HWPX와 다른 형식');
  });

  test('손상된 DOCX와 빈 파일은 오류를 알리고 앱이 계속 동작한다', async ({ page }) => {
    await page.goto('/');
    const good = Buffer.from(fx('docx/basic.docx'));
    await page.locator('#file').setInputFiles({ name: 'broken.docx', mimeType: 'application/octet-stream', buffer: good.subarray(0, good.length - 40) });
    await expect(page.locator('#status')).toContainText(/열 수 없는|지원하는 형식이 아닙니다/);
    await page.locator('#file').setInputFiles({ name: 'empty.docx', mimeType: 'application/octet-stream', buffer: Buffer.alloc(0) });
    await expect(page.locator('#status')).toContainText(/열 수 없는|지원하는 형식이 아닙니다/);
    await page.locator('#sample').selectOption({ label: HWPX_LABEL });
    await expect(page.locator('#doc')).toBeVisible();
  });

  test('직접 고른 HWPX 파일을 열고 저장한다(파일 선택기 경로)', async ({ page }) => {
    await page.goto('/');
    await page.locator('#file').setInputFiles({ name: '내문서.hwpx', mimeType: 'application/octet-stream', buffer: Buffer.from(fx('hwpx/pps-spec-14-13.hwpx')) });
    await expect(page.locator('#doc-name')).toHaveText('내문서.hwpx');
    await page.locator('#paras textarea').first().fill('선택기경로');
    const [dl] = await Promise.all([page.waitForEvent('download'), page.locator('#save').click()]);
    expect(dl.suggestedFilename()).toBe('내문서-수정.hwpx');
  });
});

test.describe('쪽 미리보기(읽기 전용)와 정보 화면', () => {
  async function pagesLoaded(page: Page, n: number): Promise<void> {
    await expect(page.locator('#pages img.page')).toHaveCount(n);
    await page.waitForFunction((count) => [...document.querySelectorAll<HTMLImageElement>('#pages img.page')].filter(i => i.complete && i.naturalWidth > 0).length === count, n);
  }

  test('HWP(바이너리)는 편집 없이 쪽 미리보기만 제공한다', async ({ page }) => {
    await page.goto('/');
    await page.locator('#file').setInputFiles({ name: 'nabo.hwp', mimeType: 'application/octet-stream', buffer: Buffer.from(readFileSync(new URL('../fixtures/hwp/assembly-nabo.hwp', import.meta.url))) });
    await expect(page.locator('#doc-format')).toContainText('HWP');
    await expect(page.locator('#paras textarea')).toHaveCount(0);
    await expect(page.locator('#toolbar')).toBeHidden();
    await page.locator('#prev-open').click();
    await expect(page.locator('#prev-info')).toContainText('총 2쪽', { timeout: 25_000 });
    await pagesLoaded(page, 2);
  });

  test('HWPX 편집본의 쪽 미리보기가 현재 쪽 수로 그려진다', async ({ page }) => {
    await openSample(page, HWPX_LABEL);
    await page.locator('#paras textarea').first().fill('미리보기 시험');
    await page.locator('#prev-open').click();
    await expect(page.locator('#prev-info')).toContainText('총 3쪽', { timeout: 25_000 });
    await pagesLoaded(page, 3);
    await expect(page.locator('#prev-more')).toBeHidden();
  });

  test('DOCX에는 쪽 미리보기 단추가 없다', async ({ page }) => {
    await openSample(page, DOCX_TABLE_LABEL);
    await expect(page.locator('#preview')).toBeHidden();
  });

  test('정보 화면에 한컴 공개 문서 표기와 상표 고지가 있다', async ({ page }) => {
    await page.goto('/');
    await page.locator('#about-open').click();
    await expect(page.locator('#about')).toBeVisible();
    await expect(page.locator('#about')).toContainText('본 제품은 한컴의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.');
    await expect(page.locator('#about')).toContainText('상표');
    await page.locator('#about-close').click();
    await expect(page.locator('#about')).toBeHidden();
  });
});

test.describe('XLSX 흐름', () => {
  const XLSX_LABEL = 'XLSX 값 종류(숫자·글자·날짜·수식)';
  const cell = (page: Page, ref: string) => page.locator(`#grid td[data-ref="${ref}"]`);

  test('칸을 고치면 숫자는 숫자로 저장되고 나머지 칸은 그대로다', async ({ page }) => {
    await openSampleSheet(page);
    await expect(page.locator('#doc-format')).toHaveText('XLSX');
    await expect(cell(page, 'B1')).toHaveText('42');
    await cell(page, 'B1').click();
    await page.locator('#cell-input').fill('43');
    await expect(cell(page, 'B1')).toHaveText('43');
    await expect(page.locator('#count')).toHaveText('변경 1');
    await cell(page, 'A11').click();
    await page.locator('#cell-input').fill('바뀐말');
    await expect(page.locator('#count')).toHaveText('변경 2');
    const saved = await download(page);
    const d = XlsxDocument.open(saved);
    assert.deepEqual([d.cell(0, 'B1').kind, d.cell(0, 'B1').text], ['number', '43']);
    assert.equal(d.cell(0, 'A11').text, '바뀐말');
    assert.equal(d.cell(0, 'B11').text, '같은말', '같은 공유 문자열을 쓰던 다른 칸은 그대로');
    assert.equal(d.cell(0, 'B2').text, '3.14159');
    const a = new ZipArchive(fx('xlsx/types.xlsx')), b = new ZipArchive(saved);
    assert.deepEqual(b.read('xl/sharedStrings.xml'), a.read('xl/sharedStrings.xml'));
  });

  test('숫자 칸에 글자를 넣으면 오류를 보이고 값은 바뀌지 않는다', async ({ page }) => {
    await openSampleSheet(page);
    await cell(page, 'B1').click();
    await page.locator('#cell-input').fill('abc');
    await expect(page.locator('#cell-msg')).toContainText('숫자만');
    await expect(page.locator('#cell-input')).toHaveClass(/invalid/);
    await expect(cell(page, 'B1')).toHaveText('42');
    await expect(page.locator('#count')).toHaveText('변경 0');
  });

  test('수식 칸은 읽기 전용이고 날짜 칸은 날짜로 보인다', async ({ page }) => {
    await openSampleSheet(page);
    await cell(page, 'B9').click();
    await expect(page.locator('#cell-input')).toHaveJSProperty('readOnly', true);
    await expect(page.locator('#cell-msg')).toContainText('SUM(B1:B3)');
    await expect(cell(page, 'B7')).toHaveText('2026-10-22');
  });

  test('시트를 바꾸고, 모두 되돌리면 원본과 같은 파일이 저장된다', async ({ page }) => {
    await openSampleSheet(page);
    await page.locator('#sheet-select').selectOption({ index: 1 });
    await expect(cell(page, 'B1')).toHaveText('100');
    await cell(page, 'B1').click();
    await page.locator('#cell-input').fill('200');
    await expect(page.locator('#count')).toHaveText('변경 1');
    page.once('dialog', dlg => dlg.accept());
    await page.locator('#revert-all').click();
    await expect(page.locator('#count')).toHaveText('변경 0');
    await expect(cell(page, 'B1')).toHaveText('100');
    const saved = await download(page);
    assert.equal(Buffer.compare(Buffer.from(saved), Buffer.from(fx('xlsx/types.xlsx'))), 0);
  });

  async function openSampleSheet(page: Page): Promise<void> {
    await page.goto('/');
    await page.locator('#sample').selectOption({ label: XLSX_LABEL });
    await expect(page.locator('#sheet')).toBeVisible();
    await expect(page.locator('#grid td').first()).toBeVisible();
  }
});

test.describe('PPTX 흐름', () => {
  test('슬라이드 글자를 고쳐 저장하고, 필드 문단은 보기만 가능하다', async ({ page }) => {
    await page.goto('/');
    await page.locator('#sample').selectOption({ label: 'PPTX 슬라이드 3장(표·노트)' });
    await expect(page.locator('#doc-format')).toHaveText('PPTX');
    await expect(page.locator('.para', { hasText: '슬라이드 1 ·' }).first()).toBeVisible();
    await expect(page.locator('.para', { hasText: '슬라이드 3 노트' })).toHaveCount(1);
    const locked = page.locator('.para', { hasText: '필드 등이 섞여 보기만 가능' }).locator('textarea');
    await expect(locked).toHaveJSProperty('readOnly', true);
    const first = page.locator('#paras textarea').first();
    await first.fill('PPTX 편집 시험');
    const saved = await download(page);
    const s = new EditSession(saved, 'x.pptx');
    expect(s.paragraphs[0].text).toBe('PPTX 편집 시험');
    expect(s.paragraphs.length).toBe(new EditSession(fx('pptx/basic.pptx'), 'o.pptx').paragraphs.length);
  });
});

test.describe('PDF 흐름', () => {
  async function openPdf(page: Page): Promise<void> {
    await page.goto('/');
    await page.locator('#sample').selectOption({ label: 'PDF 3쪽(메모 추가)' });
    await expect(page.locator('#pdf')).toBeVisible();
    await expect(page.locator('#status')).toContainText('3쪽', { timeout: 30_000 });
    await page.waitForFunction(() => {
      const c = document.getElementById('pdf-canvas') as HTMLCanvasElement | null;
      if (!c || c.width === 0) return false;
      const d = c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data;
      for (let i = 0; i < d.length; i += 4) if (d[i] < 200) return true; // 글자 픽셀(어두운 점)이 있다
      return false;
    }, undefined, { timeout: 30_000 });
  }

  test('열람: 그려지고, 쪽 이동과 확대·축소가 된다', async ({ page }) => {
    await openPdf(page);
    await expect(page.locator('#pdf-total')).toHaveText('3');
    const w0 = await page.locator('#pdf-canvas').evaluate((c: HTMLCanvasElement) => c.getBoundingClientRect().width);
    await page.locator('#pdf-zoom-in').click();
    await expect.poll(() => page.locator('#pdf-canvas').evaluate((c: HTMLCanvasElement) => c.getBoundingClientRect().width)).toBeGreaterThan(w0 * 1.2);
    await page.locator('#pdf-fit').click();
    await page.locator('#pdf-next').click();
    await expect(page.locator('#pdf-page')).toHaveValue('2');
    await page.locator('#pdf-prev').click();
    await expect(page.locator('#pdf-page')).toHaveValue('1');
  });

  test('메모 추가 → 저장 → 내려받은 파일이 원본을 그대로 포함하고 메모가 들어 있다', async ({ page }) => {
    await openPdf(page);
    await page.locator('#pdf-memo').click();
    await page.locator('#pdf-stage').click({ position: { x: 80, y: 120 } });
    await expect(page.locator('#pdf-editor')).toBeVisible();
    await page.locator('#pdf-note-input').fill('검토 메모 & <확인>');
    await page.locator('#pdf-note-add').click();
    await expect(page.locator('#count')).toHaveText('변경 1');
    await expect(page.locator('#pdf-marks .note-mark.pending')).toHaveCount(1);
    const saved = await download(page);
    const orig = fx('pdf/classic.pdf');
    assert.deepEqual(saved.subarray(0, orig.length), orig, '원본 바이트가 앞부분에 그대로');
    const d = PdfDocument.open(saved);
    assert.deepEqual(d.notes().map(n => [n.page, n.contents]), [[0, '검토 메모 & <확인>']]);
    assert.equal(d.pages().length, 3);
  });

  test('메모를 지우면 원본과 같은 파일이 저장된다', async ({ page }) => {
    await openPdf(page);
    await page.locator('#pdf-memo').click();
    await page.locator('#pdf-stage').click({ position: { x: 60, y: 60 } });
    await page.locator('#pdf-note-input').fill('지울 메모');
    await page.locator('#pdf-note-add').click();
    await expect(page.locator('#count')).toHaveText('변경 1');
    page.once('dialog', dlg => dlg.accept());
    await page.locator('#revert-all').click();
    await expect(page.locator('#count')).toHaveText('변경 0');
    const saved = await download(page);
    assert.equal(Buffer.compare(Buffer.from(saved), Buffer.from(fx('pdf/classic.pdf'))), 0);
  });

  test('임시 보존에서 메모가 복구된다', async ({ page }) => {
    await openPdf(page);
    await page.locator('#pdf-memo').click();
    await page.locator('#pdf-stage').click({ position: { x: 70, y: 90 } });
    await page.locator('#pdf-note-input').fill('복구될 메모');
    await page.locator('#pdf-note-add').click();
    await page.waitForTimeout(1200);
    await page.reload();
    await page.locator('#draft-resume').click();
    await expect(page.locator('#pdf-notes li')).toContainText('복구될 메모', { timeout: 30_000 });
    await expect(page.locator('#count')).toHaveText('변경 1');
  });
});
