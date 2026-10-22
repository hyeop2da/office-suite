import { test, expect, type Page } from '@playwright/test';

async function swReady(page: Page): Promise<void> {
  await page.goto('/');
  await page.evaluate(async () => { await navigator.serviceWorker.ready; });
  // 앱 파일 미리 저장이 끝날 때까지(설치 이벤트가 캐시를 채운 뒤 활성화됨)
  await expect.poll(() => page.evaluate(async () => {
    const keys = await caches.keys();
    if (!keys.length) return 0;
    return (await (await caches.open(keys[0])).keys()).length;
  })).toBeGreaterThan(8);
  // 서비스 워커가 이 페이지를 제어할 때까지 다시 열기
  await page.reload();
  await expect.poll(() => page.evaluate(() => !!navigator.serviceWorker.controller)).toBe(true);
}

test.describe('오프라인', () => {
  test('앱 파일이 캐시되어 네트워크 없이 새로고침·샘플 열기·저장 검증이 된다', async ({ page, context }) => {
    await swReady(page);
    await context.setOffline(true);
    await page.reload();
    await expect(page.locator('#privacy')).toContainText('서버로 전송되지 않습니다');
    await page.locator('#sample').selectOption({ label: 'DOCX 표' });
    await expect(page.locator('#doc-format')).toHaveText('DOCX');
    await page.locator('#paras textarea').first().fill('오프라인 편집');
    await expect(page.locator('#count')).toHaveText('변경 1');
    const [dl] = await Promise.all([page.waitForEvent('download'), page.locator('#save').click()]);
    expect(dl.suggestedFilename()).toBe('table-수정.docx');
    await context.setOffline(false);
  });

  test('캐시에는 앱 파일만 있고 사용자가 연 문서는 들어가지 않는다', async ({ page }) => {
    await swReady(page);
    await page.locator('#file').setInputFiles({ name: 'secret-내문서.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('PK\x03\x04 not a real docx') });
    const urls = await page.evaluate(async () => {
      const out: string[] = [];
      for (const k of await caches.keys()) for (const r of await (await caches.open(k)).keys()) out.push(r.url);
      return out;
    });
    expect(urls.some(u => /secret|%EB%82%B4/.test(u))).toBe(false);
    expect(urls.every(u => u.startsWith(new URL(page.url()).origin))).toBe(true);
  });

  test('매니페스트와 아이콘이 있다(홈 화면 추가)', async ({ page, request, baseURL }) => {
    await page.goto('/');
    const href = await page.locator('link[rel=manifest]').getAttribute('href');
    const m = await (await request.get(new URL(href!, baseURL).href)).json();
    expect(m.name).toBe('문서 작업실');
    expect(m.display).toBe('standalone');
    for (const icon of m.icons) expect((await request.get(new URL(icon.src, baseURL).href)).ok()).toBe(true);
    expect(await page.locator('link[rel=apple-touch-icon]').getAttribute('href')).toBe('icon-180.png');
  });
});
