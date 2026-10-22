// 읽기 전용 쪽 미리보기. rhwp(MIT, WASM)를 처음 쓸 때만 내려받는다. 저장 경로에는 쓰지 않는다.
// 미리보기는 글꼴·줄바꿈이 한/글과 다를 수 있다.
let ready: Promise<typeof import('@rhwp/core')> | undefined;

function installMeasure(): void {
  let ctx: CanvasRenderingContext2D | null = null;
  let last = '';
  (globalThis as unknown as Record<string, unknown>).measureTextWidth = (font: string, text: string): number => {
    if (!ctx) ctx = document.createElement('canvas').getContext('2d');
    if (!ctx) return text.length * 10;
    if (font !== last) { ctx.font = font; last = font; }
    return ctx.measureText(text).width;
  };
}

async function loadRhwp(): Promise<typeof import('@rhwp/core')> {
  if (!ready) {
    ready = (async () => {
      installMeasure();
      const m = await import('@rhwp/core');
      await m.default({ module_or_path: new URL('rhwp_bg.wasm', document.baseURI).href });
      return m;
    })();
    ready.catch(() => { ready = undefined; }); // 실패하면 다음에 다시 시도
  }
  return ready;
}

export interface PreviewDoc {
  pageCount: number;
  /** 쪽 하나를 SVG 문자열로 그린다. 화면에는 <img>(blob URL)로만 넣어 스크립트가 실행되지 않게 한다. */
  svg(page: number): string;
  free(): void;
}

export async function openPreview(bytes: Uint8Array): Promise<PreviewDoc> {
  const m = await loadRhwp();
  const doc = new m.HwpDocument(bytes);
  return { pageCount: doc.pageCount(), svg: (p: number) => doc.renderPageSvg(p), free: () => doc.free() };
}
