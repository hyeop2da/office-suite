// PDF 화면: pdf.js(Apache-2.0)로 쪽을 그리고, 메모 모드에서 쪽을 누르면 텍스트 메모를 추가한다.
// pdf.js는 PDF를 열 때만 내려받는다. 메모는 PdfSession이 증분 갱신으로 저장한다.
import type { PdfSession } from '../pdf-session.ts';

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;

interface PdfJs {
  GlobalWorkerOptions: { workerSrc: string };
  getDocument(o: unknown): { promise: Promise<PdfDocProxy> };
}
interface Viewport { width: number; height: number; convertToPdfPoint(x: number, y: number): [number, number]; convertToViewportPoint(x: number, y: number): [number, number] }
interface PdfPageProxy { getViewport(o: { scale: number }): Viewport; render(o: unknown): { promise: Promise<void>; cancel(): void } }
interface PdfDocProxy { numPages: number; getPage(n: number): Promise<PdfPageProxy>; destroy(): Promise<void> }

let lib: Promise<PdfJs> | undefined;
async function loadPdfJs(): Promise<PdfJs> {
  if (!lib) {
    lib = (async () => {
      const m = (await import('pdfjs-dist/build/pdf.min.mjs')) as unknown as PdfJs;
      m.GlobalWorkerOptions.workerSrc = new URL('pdf.worker.min.mjs', document.baseURI).href;
      return m;
    })();
    lib.catch(() => { lib = undefined; });
  }
  return lib;
}

export interface PdfHooks { onChange(): void; status(msg: string, kind?: '' | 'ok' | 'err' | 'warn'): void }
export interface PdfViewHandle { refresh(): void; destroy(): void }

export async function mountPdf(s: PdfSession, hooks: PdfHooks): Promise<PdfViewHandle> {
  const pdfjs = await loadPdfJs();
  const proxy = await pdfjs.getDocument({
    data: s.original.slice(),
    isEvalSupported: false,
    cMapUrl: new URL('cmaps/', document.baseURI).href, cMapPacked: true,
    standardFontDataUrl: new URL('standard_fonts/', document.baseURI).href,
  }).promise;
  const total = proxy.numPages;
  const canvas = $('pdf-canvas') as HTMLCanvasElement;
  const marks = $('pdf-marks');
  const stage = $('pdf-stage');
  const pageInput = $('pdf-page') as HTMLInputElement;
  const memoBtn = $('pdf-memo') as HTMLButtonElement;
  const editor = $('pdf-editor');
  const noteInput = $('pdf-note-input') as HTMLTextAreaElement;
  let cur = 0, zoom = 1, memoMode = false, token = 0;
  let view: Viewport | null = null;
  let draft: { x: number; y: number } | null = null;
  let renderTask: { cancel(): void } | null = null;

  $('pdf-total').textContent = String(total);
  pageInput.max = String(total);

  async function render(): Promise<void> {
    const my = ++token;
    renderTask?.cancel();
    const page = await proxy.getPage(cur + 1);
    if (my !== token) return;
    const fit = Math.max(120, ($('pdf-wrap').clientWidth || 360) - 2) / page.getViewport({ scale: 1 }).width;
    const scale = fit * zoom;
    const vp = page.getViewport({ scale });
    const dpr = Math.min(window.devicePixelRatio || 1, 2, Math.sqrt(16_000_000 / (vp.width * vp.height)));
    canvas.width = Math.floor(vp.width * dpr); canvas.height = Math.floor(vp.height * dpr);
    canvas.style.width = `${vp.width}px`; canvas.style.height = `${vp.height}px`;
    stage.style.width = `${vp.width}px`; stage.style.height = `${vp.height}px`;
    const ctx = canvas.getContext('2d')!;
    const task = page.render({ canvasContext: ctx, viewport: page.getViewport({ scale: scale * dpr }) });
    renderTask = task;
    try { await task.promise; } catch (e) { if ((e as Error).name === 'RenderingCancelledException') return; throw e; }
    if (my !== token) return;
    view = vp;
    drawMarks();
  }

  function drawMarks(): void {
    marks.replaceChildren();
    if (!view) return;
    const put = (x: number, y: number, text: string, pending: boolean): void => {
      const [vx, vy] = view!.convertToViewportPoint(x, y);
      const m = document.createElement('button');
      m.className = `note-mark${pending ? ' pending' : ''}`;
      m.textContent = '메모';
      m.title = text;
      m.style.left = `${Math.max(0, Math.min(vx, view!.width - 36))}px`;
      m.style.top = `${Math.max(0, Math.min(vy, view!.height - 24))}px`;
      m.addEventListener('click', (ev) => { ev.stopPropagation(); hooks.status(`${pending ? '추가한' : '기존'} 메모: ${text}`); });
      marks.append(m);
    };
    for (const n of s.existing) if (n.page === cur) put(n.x, n.y, n.contents, false);
    for (const p of s.pending()) if (p.page === cur) put(p.x, p.y, p.contents, true);
  }

  function renderList(): void {
    const list = $('pdf-notes');
    list.replaceChildren();
    const rows: { page: number; text: string; key?: string }[] = [
      ...s.existing.map(n => ({ page: n.page, text: n.contents })),
      ...s.pending().map(p => ({ page: p.page, text: p.contents, key: p.key })),
    ].sort((a, b) => a.page - b.page);
    $('pdf-notes-title').textContent = `메모 ${rows.length}개 (새로 추가 ${s.changeCount})`;
    for (const r of rows) {
      const li = document.createElement('li');
      const t = document.createElement('span');
      t.textContent = `${r.page + 1}쪽 · ${r.text}${r.key ? ' (새 메모)' : ''}`;
      li.append(t);
      if (r.key) {
        const del = document.createElement('button');
        del.className = 'btn revert'; del.textContent = '지우기';
        del.addEventListener('click', () => { s.revert(r.key!); drawMarks(); renderList(); hooks.onChange(); });
        li.append(del);
      }
      list.append(li);
    }
  }

  function go(n: number): void {
    cur = Math.max(0, Math.min(total - 1, n));
    pageInput.value = String(cur + 1);
    closeEditor();
    void render();
  }
  function closeEditor(): void { draft = null; editor.hidden = true; noteInput.value = ''; }

  $('pdf-prev').onclick = () => go(cur - 1);
  $('pdf-next').onclick = () => go(cur + 1);
  pageInput.onchange = () => go(Number(pageInput.value) - 1);
  $('pdf-zoom-in').onclick = () => { zoom = Math.min(4, zoom * 1.25); void render(); };
  $('pdf-zoom-out').onclick = () => { zoom = Math.max(0.5, zoom / 1.25); void render(); };
  $('pdf-fit').onclick = () => { zoom = 1; void render(); };
  memoBtn.onclick = () => {
    memoMode = !memoMode;
    memoBtn.classList.toggle('primary', memoMode);
    memoBtn.textContent = memoMode ? '메모 모드 켜짐: 쪽을 눌러 위치 지정' : '메모 추가';
    stage.classList.toggle('memo', memoMode);
    if (!memoMode) closeEditor();
  };
  stage.addEventListener('click', (ev) => {
    if (!memoMode || !view) return;
    const r = canvas.getBoundingClientRect();
    const [x, y] = view.convertToPdfPoint(ev.clientX - r.left, ev.clientY - r.top);
    draft = { x, y };
    editor.hidden = false;
    noteInput.focus();
  });
  $('pdf-note-add').onclick = () => {
    if (!draft) return;
    const text = noteInput.value;
    if (text.trim() === '') { hooks.status('메모 글자를 입력하세요.', 'warn'); return; }
    try { s.addNote(cur, draft.x, draft.y, text); }
    catch (e) { hooks.status((e as Error).message, 'err'); return; }
    closeEditor(); drawMarks(); renderList(); hooks.onChange();
  };
  $('pdf-note-cancel').onclick = closeEditor;

  go(0);
  renderList();
  return {
    refresh: () => { drawMarks(); renderList(); },
    destroy: () => { token++; renderTask?.cancel(); void proxy.destroy(); },
  };
}
