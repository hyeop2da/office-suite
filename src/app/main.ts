// 모바일 웹 UI. 문서 처리는 src/session.ts와 src/engine.ts가 맡고, 이 파일은 화면과 파일 입출력만 한다.
import { EditSession } from '../session.ts';
import { SheetSession } from '../sheet-session.ts';
import { mountSheet } from './sheet-view.ts';
import { PdfSession } from '../pdf-session.ts';
import { mountPdf, type PdfViewHandle } from './pdf-view.ts';
import { detectFormat, UnsupportedFormatError } from '../engine.ts';
import { loadDraft, saveDraft, deleteDraft, type Draft } from './draft.ts';
import { openPreview, type PreviewDoc } from './preview.ts';

const MAX_BYTES = 100 * 1024 * 1024;
const WARN_BYTES = 30 * 1024 * 1024;
const PAGE = 150;
const CONTEXT_LABEL: Record<string, string> = {
  body: '본문', cell: '표 셀', textbox: '글상자', header: '머리말', footer: '꼬리말', footnote: '각주', endnote: '미주',
};
const SAMPLES: [string, string][] = [
  ['samples/pps-manual-14-7.hwpx', 'HWPX 제품설명서 서식(표 많음)'],
  ['samples/pipc-privacy-20231127.hwpx', 'HWPX 개인정보 안내(그림·글상자)'],
  ['samples/basic.docx', 'DOCX 기본'],
  ['samples/table.docx', 'DOCX 표'],
  ['samples/types.xlsx', 'XLSX 값 종류(숫자·글자·날짜·수식)'],
  ['samples/basic.pptx', 'PPTX 슬라이드 3장(표·노트)'],
  ['samples/classic.pdf', 'PDF 3쪽(메모 추가)'],
];

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;
let session: EditSession | null = null;
let sheet: SheetSession | null = null;
let sheetView: { refresh(): void } | null = null;
let pdf: PdfSession | null = null;
let pdfView: PdfViewHandle | null = null;
const active = (): EditSession | SheetSession | PdfSession | null => session ?? sheet ?? pdf;

function resetPdf(): void {
  pdfView?.destroy();
  pdfView = null;
  pdf = null;
  $('pdf').hidden = true;
}
let shown = 0;
let saveTimer: number | undefined;
const textareas = new Map<string, HTMLTextAreaElement>();

// ---- 쪽 미리보기(읽기 전용) ----
let previewSource: (() => Uint8Array) | null = null;
let previewDoc: PreviewDoc | null = null;
let previewShown = 0;
let pageUrls: string[] = [];
const PREVIEW_STEP = 3;

function resetPreview(): void {
  previewDoc?.free();
  previewDoc = null;
  previewShown = 0;
  for (const u of pageUrls) URL.revokeObjectURL(u);
  pageUrls = [];
  $('pages').replaceChildren();
  $('prev-more').hidden = true;
  $('prev-info').textContent = '';
}

function renderMorePages(): void {
  if (!previewDoc) return;
  const end = Math.min(previewDoc.pageCount, previewShown + PREVIEW_STEP);
  for (let i = previewShown; i < end; i++) {
    // SVG는 <img>(blob URL)로만 넣는다: 이미지 안의 스크립트는 실행되지 않는다.
    const url = URL.createObjectURL(new Blob([previewDoc.svg(i)], { type: 'image/svg+xml' }));
    pageUrls.push(url);
    const img = document.createElement('img');
    img.className = 'page';
    img.alt = `${i + 1}쪽`;
    img.src = url;
    $('pages').append(img);
  }
  previewShown = end;
  $('prev-info').textContent = `총 ${previewDoc.pageCount}쪽 중 ${previewShown}쪽 표시`;
  $('prev-more').hidden = previewShown >= previewDoc.pageCount;
}

async function onPreview(): Promise<void> {
  if (!previewSource) return;
  const btn = $('prev-open') as HTMLButtonElement;
  btn.disabled = true;
  $('prev-info').textContent = '불러오는 중… (처음에는 큰 파일을 내려받습니다)';
  try {
    const bytes = previewSource();
    resetPreview();
    previewDoc = await openPreview(bytes);
    renderMorePages();
  } catch (e) {
    $('prev-info').textContent = '';
    setStatus(`미리보기를 만들지 못했습니다: ${(e as Error).message}`, 'err');
  } finally { btn.disabled = false; }
}

function showViewOnly(bytes: Uint8Array, name: string, info: { label: string; note: string }): void {
  session = null;
  sheet = null;
  sheetView = null;
  $('sheet').hidden = true;
  resetPdf();
  textareas.clear();
  $('paras').replaceChildren();
  ($('more') as HTMLButtonElement).hidden = true;
  $('empty').hidden = true;
  $('doc').hidden = false;
  $('toolbar').hidden = true;
  $('doc-name').textContent = name;
  $('doc-format').textContent = info.label;
  $('doc-note').textContent = info.note;
  resetPreview();
  previewSource = () => bytes;
  $('preview').hidden = false;
  document.documentElement.style.setProperty('--toolbar-h', '0px');
  setStatus(`${info.label}: ${info.note}`, 'warn');
}

function setStatus(msg: string, kind: '' | 'ok' | 'err' | 'warn' = ''): void {
  const el = $('status');
  el.textContent = msg;
  el.className = kind;
}

function updateCount(): void {
  const a = active();
  $('count').textContent = `변경 ${a ? a.changeCount : 0}`;
  ($('save') as HTMLButtonElement).disabled = !a;
  ($('revert-all') as HTMLButtonElement).disabled = !a || a.changeCount === 0;
}

function autosaveEnabled(): boolean { return ($('autosave') as HTMLInputElement).checked; }

function scheduleSave(): void {
  if (!active() || !autosaveEnabled()) return;
  window.clearTimeout(saveTimer);
  saveTimer = window.setTimeout(async () => {
    const a = active();
    if (!a) return;
    const ok = await saveDraft({ name: a.filename, bytes: a.original, edits: [...a.edits], savedAt: Date.now() });
    if (!ok) setStatus('임시 보존에 실패했습니다(브라우저 저장 공간 제한일 수 있음). 저장·다운로드로 직접 보관하세요.', 'warn');
  }, 800);
}

function autosize(t: HTMLTextAreaElement): void { t.style.height = 'auto'; t.style.height = `${t.scrollHeight}px`; }

function renderParas(): void {
  if (!session) return;
  const host = $('paras');
  const list = session.paragraphs;
  const end = Math.min(list.length, shown + PAGE);
  for (let i = shown; i < end; i++) {
    const p = list[i];
    const box = document.createElement('div');
    box.className = 'para';
    const meta = document.createElement('div');
    meta.className = 'meta';
    const label = document.createElement('span');
    label.textContent = `${i + 1}. ${p.group ? `${p.group} · ` : ''}${CONTEXT_LABEL[p.context] ?? p.context}${p.partlyLocked ? ' · 탭·필드 등이 섞여 보기만 가능' : ''}`;
    const rv = document.createElement('button');
    rv.className = 'btn revert';
    rv.textContent = '되돌리기';
    rv.hidden = true;
    meta.append(label, rv);
    const ta = document.createElement('textarea');
    ta.value = session.edits.get(p.id) ?? p.text;
    ta.rows = 1;
    ta.readOnly = p.partlyLocked;
    ta.setAttribute('aria-label', `${i + 1}번 문단`);
    ta.setAttribute('enterkeyhint', 'done');
    textareas.set(p.id, ta);
    const refresh = (): void => { const ch = session!.edits.has(p.id); box.classList.toggle('changed', ch); rv.hidden = !ch; };
    ta.addEventListener('input', () => {
      // 줄바꿈과 탭은 아직 지원하지 않으므로 입력에서 걸러 낸다.
      const clean = ta.value.replace(/[\r\n\t]+/g, ' ');
      if (clean !== ta.value) { ta.value = clean; setStatus('줄바꿈·탭은 아직 넣을 수 없어 공백으로 바꿨습니다.', 'warn'); }
      session!.setText(p.id, ta.value);
      refresh(); updateCount(); autosize(ta); scheduleSave();
    });
    rv.addEventListener('click', () => { session!.revert(p.id); ta.value = p.text; refresh(); updateCount(); autosize(ta); scheduleSave(); });
    refresh();
    box.append(meta, ta);
    host.append(box);
    autosize(ta);
  }
  shown = end;
  const more = $('more') as HTMLButtonElement;
  more.hidden = shown >= list.length;
  more.textContent = `더 보기 (${shown}/${list.length})`;
}

function showSession(s: EditSession): void {
  session = s;
  sheet = null;
  sheetView = null;
  $('sheet').hidden = true;
  resetPdf();
  shown = 0;
  textareas.clear();
  $('paras').replaceChildren();
  $('empty').hidden = true;
  $('doc').hidden = false;
  $('toolbar').hidden = false;
  $('doc-name').textContent = s.filename;
  $('doc-format').textContent = s.engine.format.label;
  $('doc-note').textContent = s.engine.format.note;
  renderParas();
  updateCount();
  resetPreview();
  previewSource = s.engine.format.id === 'hwpx' ? () => {
    const r = s.save(); // 저장 때와 같은 검증을 거친 현재 편집본을 그린다
    if (!r.verified) throw new Error(r.problems[0] ?? '저장본 검증 실패');
    return r.bytes;
  } : null;
  $('preview').hidden = previewSource === null;
}

function openSheet(bytes: Uint8Array, name: string, info: { label: string; note: string }, edits?: [string, string][]): void {
  try {
    const s = new SheetSession(bytes, name);
    if (edits) for (const [k, t] of edits) { try { s.setText(k, t); } catch { /* 문서가 달라졌거나 값이 맞지 않으면 그 수정은 버린다 */ } }
    session = null;
    sheet = s;
    resetPdf();
    textareas.clear();
    $('paras').replaceChildren();
    ($('more') as HTMLButtonElement).hidden = true;
    $('empty').hidden = true;
    $('doc').hidden = false;
    $('toolbar').hidden = false;
    $('doc-name').textContent = name;
    $('doc-format').textContent = info.label;
    $('doc-note').textContent = info.note;
    resetPreview();
    previewSource = null;
    $('preview').hidden = true;
    $('sheet').hidden = false;
    sheetView = mountSheet(s, { onChange: () => { updateCount(); scheduleSave(); } });
    updateCount();
    setStatus(`${name}을(를) 열었습니다. 시트 ${s.doc.sheets().length}개. 칸을 눌러 값을 고치세요.`, 'ok');
  } catch (e) {
    setStatus(`열 수 없는 파일입니다: ${(e as Error).message}`, 'err');
  }
}

async function openPdf(bytes: Uint8Array, name: string, info: { label: string; note: string }, edits?: [string, string][]): Promise<void> {
  let s: PdfSession;
  try { s = new PdfSession(bytes, name); }
  catch (e) { setStatus(`열 수 없는 PDF입니다: ${(e as Error).message}`, 'err'); return; }
  if (edits) for (const [k, t] of edits) { try { s.setText(k, t); } catch { /* 맞지 않는 메모는 버린다 */ } }
  session = null; sheet = null; sheetView = null;
  resetPdf();
  textareas.clear();
  $('paras').replaceChildren();
  ($('more') as HTMLButtonElement).hidden = true;
  $('sheet').hidden = true;
  $('empty').hidden = true;
  $('doc').hidden = false;
  $('toolbar').hidden = false;
  $('doc-name').textContent = name;
  $('doc-format').textContent = info.label;
  $('doc-note').textContent = info.note;
  resetPreview();
  previewSource = null;
  $('preview').hidden = true;
  pdf = s;
  $('pdf').hidden = false;
  updateCount();
  setStatus('PDF를 불러오는 중… (처음에는 보기 엔진을 내려받습니다)');
  try {
    pdfView = await mountPdf(s, { onChange: () => { updateCount(); scheduleSave(); }, status: setStatus });
    setStatus(`${name}을(를) 열었습니다. ${s.doc.pages().length}쪽. "메모 추가"를 켜고 쪽을 눌러 메모를 남기세요.`, 'ok');
  } catch (e) {
    $('pdf').hidden = true;
    pdf = null;
    updateCount();
    setStatus(`PDF를 그리지 못했습니다: ${(e as Error).message}`, 'err');
  }
}

async function openBytes(bytes: Uint8Array, name: string, edits?: [string, string][]): Promise<void> {
  if (bytes.length > MAX_BYTES) { setStatus(`파일이 너무 큽니다(${(bytes.length / 1048576).toFixed(0)}MB). ${MAX_BYTES / 1048576}MB 이하만 엽니다.`, 'err'); return; }
  const info = detectFormat(bytes, name);
  if (info.support === 'view') { showViewOnly(bytes, name, info); return; }
  if (info.support === 'none') { setStatus(`${info.label}: ${info.note}`, 'warn'); return; }
  if (info.support === 'cell-edit') { openSheet(bytes, name, info, edits); return; }
  if (info.support === 'annotate') { await openPdf(bytes, name, info, edits); return; }
  try {
    const s = new EditSession(bytes, name);
    if (edits) for (const [id, t] of edits) { try { s.setText(id, t); } catch { /* 문서가 달라졌으면 그 수정은 버린다 */ } }
    showSession(s);
    const big = bytes.length > WARN_BYTES;
    setStatus(big
      ? `큰 파일입니다(${(bytes.length / 1048576).toFixed(0)}MB). 기기 메모리가 부족하면 브라우저가 멈출 수 있습니다.`
      : s.paragraphs.length === 0 ? '고칠 수 있는 글자가 없는 문서입니다.' : `${name}을(를) 열었습니다. 문단 ${s.paragraphs.length}개`,
    big || s.paragraphs.length === 0 ? 'warn' : 'ok');
  } catch (e) {
    if (e instanceof UnsupportedFormatError) setStatus(e.message, 'warn');
    else setStatus(`열 수 없는 파일입니다: ${(e as Error).message}`, 'err');
  }
}

function download(bytes: Uint8Array, name: string, mime: string): void {
  const blob = new Blob([bytes as BlobPart], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name; a.rel = 'noopener';
  document.body.append(a);
  a.click();
  a.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

function outName(name: string): string {
  const k = name.lastIndexOf('.');
  return k > 0 ? `${name.slice(0, k)}-수정${name.slice(k)}` : `${name}-수정`;
}

function onSave(): void {
  const a = active();
  if (!a) return;
  try {
    const r = a.save();
    const fmt = session ? session.engine.format.id : pdf ? 'pdf' : 'xlsx';
    const mime = fmt === 'pdf' ? 'application/pdf' : fmt === 'docx' ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      : fmt === 'pptx' ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      : fmt === 'xlsx' ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/hwp+zip';
    if (!r.verified) { setStatus(`저장본 검증에 실패해 내려받지 않았습니다: ${r.problems.slice(0, 3).join(' / ')}`, 'err'); return; }
    (window as unknown as { __office?: unknown }).__office = { lastSave: r.bytes, name: outName(a.filename) };
    download(r.bytes, outName(a.filename), mime);
    const what = session ? `문단 ${session.paragraphs.length}개` : pdf ? `${pdf.doc.pages().length}쪽과 모든 메모` : `시트 ${sheet!.doc.sheets().length}개의 모든 칸`;
    setStatus(`${outName(a.filename)} 저장본을 다시 열어 ${what}를 확인했습니다(변경 ${a.changeCount}). 기기의 다운로드 폴더를 확인하세요.`, 'ok');
  } catch (e) {
    setStatus(`저장하지 못했습니다: ${(e as Error).message}`, 'err');
  }
}

async function showDraftBanner(): Promise<void> {
  const d: Draft | null = await loadDraft();
  const banner = $('draft-banner');
  if (!d) { banner.hidden = true; return; }
  banner.hidden = false;
  $('draft-text').textContent = `임시 보존된 작업이 있습니다: ${d.name}, 변경 ${d.edits.length}개, ${new Date(d.savedAt).toLocaleString('ko-KR')}`;
  $('draft-resume').onclick = async () => { await openBytes(new Uint8Array(d.bytes), d.name, d.edits); banner.hidden = true; };
  $('draft-delete').onclick = async () => { await deleteDraft(); banner.hidden = true; setStatus('임시 보존을 지웠습니다.', 'ok'); };
}

function keepToolbarAboveKeyboard(): void {
  const vv = window.visualViewport;
  const tb = $('toolbar');
  if (!vv) return;
  const move = (): void => { tb.style.transform = `translateY(${-(window.innerHeight - vv.height - vv.offsetTop)}px)`; };
  vv.addEventListener('resize', move);
  vv.addEventListener('scroll', move);
}

/** 도구 모음 높이(줄바꿈으로 달라짐)만큼 본문 아래 여백을 둬서 마지막 문단이 가려지지 않게 한다. */
function trackToolbarHeight(): void {
  const tb = $('toolbar');
  const set = (): void => { document.documentElement.style.setProperty('--toolbar-h', `${tb.offsetHeight}px`); };
  if (typeof ResizeObserver !== 'undefined') new ResizeObserver(set).observe(tb);
  window.addEventListener('resize', set);
  set();
}

function init(): void {
  const sel = $('sample') as HTMLSelectElement;
  for (const [path, label] of SAMPLES) { const o = document.createElement('option'); o.value = path; o.textContent = label; sel.append(o); }
  sel.addEventListener('change', async () => {
    if (!sel.value) return;
    try {
      const res = await fetch(sel.value);
      if (!res.ok) throw new Error(`${res.status}`);
      await openBytes(new Uint8Array(await res.arrayBuffer()), sel.value.split('/').pop()!);
    } catch (e) { setStatus(`샘플을 불러오지 못했습니다: ${(e as Error).message}`, 'err'); }
    sel.value = '';
  });
  ($('file') as HTMLInputElement).addEventListener('change', async (ev) => {
    const input = ev.target as HTMLInputElement;
    const f = input.files?.[0];
    if (!f) return;
    try { await openBytes(new Uint8Array(await f.arrayBuffer()), f.name); }
    catch (e) { setStatus(`파일을 읽지 못했습니다: ${(e as Error).message}`, 'err'); }
    input.value = '';
  });
  $('more').addEventListener('click', renderParas);
  $('prev-open').addEventListener('click', () => { void onPreview(); });
  $('prev-more').addEventListener('click', renderMorePages);
  const about = $('about') as HTMLDialogElement;
  $('about-open').addEventListener('click', () => about.showModal());
  $('about-close').addEventListener('click', () => about.close());
  $('save').addEventListener('click', onSave);
  $('revert-all').addEventListener('click', () => {
    if (pdf) {
      if (!confirm('추가한 메모를 모두 지울까요?')) return;
      pdf.revertAll();
      pdfView?.refresh();
      updateCount(); scheduleSave();
      return;
    }
    if (sheet) {
      if (!confirm('모든 변경을 되돌릴까요?')) return;
      sheet.revertAll();
      sheetView?.refresh();
      updateCount(); scheduleSave();
      return;
    }
    if (!session || !confirm('모든 변경을 되돌릴까요?')) return;
    session.revertAll();
    for (const p of session.paragraphs) {
      const ta = textareas.get(p.id);
      if (!ta) continue;
      ta.value = p.text;
      const box = ta.closest('.para');
      box?.classList.remove('changed');
      box?.querySelector<HTMLElement>('.revert')?.setAttribute('hidden', '');
      autosize(ta);
    }
    updateCount(); scheduleSave();
  });
  $('clear-draft').addEventListener('click', async () => { await deleteDraft(); setStatus('임시 보존을 지웠습니다.', 'ok'); });
  $('autosave').addEventListener('change', async () => {
    if (!autosaveEnabled()) { await deleteDraft(); setStatus('임시 보존을 끄고 저장된 내용을 지웠습니다.', 'ok'); } else scheduleSave();
  });
  window.addEventListener('beforeunload', (e) => { const a = active(); if (a && a.changeCount > 0 && !autosaveEnabled()) e.preventDefault(); });
  keepToolbarAboveKeyboard();
  trackToolbarHeight();
  void showDraftBanner();
  updateCount();
}

init();

// 오프라인 지원(서비스 워커). 앱 파일만 캐시하고 사용자 문서는 캐시하지 않는다.
if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
  window.addEventListener('load', () => { navigator.serviceWorker.register('sw.js').catch(() => { /* 지원하지 않거나 막혔으면 온라인으로만 동작 */ }); });
}
