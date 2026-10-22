// XLSX 칸 격자 화면. 칸을 누르면 위쪽 입력 칸에 값이 나오고, 고친 값은 바로 검사해서 수정 목록에 넣는다.
import { SheetSession, cellKey } from '../sheet-session.ts';
import { numToCol, type CellView } from '../xlsx.ts';

const ROWS_PER_PAGE = 50;
const MAX_COLS = 40;
const KIND_LABEL: Record<string, string> = { empty: '빈 칸', number: '숫자', string: '글자', bool: '논리값', error: '오류', formula: '수식', other: '기타' };

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;

export interface SheetHooks { onChange(): void }

export function mountSheet(s: SheetSession, hooks: SheetHooks): { refresh(): void } {
  const sheets = s.doc.sheets();
  const sel = $('sheet-select') as HTMLSelectElement;
  const grid = $('grid');
  const input = $('cell-input') as HTMLTextAreaElement;
  sel.replaceChildren();
  sheets.forEach((sh, i) => { const o = document.createElement('option'); o.value = String(i); o.textContent = sh.state === 'visible' ? sh.name : `${sh.name} (숨김)`; sel.append(o); });
  let sheet = Math.max(0, sheets.findIndex(x => x.state === 'visible'));
  sel.value = String(sheet);
  let rowsShown = ROWS_PER_PAGE;
  let selected: string | null = null;

  const used = (): { rows: number; cols: number } => s.doc.usedRange(sheet);
  const colCount = (): number => Math.min(MAX_COLS, Math.max(used().cols + 1, 6));
  const rowLimit = (): number => Math.max(used().rows + 1, 10);

  function labelOf(v: CellView, text: string): string {
    return v.display !== undefined && !s.edits.has(cellKey(sheet, v.ref)) ? v.display : text;
  }

  function paintCell(td: HTMLElement, ref: string): void {
    const v = s.doc.cell(sheet, ref);
    const text = s.textOf(sheet, ref);
    td.textContent = labelOf(v, text);
    td.classList.toggle('changed', s.edits.has(cellKey(sheet, ref)));
    td.classList.toggle('num', v.kind === 'number');
    td.classList.toggle('locked', !v.editable);
  }

  function render(): void {
    grid.replaceChildren();
    const cols = colCount(), rows = Math.min(rowsShown, rowLimit());
    const head = document.createElement('tr');
    head.append(document.createElement('th'));
    for (let c = 1; c <= cols; c++) { const th = document.createElement('th'); th.textContent = numToCol(c); head.append(th); }
    grid.append(head);
    for (let r = 1; r <= rows; r++) {
      const tr = document.createElement('tr');
      const rh = document.createElement('th'); rh.textContent = String(r); tr.append(rh);
      for (let c = 1; c <= cols; c++) {
        const ref = `${numToCol(c)}${r}`;
        const td = document.createElement('td');
        td.dataset.ref = ref;
        td.tabIndex = 0;
        paintCell(td, ref);
        if (ref === selected) td.classList.add('sel');
        tr.append(td);
      }
      grid.append(tr);
    }
    $('grid-more').hidden = rows >= rowLimit();
  }

  function msg(text: string, bad = false): void { const m = $('cell-msg'); m.textContent = text; m.classList.toggle('bad', bad); }

  function select(ref: string): void {
    selected = ref;
    grid.querySelectorAll('td.sel').forEach(e => e.classList.remove('sel'));
    grid.querySelector(`td[data-ref="${ref}"]`)?.classList.add('sel');
    const v = s.doc.cell(sheet, ref);
    $('cell-ref').textContent = `${sheets[sheet].name}!${ref} · ${KIND_LABEL[v.kind] ?? v.kind}`;
    input.disabled = false;
    input.value = s.textOf(sheet, ref);
    input.readOnly = !v.editable;
    msg(v.kind === 'formula' ? `수식: =${v.formula} (수식 칸은 고칠 수 없습니다)` : v.rich ? '서식이 섞인 글자입니다. 고치면 서식이 사라집니다.' : v.display ? `날짜: ${v.display} (YYYY-MM-DD 또는 일련번호로 입력)` : '');
  }

  grid.addEventListener('click', (e) => { const td = (e.target as HTMLElement).closest('td'); if (td?.dataset.ref) select(td.dataset.ref); });
  grid.addEventListener('keydown', (e) => { if (e.key === 'Enter') { const td = (e.target as HTMLElement).closest('td'); if (td?.dataset.ref) select(td.dataset.ref); } });

  input.addEventListener('input', () => {
    if (!selected) return;
    try {
      s.setText(cellKey(sheet, selected), input.value);
      input.classList.remove('invalid');
      const v = s.doc.cell(sheet, selected);
      msg(v.display ? `날짜: ${v.display}` : '');
      const td = grid.querySelector<HTMLElement>(`td[data-ref="${selected}"]`);
      if (td) paintCell(td, selected);
      hooks.onChange();
    } catch (e) {
      input.classList.add('invalid');
      msg((e as Error).message, true);
    }
  });

  sel.addEventListener('change', () => { sheet = Number(sel.value); rowsShown = ROWS_PER_PAGE; selected = null; input.value = ''; input.disabled = true; $('cell-ref').textContent = '칸을 선택하세요'; msg(''); render(); });
  $('grid-more').onclick = () => { rowsShown += ROWS_PER_PAGE; render(); };

  input.disabled = true;
  $('cell-ref').textContent = '칸을 선택하세요';
  msg('');
  render();
  return { refresh: () => { render(); if (selected) select(selected); } };
}
