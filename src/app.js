import { Workbook, MAX_ROWS, MAX_COLS, MAX_RANGE_CELLS, FUNCTIONS, FormulaError, address, parseAddress, parseRange, normalizedRange, rangeAddress, cellsIn, keyOf, shiftFormula, formatValue, colName, DATE_FORMATS, TIME_FORMATS, CURRENCIES, numberFormatStyle } from './engine.js';
import { GridRenderer } from './renderer.js';
import { createSampleWorkbook } from './sample.js';
import { parseDelimited, serializeDelimited, exportCSV, workbookFromCSV, downloadFile, exportXLSX, importXLSX } from './io.js';
const $ = selector => document.querySelector(selector);
const $$ = selector => [...document.querySelectorAll(selector)];
const escapeHTML = s => String(s ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;');
const ICONS = {
  undo: '<path d="M8 5 3 10l5 5M3 10h10a6 6 0 0 1 0 12" transform="translate(0 -2)"/>', redo: '<path d="m16 3 5 5-5 5M21 8H11a6 6 0 0 0 0 12"/>',
  search: '<circle cx="10.5" cy="10.5" r="6.5"/><path d="m16 16 5 5"/>', edit: '<path d="m15 4 5 5M4 20l5-1L21 7a2.1 2.1 0 0 0-4-4L5 15z"/>',
  comment: '<path d="M5 4h14a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2h-9l-5 3v-3H4a2 2 0 0 1-2-2V6a2 2 0 0 1 3-2Z"/><path d="M7 9h10M7 13h7"/>',
  export: '<path d="M12 15V3m-4 4 4-4 4 4M4 13v7h16v-7"/>', paste: '<rect x="5" y="5" width="14" height="16" rx="2"/><rect x="9" y="2" width="6" height="5" rx="1"/><path d="M9 12h6m-6 4h6"/>',
  copy: '<rect x="8" y="8" width="12" height="13" rx="2"/><path d="M15 8V3H3v13h5"/>', cut: '<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="m8 8 12 12M8 16 20 4"/>',
  paint: '<path d="M4 3h16v6H4zM8 9v3h8v3h-4v6"/>', border: '<rect x="4" y="4" width="16" height="16"/><path d="M4 12h16M12 4v16" stroke-dasharray="2 2"/>', fill: '<path d="m4 11 8-8 9 9-8 8zM4 11h17M8 2l6 6M4 19c-3 4 2 4 0 0"/>',
  alignLeft: '<path d="M4 5h16M4 10h11M4 15h16M4 20h11"/>', alignCenter: '<path d="M4 5h16M7 10h10M4 15h16M7 20h10"/>', alignRight: '<path d="M4 5h16M9 10h11M4 15h16M9 20h11"/>',
  wrap: '<path d="M3 5h18M3 10h14a4 4 0 0 1 0 8h-4m2-3-3 3 3 3M3 16h5"/>', merge: '<rect x="3" y="5" width="18" height="14"/><path d="M8 9v6M16 9v6m-6-3h4"/>',
  table: '<rect x="3" y="4" width="18" height="16" rx="1"/><path d="M3 9h18M3 14h18M9 9v11M15 9v11"/>', conditional: '<rect x="3" y="4" width="18" height="16" rx="1"/><path d="M3 10h18M9 4v16M15 4v16"/><path d="M4 15h4" stroke="#d8ad46" stroke-width="3"/>',
  insert: '<rect x="3" y="8" width="18" height="12"/><path d="M3 14h18M9 8v12M15 8v12M8 3h8m-4-3v6"/>', delete: '<rect x="3" y="8" width="18" height="12"/><path d="M3 14h18M9 8v12M15 8v12M8 3h8"/>',
  sort: '<path d="M8 4v16m-4-4 4 4 4-4M15 5h6m-6 5h4m-4 5h2"/>', filter: '<path d="M3 4h18l-7 8v7l-4 2v-9z"/>', grid: '<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/>',
  moon: '<path d="M20 15A9 9 0 0 1 9 4a9 9 0 1 0 11 11Z"/>', list: '<path d="M7 5h14M7 12h14M7 19h14M3 5h.1M3 12h.1M3 19h.1"/>',
  chart: '<path d="M3 3v18h18M7 17V9h3v8M13 17V4h3v13M19 17v-6h3v6"/>', line: '<path d="M3 3v18h18M6 16l5-7 5 3 5-8"/>', donut: '<path d="M12 3a9 9 0 1 0 9 9h-9Z"/><path d="M15 2v7h7a8 8 0 0 0-7-7Z"/>',
  file: '<path d="M5 2h9l5 5v15H5zM14 2v6h5M8 12h8M8 16h8"/>', open: '<path d="M3 7V4h7l3 3h8v12H3V7m0 4h18"/>', save: '<path d="M3 3h15l3 3v15H3zM7 3v6h9V3M7 21v-8h10v8"/>',
  plus: '<path d="M12 4v16M4 12h16"/>', freeze: '<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M3 9h18M9 3v18" stroke-width="2.5"/>', print: '<path d="M7 8V3h10v5M7 17H3V8h18v9h-4M7 13h10v8H7zM17 11h1"/>',
  sum: '<path d="M19 4H5l7 8-7 8h14"/>', function: '<path d="M17 4c-4-4-6 1-7 7l-2 7c-1 4-4 4-5 1M6 10h10m1 4 5 7m0-7-5 7"/>',
  check: '<path d="m4 12 5 5L20 6"/>', clear: '<path d="m9 3 12 10-8 8H7l-6-6ZM5 11l10 9"/>', lock: '<rect x="4" y="10" width="16" height="12" rx="2"/><path d="M8 10V6a4 4 0 0 1 8 0v4M12 15v3"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7h.01"/>', refresh: '<path d="M20 5v6h-6M4 19v-6h6M20 11a8 8 0 0 0-14-6M4 13a8 8 0 0 0 14 6"/>',
  name: '<path d="M3 5h18v14H3zM7 15l3-7 3 7m-5-2h4m4-5v7"/>', percent: '<circle cx="6" cy="6" r="3"/><circle cx="18" cy="18" r="3"/><path d="m4 20 16-16"/>'
};
function icon(name, large = false) { return `<svg class="icon${large ? ' large' : ''}" viewBox="0 0 24 24" aria-hidden="true">${ICONS[name] || ICONS.grid}</svg>`; }
function hydrateIcons(root = document) { root.querySelectorAll('[data-icon]').forEach(el => { el.innerHTML = icon(el.dataset.icon); }); }
const tool = (action, label, image, large = false, more = false) => `<button class="${large ? 'big-tool' : 'small-tool'}" data-action="${action}" title="${escapeHTML(label)}">${icon(image, large)}<span>${label}</span>${more ? '<span class="chevron">⌄</span>' : ''}</button>`;
const mini = (action, image, title, content = '') => `<button class="mini-tool" data-action="${action}" title="${escapeHTML(title)}" aria-label="${escapeHTML(title)}">${content || icon(image)}</button>`;
const group = (label, content, extra = '') => `<div class="ribbon-group ${extra}"><div class="group-content">${content}</div><div class="group-label">${label}</div></div>`;
const STORAGE_KEY = 'gridline.workbook.v1';
const REOPEN_KEY = 'gridline.reopen-last.v1';
const COMMANDS = [
  ['new','New workbook','file','Ctrl/⌘ N'],['open','Open workbook','open','Ctrl/⌘ O'],['save','Save Gridline workbook','save','Ctrl/⌘ S'],['export-xlsx','Export Excel workbook (.xlsx)','export',''],['export-csv','Export current sheet as CSV','export',''],
  ['hide-columns','Hide columns','table',''],['unhide-columns','Unhide columns','table',''],['unhide-all-columns','Unhide all columns','table',''],['text-to-columns','Text to Columns','table',''],['format-cells','Format cells…','table','Ctrl/⌘ 1'],['find','Find and replace','search','Ctrl/⌘ F'],['chart','Insert chart','chart',''],['functions','Insert function','function',''],['name-manager','Named ranges','name',''],['sort','Sort range','sort',''],['filter','Filter values','filter',''],['conditional','Conditional formatting','conditional',''],['format-table','Format as table','table',''],['freeze-top','Freeze top row','freeze',''],['freeze-first','Freeze first column','freeze',''],['freeze','Freeze at active cell','freeze',''],['unfreeze','Unfreeze panes','freeze',''],['toggle-gridlines','Toggle gridlines','grid',''],['show-formulas','Show formulas','function','Ctrl/⌘ `'],['add-note','Add a cell note','comment',''],['notes','View notes','comment',''],['insert-row','Insert row','insert',''],['insert-column','Insert column','insert',''],['delete-row','Delete row','delete',''],['delete-column','Delete column','delete',''],['add-sheet','Add worksheet','plus',''],['duplicate-sheet','Duplicate worksheet','copy',''],['theme','Toggle dark mode','moon',''],['recalculate','Recalculate workbook','refresh',''],['performance','Renderer diagnostics','grid',''],['print','Print worksheet','print','Ctrl/⌘ P'],['help','Keyboard shortcuts','info','F1']
];
class GridlineApp {
  constructor() {
    this.active = { r: 12, c: 6 }; this.anchor = { ...this.active }; this.selection = normalizedRange(this.active); this.tab = 'Home'; this.autosave = true; this.clipboard = null; this.editing = false; this.barEditing = false; this.sheetViews = new Map(); this.chartRevision = -1; this.panelType = null;
    this.host = $('#grid-host'); this.editor = $('#cell-editor'); this.formulaInput = $('#formula-input'); this.dialog = $('#dialog'); this.drag = null; this.composing = false;
    let workbook; const params = new URLSearchParams(location.search);
    this.reopenLast = false;
    try { this.reopenLast = localStorage.getItem(REOPEN_KEY) === 'true'; const saved = this.reopenLast && !params.has('fresh') && !params.has('open-file') && localStorage.getItem(STORAGE_KEY); if (saved) workbook = Workbook.fromJSON(JSON.parse(saved)); } catch (e) { console.warn('Gridline restore:', e); }
    this.startupPending = !workbook; this.workbook = workbook || new Workbook();
    this.active = this.anchor = {r:0,c:0};
    this.renderer = new GridRenderer(this.host, this.workbook, (backend, reason) => { $('#backend-label').textContent = backend === 'webgpu' ? 'WebGPU accelerated' : 'Canvas2D fallback'; $('#renderer-status').title = reason || 'GPU-instanced grid and glyph atlas'; });
    this.renderer.onFrame = () => { this.positionEditor(); this.positionCharts(); this.updateScrollbars(); if (this.panelType === 'performance') this.refreshPerformance(); };
    this.observeWorkbook(); this.bindEvents(); this.renderRibbon(); this.renderTabs(); this.updateUI(); hydrateIcons(); this.renderCharts();
    this.renderer.initialize(params.get('renderer') === 'canvas').then(() => { $('#loading').remove(); this.select(this.selection, this.active, false); this.host.focus(); this.ready = true; window.launchQueue?.setConsumer(params => { if (params.files?.length) this.errorBoundary(async () => { await this.openFile(await params.files[0].getFile()); this.closeDialog(); }); }); if (!workbook && !params.has('open-file')) this.showFile(); window.dispatchEvent(new Event('gridline-ready')); });
  }
  get sheet() { return this.workbook.activeSheet; }
  observeWorkbook() {
    this.unsubscribe?.(); this.unsubscribe = this.workbook.onChange(() => {
      this.startupPending = false;
      this.renderer.workbook = this.workbook; this.renderer.syncLayout(); this.updateUI(); this.renderTabs(); this.renderCharts();
      if (this.panelType === 'notes') this.showNotes();
      $('#save-status').textContent = this.autosave ? 'Saving locally…' : 'Autosave is off';
      clearTimeout(this.saveTimer); if (this.autosave) this.saveTimer = setTimeout(() => this.persist(), 650);
    });
  }
  persist() {
    if (this.startupPending) return true;
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(this.workbook.toJSON())); $('#save-status').textContent = 'Saved on this device'; return true; }
    catch { $('#save-status').textContent = 'Local storage full — export to save'; this.toast('Local storage is unavailable or full. Export a .gridline file to save your workbook.', true); return false; }
  }
  setWorkbook(workbook) {
    this.commitEdit(false); this.startupPending = false; this.workbook = workbook; this.renderer.workbook = workbook; this.renderer.scrollX = this.renderer.scrollY = 0; this.sheetViews.clear(); this.clipboard = null; this.renderer.copyRange = null; this.observeWorkbook(); this.renderer.syncLayout(); this.select(normalizedRange({ r: 0, c: 0 }), { r: 0, c: 0 }, false); this.closePanel(); this.renderTabs(); this.renderCharts(); this.updateUI(); this.persist(); this.host.focus();
  }
  updateUI() {
    $('#workbook-title').value = this.workbook.title; document.title = `${this.workbook.title} — Gridline`;
    $('#name-box').value = rangeAddress(this.selection);
    if (!this.editing && !this.barEditing) this.formulaInput.value = this.workbook.editValue(this.sheet, this.active.r, this.active.c);
    this.renderer.selection = this.selection; this.renderer.active = this.active;
    const style = this.sheet.style(this.active.r, this.active.c);
    for (const prop of ['bold', 'italic', 'underline', 'wrap']) $$(`[data-action="${prop}"]`).forEach(b => b.classList.toggle('active', !!style[prop]));
    for (const align of ['left', 'center', 'right']) $$(`[data-action="align-${align}"]`).forEach(b => b.classList.toggle('active', style.align === align));
    const format = $('#number-format'); if (format) format.value = ['general','number','currency','percent','date','integer'].includes(style.format) ? style.format : 'general';
    const size = $('#font-size'); if (size) size.value = Math.round((style.fontSize || 13) * 0.75);
    $$('[data-action="undo"]').forEach(b => b.disabled = !this.workbook.undoStack.length); $$('[data-action="redo"]').forEach(b => b.disabled = !this.workbook.redoStack.length);
    const text = this.workbook.display(this.sheet, this.active.r, this.active.c), formula = this.sheet.raw(this.active.r, this.active.c);
    $('#active-cell-description').textContent = `${this.sheet.name}. ${address(this.active.r, this.active.c)}. ${text}${formula.startsWith('=') ? '. Formula ' + formula : ''}`;
    $('#active-cell-description').setAttribute('aria-rowindex', String(this.active.r + 1)); $('#active-cell-description').setAttribute('aria-colindex', String(this.active.c + 1));
    $('#mode-status').textContent = this.sheet.protected ? 'Read-only sheet' : this.editing || this.barEditing ? 'Edit' : 'Ready';
    let count = 0, numbers = 0, total = 0;
    for (const [key, cell] of this.sheet.cells) {
      const [r, c] = key.split(',').map(Number); const q = this.selection;
      if (r >= q.r1 && r <= q.r2 && c >= q.c1 && c <= q.c2 && cell.raw) { const v = this.workbook.value(this.sheet, r, c); count++; if (typeof v === 'number') { total += v; numbers++; } }
    }
    $('#selection-stats').innerHTML = count > 1 ? `${numbers ? `<span>Average: <b>${escapeHTML(formatValue(total / numbers, { format: 'number' }))}</b></span>` : ''}<span>Count: <b>${count.toLocaleString()}</b></span>${numbers ? `<span>Sum: <b>${escapeHTML(formatValue(total, { format: 'number' }))}</b></span>` : ''}` : '';
    this.renderer.requestFrame();
  }
  select(q, active = { r: q.r1, c: q.c1 }, reveal = false) {
    q = { r1: Math.max(0, Math.min(MAX_ROWS - 1, q.r1)), r2: Math.max(0, Math.min(MAX_ROWS - 1, q.r2)), c1: Math.max(0, Math.min(MAX_COLS - 1, q.c1)), c2: Math.max(0, Math.min(MAX_COLS - 1, q.c2)) };
    const merged = q.r1 === q.r2 && q.c1 === q.c2 && this.sheet.mergeAt(q.r1, q.c1); if (merged) q = { ...merged };
    this.selection = q; this.active = { r: Math.max(0, Math.min(MAX_ROWS - 1, active.r)), c: Math.max(0, Math.min(MAX_COLS - 1, active.c)) }; this.updateUI(); if (reveal) this.renderer.ensureVisible(this.active.r, this.active.c);
  }
  goto(r, c, extend = false) { r = Math.max(0, Math.min(MAX_ROWS - 1, r)); c = Math.max(0, Math.min(MAX_COLS - 1, c)); const p = { r, c }; if (!extend) this.anchor = p; this.select(extend ? normalizedRange(this.anchor, p) : normalizedRange(p), p, true); }
  errorBoundary(fn) { try { const value = fn(); if (value instanceof Promise) value.catch(e => this.toast(e.message, true)); return value; } catch (e) { this.toast(e.message, true); console.error(e); } }
  toast(message, error = false) { const t = $('#toast'); t.textContent = message; t.hidden = false; t.classList.toggle('error', error); clearTimeout(this.toastTimer); this.toastTimer = setTimeout(() => t.hidden = true, error ? 6500 : 3500); }
  editable() { if (this.sheet.protected) { this.toast('This sheet is read-only. Turn off sheet protection from Review to edit.', true); return false; } return true; }
  bindEvents() {
    document.addEventListener('click', e => {
      const tab = e.target.closest('[data-tab]'); if (tab) { this.tab = tab.dataset.tab; this.renderRibbon(); return; }
      const action = e.target.closest('[data-action]'); if (action && !action.disabled) this.errorBoundary(() => this.run(action.dataset.action, action));
      const sheetTab = e.target.closest('[data-sheet]'); if (sheetTab) this.switchSheet(sheetTab.dataset.sheet);
    });
    document.addEventListener('pointerdown', e => {
      if (!e.target.closest('#context-menu')) $('#context-menu').hidden = true;
      if (this.editing && !e.target.closest('#grid-host') && !e.target.closest('#formula-input') && !e.target.closest('.formula-control')) this.errorBoundary(() => this.commitEdit(false));
      if (e.target.closest('.ribbon button,.formula-control,.fx') && !e.target.closest('select,input')) e.preventDefault();
    });
    document.addEventListener('keydown', e => this.errorBoundary(() => this.onKey(e)));
    this.host.addEventListener('pointerdown', e => this.errorBoundary(() => this.onPointerDown(e)));
    this.host.addEventListener('pointermove', e => this.onPointerMove(e));
    this.host.addEventListener('pointerup', e => this.onPointerUp(e));
    this.host.addEventListener('pointercancel', e => this.onPointerUp(e));
    this.host.addEventListener('dblclick', e => { if (e.target.closest('.chart-card,.scroll-thumb,.cell-editor')) return; const p = this.localPoint(e), hit = this.renderer.hitTest(p.x, p.y); if (hit.colHeader) this.autoFit(this.selection.r1 === 0 && this.selection.r2 === MAX_ROWS - 1 && hit.c >= this.selection.c1 && hit.c <= this.selection.c2 ? null : hit.c); else if (!hit.rowHeader) this.startEdit(); });
    this.host.addEventListener('contextmenu', e => {
      if (e.target.closest('.cell-editor')) return; e.preventDefault(); this.errorBoundary(() => {
        this.commitEdit(false); this.commitFormula();
        const p = this.localPoint(e), hit = this.renderer.hitTest(p.x, p.y), q = this.selection;
        if (hit.colHeader || hit.rowHeader) {
          const selected = hit.colHeader && !hit.rowHeader ? q.r1 === 0 && q.r2 === MAX_ROWS - 1 && hit.c >= q.c1 && hit.c <= q.c2 : hit.rowHeader && !hit.colHeader && q.c1 === 0 && q.c2 === MAX_COLS - 1 && hit.r >= q.r1 && hit.r <= q.r2;
          if (!selected) this.select({ r1: hit.colHeader ? 0 : hit.r, r2: hit.colHeader ? MAX_ROWS - 1 : hit.r, c1: hit.rowHeader ? 0 : hit.c, c2: hit.rowHeader ? MAX_COLS - 1 : hit.c }, { r: hit.colHeader ? 0 : hit.r, c: hit.rowHeader ? 0 : hit.c });
        } else if (hit.r < q.r1 || hit.r > q.r2 || hit.c < q.c1 || hit.c > q.c2) this.goto(hit.r, hit.c);
        this.cellContextMenu(e.clientX, e.clientY, hit.colHeader && !hit.rowHeader ? 'column' : hit.rowHeader && !hit.colHeader ? 'row' : null);
      });
    });
    this.host.addEventListener('wheel', e => this.errorBoundary(() => { e.preventDefault(); if (this.editing) this.commitEdit(false); if (e.ctrlKey || e.metaKey) this.setZoom(this.renderer.zoom + (e.deltaY > 0 ? -.1 : .1)); else { const unit = e.deltaMode === 1 ? 25 : e.deltaMode === 2 ? this.renderer.height : 1; this.renderer.scrollY += e.shiftKey ? 0 : e.deltaY * unit; this.renderer.scrollX += (e.shiftKey ? e.deltaX || e.deltaY : e.deltaX) * unit; this.renderer.clampScroll(); this.renderer.requestFrame(); } }), { passive: false });
    this.editor.addEventListener('compositionstart', () => this.composing = true); this.editor.addEventListener('compositionend', () => this.composing = false);
    this.editor.addEventListener('input', () => { this.formulaInput.value = this.editor.value; this.showFormulaSuggestions(this.editor); });
    this.formulaInput.addEventListener('focus', () => { if (this.editing) { this.formulaInput.value = this.editor.value; this.editing = false; this.editor.hidden = true; } this.barEditing = true; $('#mode-status').textContent = 'Edit'; });
    this.formulaInput.addEventListener('input', () => this.showFormulaSuggestions(this.formulaInput));
    this.formulaInput.addEventListener('blur', () => { setTimeout(() => { if (this.barEditing && document.activeElement !== this.formulaInput && !this.drag?.reference) this.errorBoundary(() => this.commitFormula()); }, 0); });
    $('#name-box').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); const value = e.target.value.trim(), named = this.workbook.names[value.toUpperCase()]; let q = parseRange(value); if (!q && named) { const parts = named.split('!'); const name = parts[0].replace(/^'|'$/g, '').replaceAll("''", "'"); const sheet = this.workbook.sheetByName(name); if (sheet) { this.switchSheet(sheet.id); q = parseRange(parts[1]); } } if (q) { this.anchor = { r: q.r1, c: q.c1 }; this.select(q, this.anchor, true); this.host.focus(); } else this.toast('Enter a cell or range, for example A1, B2:F20, or a defined name.', true); } });
    $('#workbook-title').addEventListener('change', e => this.workbook.mutate('Rename workbook', () => this.workbook.title = e.target.value.trim().slice(0, 200) || 'Untitled workbook'));
    $('#zoom-slider').addEventListener('input', e => this.setZoom(+e.target.value / 100));
    $('#file-input').addEventListener('change', e => { const file = e.target.files[0]; if (file) this.errorBoundary(() => this.openFile(file)); e.target.value = ''; });
    this.host.addEventListener('dragover', e => { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; });
    this.host.addEventListener('drop', e => { e.preventDefault(); const file = e.dataTransfer.files[0]; if (file) this.errorBoundary(() => this.openFile(file)); });
    document.addEventListener('copy', e => { if (this.isInputFocus()) return; this.copyToEvent(e, false); });
    document.addEventListener('cut', e => { if (this.isInputFocus()) return; this.copyToEvent(e, true); });
    document.addEventListener('paste', e => { if (this.isInputFocus() || this.dialog.open) return; e.preventDefault(); this.errorBoundary(() => this.pasteText(e.clipboardData.getData('text/plain'))); });
    $('#sheet-tabs').addEventListener('dblclick', e => { const t = e.target.closest('[data-sheet]'); if (t) { this.switchSheet(t.dataset.sheet); this.renameSheet(); } });
    $('#sheet-tabs').addEventListener('contextmenu', e => { const t = e.target.closest('[data-sheet]'); if (t) { e.preventDefault(); this.switchSheet(t.dataset.sheet); this.contextMenu(e.clientX, e.clientY, [['rename-sheet','Rename…'],['duplicate-sheet','Duplicate'],['add-sheet','Insert worksheet'],null,['delete-sheet','Delete worksheet']]); } });
    for (const axis of ['v', 'h']) $(`#${axis}-scrollbar`).addEventListener('pointerdown', e => this.scrollbarDown(e, axis));
    this.dialog.addEventListener('click', e => { if (e.target === this.dialog) { const b = this.dialog.getBoundingClientRect(); if (e.clientX < b.left || e.clientX > b.right || e.clientY < b.top || e.clientY > b.bottom) this.closeDialog(); } });
    this.dialog.addEventListener('close', () => { if (!this.isInputFocus()) this.host.focus(); });
    window.addEventListener('beforeunload', () => { if (this.autosave) this.persist(); });
    window.addEventListener('error', e => { $('#loading')?.remove(); this.toast(`Application error: ${e.message}`, true); });
  }
  isInputFocus() { return ['INPUT','TEXTAREA','SELECT'].includes(document.activeElement?.tagName) || document.activeElement?.isContentEditable; }
  localPoint(e) { const rect = this.host.getBoundingClientRect(); return { x: e.clientX - rect.left, y: e.clientY - rect.top }; }
  onPointerDown(e) {
    if (e.button !== 0 || e.ctrlKey || e.target.closest('.chart-card,.v-scrollbar,.h-scrollbar,.cell-editor,.formula-suggestions')) return;
    const p = this.localPoint(e), hit = this.renderer.hitTest(p.x, p.y), refInput = this.editing ? this.editor : this.barEditing ? this.formulaInput : null;
    if (refInput && refInput.value.startsWith('=') && /[=+\-*/^(,:<> ]$/.test(refInput.value.slice(0, refInput.selectionStart))) {
      e.preventDefault(); const at = refInput.selectionStart; const text = address(hit.r, hit.c); refInput.setRangeText(text, at, refInput.selectionEnd, 'end'); this.formulaInput.value = refInput.value;
      this.drag = { mode: 'reference', reference: true, input: refInput, start: at, length: text.length, anchor: hit }; this.host.setPointerCapture(e.pointerId); return;
    }
    e.preventDefault(); this.commitEdit(false); this.commitFormula(); this.host.focus();
    const colRect = this.renderer.cellRect(0, hit.c, false), rowRect = this.renderer.cellRect(hit.r, 0, false);
    if (hit.colHeader && !hit.rowHeader && Math.abs(p.x - colRect.x - colRect.w) < 6 && this.editable()) this.drag = { mode: 'resize-col', index: hit.c, start: p.x, size: this.renderer.cols.size(hit.c) };
    else if (hit.rowHeader && !hit.colHeader && Math.abs(p.y - rowRect.y - rowRect.h) < 6 && this.editable()) this.drag = { mode: 'resize-row', index: hit.r, start: p.y, size: this.renderer.rows.size(hit.r) };
    else {
      const end = this.renderer.cellRect(this.selection.r2, this.selection.c2, false);
      if (!hit.colHeader && !hit.rowHeader && Math.abs(p.x - end.x - end.w) < 7 && Math.abs(p.y - end.y - end.h) < 7 && this.editable()) this.drag = { mode: 'fill', source: { ...this.selection } };
      else {
        if (!e.shiftKey) this.anchor = { r: hit.r, c: hit.c };
        let q = normalizedRange(this.anchor, hit);
        if (hit.colHeader && hit.rowHeader) q = { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 };
        else if (hit.colHeader) { q.r1 = 0; q.r2 = MAX_ROWS - 1; }
        else if (hit.rowHeader) { q.c1 = 0; q.c2 = MAX_COLS - 1; }
        this.select(q, { r: hit.r, c: hit.c }); this.drag = { mode: 'select', rowHeader: hit.rowHeader, colHeader: hit.colHeader };
        if (this.paintStyle) { this.format(this.paintStyle); this.paintStyle = null; this.toast('Formatting applied.'); }
      }
    }
    this.host.setPointerCapture(e.pointerId);
  }
  onPointerMove(e) {
    const p = this.localPoint(e), hit = this.renderer.hitTest(p.x, p.y);
    if (!this.drag) {
      const rect = this.renderer.cellRect(hit.r, hit.c, false), end = this.renderer.cellRect(this.selection.r2, this.selection.c2, false);
      this.host.style.cursor = hit.colHeader && Math.abs(p.x - rect.x - rect.w) < 6 ? 'col-resize' : hit.rowHeader && Math.abs(p.y - rect.y - rect.h) < 6 ? 'row-resize' : Math.abs(p.x - end.x - end.w) < 7 && Math.abs(p.y - end.y - end.h) < 7 ? 'crosshair' : 'cell'; return;
    }
    const drag = this.drag;
    if (drag.mode === 'reference') {
      const q = normalizedRange(drag.anchor, hit), text = rangeAddress(q); drag.input.setRangeText(text, drag.start, drag.start + drag.length, 'end'); drag.length = text.length; this.formulaInput.value = drag.input.value; return;
    }
    if (drag.mode === 'resize-col' || drag.mode === 'resize-row') {
      const horizontal = drag.mode === 'resize-col', delta = (horizontal ? p.x : p.y) - drag.start;
      drag.next = Math.max(horizontal ? 28 : 18, Math.min(horizontal ? 1000 : 400, drag.size + delta / this.renderer.zoom));
      const sizes = horizontal ? 'colWidths' : 'rowHeights';
      if (drag.old === undefined) drag.old = new Map(this.sheet[sizes]);
      this.sheet[sizes].set(drag.index, drag.next); this.renderer.syncLayout(); return;
    }
    if (p.y > this.renderer.height - 22) this.renderer.scrollY += 16; else if (p.y < this.renderer.headerH + 8) this.renderer.scrollY -= 16;
    if (p.x > this.renderer.width - 22) this.renderer.scrollX += 16; else if (p.x < this.renderer.headerW + 8) this.renderer.scrollX -= 16;
    this.renderer.clampScroll();
    if (drag.mode === 'fill') { const q = { r1: Math.min(drag.source.r1, hit.r), c1: Math.min(drag.source.c1, hit.c), r2: Math.max(drag.source.r2, hit.r), c2: Math.max(drag.source.c2, hit.c) }; drag.target = q; this.renderer.fillPreview = q; this.renderer.requestFrame(); }
    else { let q = normalizedRange(this.anchor, hit); if (drag.colHeader) { q.r1 = 0; q.r2 = MAX_ROWS - 1; } if (drag.rowHeader) { q.c1 = 0; q.c2 = MAX_COLS - 1; } this.select(q, this.anchor); }
  }
  onPointerUp(e) {
    const drag = this.drag; this.drag = null; if (this.host.hasPointerCapture(e.pointerId)) this.host.releasePointerCapture(e.pointerId); if (!drag) return;
    if (drag.mode === 'reference') { drag.input.focus(); return; }
    if (drag.mode.startsWith('resize') && drag.next !== undefined) {
      const sizes = drag.mode === 'resize-col' ? 'colWidths' : 'rowHeights'; this.sheet[sizes] = drag.old;
      this.workbook.mutate('Resize ' + (sizes === 'colWidths' ? 'column' : 'row'), () => this.sheet[sizes].set(drag.index, drag.next));
    }
    if (drag.mode === 'fill') { this.renderer.fillPreview = null; if (drag.target) this.errorBoundary(() => { this.workbook.fill(this.sheet, drag.source, drag.target); this.select(drag.target); }); this.renderer.requestFrame(); }
  }
  scrollbarDown(e, axis) {
    e.preventDefault(); e.stopPropagation(); const track = e.currentTarget, thumb = $(`#${axis}-thumb`), rect = track.getBoundingClientRect(), vertical = axis === 'v';
    const length = vertical ? rect.height : rect.width, thumbLength = vertical ? thumb.offsetHeight : thumb.offsetWidth, max = vertical ? Math.max(0, this.renderer.rows.offset(this.renderer.populatedRows) * this.renderer.zoom - this.renderer.height + this.renderer.headerH) : this.renderer.cols.offset(MAX_COLS) * this.renderer.zoom - this.renderer.width + this.renderer.headerW;
    const grab = e.target === thumb ? (vertical ? e.clientY - thumb.getBoundingClientRect().top : e.clientX - thumb.getBoundingClientRect().left) : thumbLength / 2;
    const move = event => { const pos = (vertical ? event.clientY - rect.top : event.clientX - rect.left) - grab; this.renderer[vertical ? 'scrollY' : 'scrollX'] = Math.max(0, Math.min(1, pos / Math.max(1, length - thumbLength))) * max; this.renderer.requestFrame(); };
    track.setPointerCapture(e.pointerId); if (e.target !== thumb) move(e);
    const up = () => { track.removeEventListener('pointermove', move); track.removeEventListener('pointerup', up); track.removeEventListener('pointercancel', up); };
    track.addEventListener('pointermove', move); track.addEventListener('pointerup', up); track.addEventListener('pointercancel', up);
  }
  updateScrollbars() {
    for (const axis of ['v','h']) {
      const vertical = axis === 'v', track = $(`#${axis}-scrollbar`), thumb = $(`#${axis}-thumb`), length = vertical ? track.clientHeight : track.clientWidth;
      const total = (vertical ? this.renderer.rows.offset(this.renderer.populatedRows) : this.renderer.cols.offset(MAX_COLS)) * this.renderer.zoom, view = vertical ? this.renderer.height - this.renderer.headerH : this.renderer.width - this.renderer.headerW, scroll = vertical ? this.renderer.scrollY : this.renderer.scrollX;
      const size = Math.min(length, Math.max(28, length * view / total)), position = (length - size) * scroll / Math.max(1, total - view);
      thumb.style[vertical ? 'height' : 'width'] = size + 'px'; thumb.style[vertical ? 'top' : 'left'] = Math.max(0, Math.min(length - size, position)) + 'px';
    }
  }
  startEdit(initial) {
    if (!this.editable()) return;
    this.editing = true; this.barEditing = false; this.editor.hidden = false; this.editor.value = initial === undefined ? this.workbook.editValue(this.sheet, this.active.r, this.active.c) : initial; this.formulaInput.value = this.editor.value;
    const style = this.sheet.style(this.active.r, this.active.c); this.editor.style.fontFamily = style.fontFamily || 'Aptos, "Segoe UI", Arial, sans-serif'; this.editor.style.fontSize = (style.fontSize || 13) * this.renderer.zoom + 'px'; this.editor.style.fontWeight = style.bold ? '600' : '400';
    this.positionEditor(); this.editor.focus(); this.editor.setSelectionRange(this.editor.value.length, this.editor.value.length); $('#mode-status').textContent = 'Edit';
  }
  positionEditor() {
    if (!this.editing) return; const rect = this.renderer.cellRect(this.active.r, this.active.c);
    this.editor.style.left = rect.x + 'px'; this.editor.style.top = rect.y + 'px'; this.editor.style.width = Math.min(Math.max(rect.w + 1, 150), this.renderer.width - rect.x - 12) + 'px'; this.editor.style.height = Math.max(rect.h + 1, 29) + 'px';
  }
  commitEdit(move = false) {
    if (!this.editing) return; const raw = this.editor.value;
    if (raw !== this.workbook.editValue(this.sheet, this.active.r, this.active.c)) this.workbook.setRaw(this.sheet, this.active.r, this.active.c, raw);
    this.editing = false; this.editor.hidden = true; $('#formula-suggestions').hidden = true;
    this.updateUI(); if (move) this.goto(this.active.r + 1, this.active.c);
  }
  commitFormula() {
    if (!this.barEditing) return;
    if (this.editable() && this.formulaInput.value !== this.workbook.editValue(this.sheet, this.active.r, this.active.c)) this.workbook.setRaw(this.sheet, this.active.r, this.active.c, this.formulaInput.value);
    this.barEditing = false; $('#formula-suggestions').hidden = true; this.updateUI();
  }
  cancelEdit() { this.editing = this.barEditing = false; this.editor.hidden = true; $('#formula-suggestions').hidden = true; this.updateUI(); this.host.focus(); }
  showFormulaSuggestions(input) {
    const list = $('#formula-suggestions'), before = input.value.slice(0, input.selectionStart), match = input.value.startsWith('=') && /(?:^=|[+\-*/(,])([A-Z][A-Z0-9.]*)$/i.exec(before);
    if (!match || match[1].length < 1) { list.hidden = true; return; }
    const names = [...FUNCTIONS.keys()].filter(n => n.startsWith(match[1].toUpperCase())).slice(0, 7); if (!names.length) { list.hidden = true; return; }
    list.innerHTML = names.map(n => `<button data-completion="${n}"><span>${n}</span><span>${escapeHTML(FUNCTIONS.get(n).description)}</span></button>`).join('');
    const rect = this.renderer.cellRect(this.active.r, this.active.c); list.style.left = Math.max(46, Math.min(rect.x, this.renderer.width - 380)) + 'px'; list.style.top = (input === this.formulaInput ? 0 : rect.y + rect.h + 3) + 'px'; list.hidden = false;
    list.onpointerdown = e => { const btn = e.target.closest('[data-completion]'); if (!btn) return; e.preventDefault(); e.stopPropagation(); const end = input.selectionStart; input.setRangeText(btn.dataset.completion + '(', end - match[1].length, end, 'end'); this.formulaInput.value = input.value; list.hidden = true; input.focus(); };
  }
  onKey(e) {
    if (this.composing || e.isComposing) return;
    const mod = e.ctrlKey || e.metaKey, key = e.key.toLowerCase();
    if (e.key === 'Escape') { $('#context-menu').hidden = true; if (this.dialog.open) { this.closeDialog(); return; } this.cancelEdit(); this.renderer.copyRange = null; this.renderer.requestFrame(); return; }
    if (this.dialog.open) return;
    if (mod && ['s','o','p','k'].includes(key)) { e.preventDefault(); this.errorBoundary(() => this.run({ s:'save', o:'open', p:'print', k:'commands' }[key])); return; }
    if (this.editing && !mod && !e.shiftKey && !e.altKey && ['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key)) { e.preventDefault(); this.commitEdit(false); this.host.focus(); }
    if (this.editing || this.barEditing) {
      if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); if (this.editing) this.commitEdit(false); else this.commitFormula(); this.host.focus(); this.goto(this.active.r + (e.shiftKey ? -1 : 1), this.active.c); }
      else if (e.key === 'Tab') { e.preventDefault(); if (this.editing) this.commitEdit(false); else this.commitFormula(); this.host.focus(); this.goto(this.active.r, this.active.c + (e.shiftKey ? -1 : 1)); }
      else if (e.key === 'F4') { e.preventDefault(); const input = this.editing ? this.editor : this.formulaInput; const at = input.selectionStart, left = input.value.slice(0, at), m = /\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(left); if (m) { const text = m[0], next = !text.includes('$') ? `$${m[1]}$${m[2]}` : text.startsWith('$') && /\$\d/.test(text) ? `${m[1]}$${m[2]}` : !text.startsWith('$') ? `$${m[1]}${m[2]}` : `${m[1]}${m[2]}`; input.setRangeText(next, at - text.length, at, 'end'); this.formulaInput.value = input.value; } }
      return;
    }
    if (this.isInputFocus()) return;
    if (mod && key === '1') { e.preventDefault(); this.errorBoundary(() => this.run('format-cells')); return; }
    if (mod && ['z','y','f','h','b','i','u','a','d','r','`','n'].includes(key)) {
      e.preventDefault();
      const command = { z:e.shiftKey ? 'redo' : 'undo', y:'redo', f:'find', h:'find', b:'bold', i:'italic', u:'underline', a:'select-all', d:'fill-down', r:'fill-right', '`':'show-formulas', n:'new' }[key]; this.errorBoundary(() => this.run(command)); return;
    }
    if (e.key === 'F1') { e.preventDefault(); this.run('help'); return; }
    if (e.key === 'F2') { e.preventDefault(); this.startEdit(); return; }
    if (e.altKey && e.key === '=') { e.preventDefault(); this.run('autosum'); return; }
    if (['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Tab','Enter','Home','End','PageUp','PageDown'].includes(e.key)) {
      e.preventDefault(); let { r, c } = this.active;
      if (e.key === 'Home') { c = 0; if (mod) r = 0; }
      else if (e.key === 'End') { const used = this.sheet.usedRange(); c = used.c2; if (mod) r = used.r2; }
      else if (e.key === 'PageDown' || e.key === 'PageUp') r += Math.max(1, Math.floor((this.renderer.height - 27) / (27 * this.renderer.zoom)) - 1) * (e.key === 'PageDown' ? 1 : -1);
      else { const dr = e.key === 'ArrowDown' || e.key === 'Enter' && !e.shiftKey ? 1 : e.key === 'ArrowUp' || e.key === 'Enter' && e.shiftKey ? -1 : 0, dc = e.key === 'ArrowRight' || e.key === 'Tab' && !e.shiftKey ? 1 : e.key === 'ArrowLeft' || e.key === 'Tab' && e.shiftKey ? -1 : 0;
        if (mod && e.key.startsWith('Arrow')) { const dest = this.jumpEdge(dr, dc); r = dest.r; c = dest.c; } else { const merge = this.sheet.mergeAt(r, c); if (merge) { if (dr > 0) r = merge.r2; if (dc > 0) c = merge.c2; } r += dr; c += dc; }
      }
      while (this.sheet.hiddenRows.has(r) && r >= 0 && r < MAX_ROWS) r += e.key === 'ArrowUp' || e.key === 'PageUp' ? -1 : 1;
      const direction = e.key === 'ArrowLeft' || e.key === 'Tab' && e.shiftKey ? -1 : 1;
      while (this.sheet.hiddenCols.has(c) && c >= 0 && c < MAX_COLS) c += direction;
      if (c < 0 || c >= MAX_COLS) c = this.active.c;
      this.goto(r, c, e.shiftKey && !['Enter','Tab'].includes(e.key)); return;
    }
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); if (this.editable()) this.workbook.clear(this.sheet, this.selection); return; }
    if (!mod && !e.altKey && e.key.length === 1) { e.preventDefault(); this.startEdit(e.key); }
    else if (e.key === 'Process' || e.keyCode === 229) this.startEdit('');
  }
  jumpEdge(dr, dc) {
    let { r, c } = this.active; const step = () => ({ r: r + dr, c: c + dc }), within = p => p.r >= 0 && p.c >= 0 && p.r < MAX_ROWS && p.c < MAX_COLS;
    if (this.sheet.raw(r, c) && within(step()) && this.sheet.raw(r + dr, c + dc)) { while (within(step()) && this.sheet.raw(r + dr, c + dc)) { r += dr; c += dc; } return { r, c }; }
    let best = null, distance = Infinity;
    for (const [key, cell] of this.sheet.cells) { if (!cell.raw) continue; const [rr, cc] = key.split(',').map(Number), delta = dr ? (rr - r) * dr : (cc - c) * dc; if ((dr ? cc === c : rr === r) && delta > 0 && delta < distance) { distance = delta; best = { r: rr, c: cc }; } }
    return best || { r: dr ? dr > 0 ? MAX_ROWS - 1 : 0 : r, c: dc ? dc > 0 ? MAX_COLS - 1 : 0 : c };
  }
  renderRibbon() {
    $$('[data-tab]').forEach(b => b.classList.toggle('selected', b.dataset.tab === this.tab));
    const stack = (...items) => `<div class="tool-stack">${items.join('')}</div>`;
    let html = '';
    if (this.tab === 'Home') {
      html += group('Clipboard', tool('paste','Paste','paste',true,true) + stack(tool('cut','Cut','cut'),tool('copy','Copy','copy')) + stack(tool('format-painter','Format painter','paint')));
      html += group('Font', `<div class="font-tools"><div class="ribbon-row"><select id="font-family" class="font-select" aria-label="Font family"><option>Aptos</option><option>Arial</option><option>Georgia</option><option>Verdana</option><option>Courier New</option></select><select id="font-size" class="font-size" aria-label="Font size">${[8,9,10,11,12,14,16,18,20,24,28,32,36,48].map(s => `<option${s === 10 ? ' selected' : ''}>${s}</option>`).join('')}</select>${mini('font-larger','plus','Increase font size','A⁺')}${mini('font-smaller','plus','Decrease font size','A⁻')}</div><div class="ribbon-row">${mini('bold','','Bold (Ctrl/⌘ B)','<span class="text-bold">B</span>')}${mini('italic','','Italic (Ctrl/⌘ I)','<span class="text-italic">I</span>')}${mini('underline','','Underline (Ctrl/⌘ U)','<span class="text-underline">U</span>')}<span class="ribbon-sep"></span>${mini('borders','border','Toggle cell borders')}<span class="ribbon-sep"></span><button class="mini-tool fill-color-tool" data-action="fill-color" title="Fill color" aria-label="Fill color">${icon('fill')}</button><button class="mini-tool font-color-tool" data-action="text-color" title="Font color" aria-label="Font color">A</button></div></div>`);
      html += group('Alignment', `<div class="font-tools"><div class="ribbon-row">${mini('align-left','alignLeft','Align left')}${mini('align-center','alignCenter','Center')}${mini('align-right','alignRight','Align right')}${tool('wrap','Wrap text','wrap')}</div><div class="ribbon-row">${tool('merge','Merge & center','merge')}</div></div>`);
      html += group('Number', `<div class="font-tools"><div class="ribbon-row"><select id="number-format" class="number-format" aria-label="Number format"><option value="general">General</option><option value="number">Number</option><option value="currency">Currency</option><option value="percent">Percentage</option><option value="date">Short date</option><option value="integer">Integer</option></select></div><div class="ribbon-row">${mini('currency','','Currency','$')}${mini('percent','percent','Percentage')}${mini('number','','Number with separator',',')}<span class="ribbon-sep"></span>${mini('decimal-less','','Decrease decimals','.0←')}${mini('decimal-more','','Increase decimals','→.00')}${mini('format-cells','table','Format cells (Ctrl/⌘ 1)')}</div></div>`);
      html += group('Styles', tool('conditional','Conditional<br>formatting','conditional',true,true) + tool('format-table','Format as<br>table','table',true,true) + `<div class="style-gallery"><button class="style-chip" data-action="style-normal">Normal</button><button class="style-chip good" data-action="style-good">Good</button><button class="style-chip heading" data-action="style-heading">Heading</button><button class="style-chip warning" data-action="style-warning">Warning</button></div>`, 'styles-group');
      html += group('Cells', tool('insert-menu','Insert','insert',true,true) + tool('delete-menu','Delete','delete',true,true));
      html += group('Editing', stack(tool('autosum','AutoSum','sum'),tool('clear-menu','Clear','clear')) + tool('sort-filter','Sort &<br>filter','sort',true,true) + tool('find','Find &<br>select','search',true,true));
    } else if (this.tab === 'Insert') {
      html = group('Tables',tool('format-table','Table','table',true)) + group('Charts',tool('chart','Column chart','chart',true)+tool('chart-line','Line chart','line',true)+tool('chart-bar','Bar chart','sort',true)+tool('chart-donut','Doughnut chart','donut',true)) + group('Text & references',tool('add-note','Cell note','comment',true)+tool('name-manager','Named range','name',true)) + group('Worksheets',tool('add-sheet','New sheet','plus',true)+tool('duplicate-sheet','Duplicate sheet','copy',true));
    } else if (this.tab === 'Page Layout') {
      html = group('Page setup',tool('print','Print worksheet','print',true)+tool('print-selection','Print selection','grid',true)) + group('Sheet options',tool('toggle-gridlines','Gridlines','grid',true)+tool('auto-fit','Auto-fit columns','table',true)+tool('wrap','Wrap text','wrap',true)) + group('Workbook appearance',tool('theme','Light / dark','moon',true)+tool('zoom-reset','Actual size','search',true));
    } else if (this.tab === 'Formulas') {
      html = group('Function library',tool('functions','Insert function','function',true)+tool('autosum','AutoSum','sum',true)) + group('Defined names',tool('name-manager','Name manager','name',true)) + group('Formula auditing',tool('show-formulas','Show formulas','function',true)+tool('inspect-formula','Inspect cell','search',true)) + group('Calculation',tool('recalculate','Calculate now','refresh',true));
    } else if (this.tab === 'Data') {
      html = group('Get data',tool('open','From CSV / XLSX','open',true)+tool('export-csv','Export CSV','export',true)) + group('Sort & filter',tool('sort-asc','Sort A to Z','sort',true)+tool('sort-desc','Sort Z to A','sort',true)+tool('sort','Custom sort','table',true)+tool('filter','Filter','filter',true)+tool('clear-filter','Clear filters','clear',true)) + group('Data tools',tool('text-to-columns','Text to Columns','table',true)+tool('remove-duplicates','Remove duplicates','table',true)+tool('recalculate','Recalculate','refresh',true));
    } else if (this.tab === 'Review') {
      html = group('Notes',tool('add-note','New note','comment',true)+tool('notes','Show all notes','comment',true)) + group('Protection',tool('protect',this.sheet.protected ? 'Enable editing' : 'Read-only sheet','lock',true)) + group('Workbook',tool('inspect-formula','Inspect active cell','search',true)+tool('about','About Gridline','info',true));
    } else if (this.tab === 'View') {
      html = group('Show',tool('toggle-gridlines','Gridlines','grid',true)+tool('show-formulas','Formulas','function',true)+tool('theme','Light / dark','moon',true)) + group('Zoom',tool('zoom-in','Zoom in','search',true)+tool('zoom-out','Zoom out','search',true)+tool('zoom-reset','100%','grid',true)) + group('Columns',tool('hide-columns','Hide columns','table',true)+tool('unhide-columns','Unhide columns','table',true)+tool('unhide-all-columns','Unhide all','table',true)) + group('Window',tool('freeze','Freeze panes','freeze',true)+tool('freeze-top','Freeze top row','table',true)+tool('freeze-first','Freeze first column','table',true)+tool('unfreeze','Unfreeze panes','clear',true)) + group('Engine',tool('performance','Performance','grid',true));
    } else html = group('Get started',tool('help','Keyboard shortcuts','info',true)+tool('commands','Find a command','search',true)+tool('sample','Load demo workbook','table',true)) + group('Gridline',tool('about','About & limitations','info',true)+tool('performance','Engine diagnostics','grid',true));
    $('#ribbon').innerHTML = html;
    $('#font-family')?.addEventListener('change', e => this.errorBoundary(() => this.format({ fontFamily: e.target.value })));
    $('#font-size')?.addEventListener('change', e => this.errorBoundary(() => this.format({ fontSize: +e.target.value / .75 })));
    $('#number-format')?.addEventListener('change', e => this.errorBoundary(() => this.format({ format: e.target.value, decimals: undefined })));
    this.updateUI();
  }
  renderTabs() {
    $('#sheet-tabs').innerHTML = this.workbook.sheets.map(s => `<button class="sheet-tab ${s.id === this.sheet.id ? 'selected' : ''}" data-sheet="${escapeHTML(s.id)}" title="Double-click to rename · Right-click for options"><span class="sheet-tab-dot" style="background:${/^#[0-9a-f]{6}$/i.test(s.color) ? s.color : '#18835a'}"></span>${escapeHTML(s.name)}</button>`).join('');
  }
  switchSheet(id) {
    if (id === this.sheet.id) return; if (this.panelType === 'find') this.closePanel(); this.commitEdit(false); this.commitFormula(); this.sheetViews.set(this.sheet.id, { selection: { ...this.selection }, active: { ...this.active }, scrollX: this.renderer.scrollX, scrollY: this.renderer.scrollY });
    this.workbook.activeSheetId = id; const view = this.sheetViews.get(id); this.renderer.scrollX = view?.scrollX ?? 0; this.renderer.scrollY = view?.scrollY ?? 0; this.renderer.copyRange = null; this.renderer.syncLayout(); this.anchor = view?.active || { r: 0, c: 0 }; this.select(view?.selection || normalizedRange(this.anchor), this.anchor, false); this.renderTabs(); this.renderCharts(); if (this.panelType === 'notes') this.showNotes(); this.host.focus();
  }
  setZoom(zoom) { this.renderer.setZoom(Math.round(zoom * 10) / 10); $('#zoom-value').textContent = Math.round(this.renderer.zoom * 100) + '%'; $('#zoom-slider').value = this.renderer.zoom * 100; this.positionCharts(); }
  format(style) { if (this.editable()) this.workbook.applyStyle(this.sheet, this.selection, style); }
  showColorPicker(property) {
    if (!this.editable()) return;
    const sheet = this.sheet, q = {...this.selection}, current = sheet.style(this.active.r,this.active.c)[property];
    const initial = /^#[0-9a-f]{6}$/i.test(current || '') ? current : property === 'fill' ? '#e2efda' : '#176b4a';
    const colors = ['#ffffff','#000000','#293b32','#176b4a','#e2efda','#ff0000','#ffc000','#ffff00','#00b050','#00b0f0','#0070c0','#7030a0'];
    this.openDialog(property === 'fill' ? 'Fill Color' : 'Text Color', `<div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px;margin-bottom:18px">${colors.map(color=>`<button type="button" data-swatch="${color}" aria-label="Color ${color}" title="${color}" style="height:32px;background:${color};border:1px solid var(--line)"></button>`).join('')}</div><label class="field-label" for="color-picker">Custom color</label><input id="color-picker" type="color" value="${initial}" style="width:100%;height:40px"><label class="field-label" for="color-hex">Hex color</label><input id="color-hex" class="dialog-input" value="${initial}" pattern="#[0-9a-fA-F]{6}" required maxlength="7"><div class="dialog-actions"><button class="secondary-btn" data-action="close-dialog">Cancel</button><button id="color-apply" class="primary-btn">Apply</button></div>`, 360);
    const picker = $('#color-picker'), hex = $('#color-hex');
    $$('[data-swatch]').forEach(button => button.onclick = () => { picker.value = hex.value = button.dataset.swatch; });
    picker.oninput = () => { hex.value = picker.value; };
    hex.oninput = () => { if (hex.validity.valid) picker.value = hex.value; };
    $('#color-apply').onclick = () => this.errorBoundary(() => { if (!hex.reportValidity()) return; this.workbook.applyStyle(sheet,q,{[property]:hex.value.toLowerCase()}); this.closeDialog(); });
  }
  showFormatCells() {
    if (!this.editable()) return;
    const q = { ...this.selection }, sheet = this.sheet, current = sheet.style(this.active.r, this.active.c);
    const initial = { ...current, ...numberFormatStyle(current.format), ...Object.fromEntries(Object.entries(current).filter(([k,v]) => k !== 'format' && v != null)) }, categories = { general:'General', number:'Number', currency:'Currency', date:'Date', time:'Time', percent:'Percentage', text:'Text' };
    const scope = q.r1 === 0 && q.r2 === MAX_ROWS - 1 ? `Columns ${colName(q.c1)}:${colName(q.c2)}` : q.c1 === 0 && q.c2 === MAX_COLS - 1 ? `Rows ${q.r1 + 1}:${q.r2 + 1}` : rangeAddress(q);
    this.openDialog('Format Cells', `<p class="help-text">${escapeHTML(scope)} · Number format</p><div class="format-layout"><div><label class="field-label" for="format-category">Category</label><select id="format-category" class="dialog-input format-categories" size="7">${Object.entries(categories).map(([v,label]) => `<option value="${v}">${label}</option>`).join('')}</select></div><div><label class="field-label" for="format-sample">Sample</label><output id="format-sample" class="format-sample"></output><div id="format-options"></div><p id="format-description" class="help-text"></p></div></div><div class="dialog-actions"><button class="secondary-btn" data-action="close-dialog">Cancel</button><button id="format-apply" class="primary-btn">OK</button></div>`, 620);
    const category = $('#format-category'); category.value = initial.format === 'integer' ? 'number' : categories[initial.format] ? initial.format : 'general';
    const read = () => ({ format: category.value, decimals: $('#format-decimals') ? +$('#format-decimals').value : null, currency: $('#format-currency')?.value || 'USD', grouping: $('#format-grouping')?.checked ?? true, pattern: $('#format-pattern')?.value || null });
    const preview = () => {
      const style = read(), value = this.workbook.value(sheet, this.active.r, this.active.c);
      $('#format-sample').textContent = formatValue(value ?? (['date','time'].includes(style.format) ? 45292.5 : 1234.567), style);
    };
    const options = (useCurrent = false) => {
      const fmt = category.value, patterns = [...new Set([...(fmt === 'date' ? DATE_FORMATS : TIME_FORMATS), ...(useCurrent && initial.pattern ? [initial.pattern] : [])])];
      $('#format-options').innerHTML = ['number','currency','percent'].includes(fmt) ? `<label class="field-label" for="format-decimals">Decimal places</label><input id="format-decimals" class="dialog-input" type="number" min="0" max="10" step="1" required value="2">${fmt === 'currency' ? `<label class="field-label" for="format-currency">Currency symbol</label><select id="format-currency" class="dialog-input">${Object.entries(CURRENCIES).map(([code,symbol]) => `<option value="${code}">${symbol} — ${code}</option>`).join('')}</select>` : ''}${fmt !== 'percent' ? '<p><label><input id="format-grouping" type="checkbox" checked> Use thousands separator</label></p>' : ''}` : ['date','time'].includes(fmt) ? `<label class="field-label" for="format-pattern">Type</label><select id="format-pattern" class="dialog-input" size="${patterns.length}">${patterns.map(pattern => `<option value="${escapeHTML(pattern)}">${escapeHTML(formatValue(45292.5625,{format:fmt,pattern}))}</option>`).join('')}</select>` : '';
      if ($('#format-decimals')) $('#format-decimals').value = useCurrent && Number.isInteger(initial.decimals) && initial.decimals >= 0 && initial.decimals <= 10 ? initial.decimals : useCurrent ? (initial.format === 'number' ? 2 : initial.format === 'percent' ? 1 : 0) : 2;
      if (useCurrent) { if ($('#format-currency')) $('#format-currency').value = initial.currency || 'USD'; if ($('#format-grouping')) $('#format-grouping').checked = initial.grouping !== false; }
      if ($('#format-pattern')) $('#format-pattern').value = useCurrent && patterns.includes(initial.pattern) ? initial.pattern : patterns[0];
      $('#format-description').textContent = { general:'General displays values without a specific number format.', number:'Change how numbers appear without changing their values.', currency:'Display amounts with the selected currency symbol.', percent:'Display values multiplied by 100 with a percent sign.', date:'Enter dates as yyyy-mm-dd or m/d/yyyy (four-digit year). Dates remain numeric values for calculations.', time:'Enter times as hh:mm or hh:mm:ss, optionally with AM/PM.', text:'New entries are kept as text, including leading zeros. Existing values and formulas are unchanged.' }[fmt];
      preview();
    };
    category.onchange = () => options();
    $('#format-options').oninput = () => { if (!$('#format-decimals') || $('#format-decimals').validity.valid) preview(); };
    $('#format-apply').onclick = () => this.errorBoundary(() => { const decimals = $('#format-decimals'); if (decimals && !decimals.reportValidity()) return; this.workbook.applyStyle(sheet, q, read()); this.closeDialog(); });
    options(true);
  }
  autoFit(column = null) {
    if (!this.editable()) return; const cols = column === null ? [this.selection.c1, this.selection.c2] : [column, column], sizes = new Map();
    for (let c = cols[0]; c <= cols[1]; c++) sizes.set(c,54);
    for (const [key] of this.sheet.cells) { const [r,c] = key.split(',').map(Number); if (!sizes.has(c) || this.sheet.mergeAt(r,c)) continue; const text = this.workbook.display(this.sheet,r,c), font = this.renderer.font(this.sheet.style(r,c)); sizes.set(c,Math.min(500,Math.max(sizes.get(c),Math.ceil(this.renderer.measureText(text,font) / this.renderer.zoom + 23)))); }
    this.workbook.mutate('Auto-fit columns', () => { for (const [c, width] of sizes) this.sheet.colWidths.set(c, width); });
  }
  dataRange() {
    const q = this.selection;
    if (q.r1 !== q.r2 || q.c1 !== q.c2) {
      const used = this.sheet.usedRange(); return { r1: q.r1, c1: q.c1, r2: Math.min(q.r2, used.r2), c2: Math.min(q.c2, used.c2) };
    }
    const filter = this.sheet.filters?.range || this.sheet.dataRegion;
    if (filter && this.active.r >= filter.r1 && this.active.r <= filter.r2 && this.active.c >= filter.c1 && this.active.c <= filter.c2) return { ...filter };
    let { r, c } = this.active, r1 = r, r2 = r, c1 = c, c2 = c;
    while (c1 > 0 && this.sheet.raw(r, c1 - 1)) c1--; while (c2 < MAX_COLS - 1 && this.sheet.raw(r, c2 + 1)) c2++;
    const filled = rr => { for (let cc = c1; cc <= c2; cc++) if (this.sheet.raw(rr, cc)) return true; return false; };
    while (r1 > 0 && filled(r1 - 1)) r1--; while (r2 < MAX_ROWS - 1 && r2 - r1 < 10000 && filled(r2 + 1)) r2++;
    return { r1, c1, r2, c2 };
  }
  async run(action, element = null) {
    if (!['cancel-edit','commit-edit','close-dialog'].includes(action)) { this.commitEdit(false); this.commitFormula(); }
    const style = this.sheet.style(this.active.r, this.active.c), q = this.selection;
    if (['bold','italic','underline','wrap'].includes(action)) return this.format({ [action]: !style[action] });
    if (action.startsWith('align-')) return this.format({ align: action.slice(6) });
    switch (action) {
      case 'file': return this.showFile();
      case 'export': return this.showExport();
      case 'new': return this.confirm('Create a new workbook?', 'Export your current workbook first to keep a separate copy. The new workbook will replace the local autosave.', () => this.setWorkbook(new Workbook()), 'Create workbook');
      case 'sample': return this.confirm('Load the demo workbook?', 'This replaces the current local workbook with the illustrative revenue workbook.', () => { this.setWorkbook(createSampleWorkbook()); this.goto(12, 6); }, 'Load demo');
      case 'open': this.closeDialog(); $('#file-input').click(); return;
      case 'save': this.persist(); downloadFile(this.fileName('.gridline'), JSON.stringify(this.workbook.toJSON(), null, 2), 'application/json'); this.toast('Gridline workbook exported with all app features.'); return;
      case 'export-xlsx': {
        const name = this.fileName('.xlsx'), workbook = this.workbook, type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        if (typeof window.showSaveFilePicker === 'function') {
          let handle;
          try { handle = await window.showSaveFilePicker({suggestedName:name,types:[{description:'Excel workbook',accept:{[type]:['.xlsx']}}]}); }
          catch (error) { if (error.name === 'AbortError') return; throw error; }
          const blob = new Blob([exportXLSX(workbook)],{type});
          await blob.stream().pipeTo(await handle.createWritable());
        } else downloadFile(name,exportXLSX(workbook),type);
        this.closeDialog(); this.toast('XLSX exported. Charts, notes, and conditional rules remain in the .gridline format.'); return;
      }
      case 'export-csv': downloadFile(this.sheet.name + '.csv', exportCSV(this.workbook), 'text/csv;charset=utf-8'); this.closeDialog(); this.toast('Current sheet exported as CSV values.'); return;
      case 'autosave': this.autosave = !this.autosave; $('#autosave-toggle').classList.toggle('on', this.autosave); if (this.autosave) this.persist(); else $('#save-status').textContent = 'Autosave is off'; return;
      case 'undo': this.workbook.undo(); return;
      case 'redo': this.workbook.redo(); return;
      case 'copy': return this.copy(false);
      case 'cut': return this.copy(true);
      case 'paste': return this.paste();
      case 'format-cells': return this.showFormatCells();
      case 'format-painter': this.paintStyle = structuredClone(style); this.toast('Select a cell or range to apply the current formatting.'); return;
      case 'fill-color': return this.showColorPicker('fill');
      case 'text-color': return this.showColorPicker('color');
      case 'borders': return this.format({ border: !style.border });
      case 'currency': case 'percent': case 'number': return this.format({ format: action, decimals: undefined });
      case 'decimal-less': return this.format({ decimals: Math.max(0, (style.decimals ?? 2) - 1) });
      case 'decimal-more': return this.format({ decimals: Math.min(10, (style.decimals ?? 0) + 1) });
      case 'font-larger': return this.format({ fontSize: Math.min(72, (style.fontSize || 13) + 2) });
      case 'font-smaller': return this.format({ fontSize: Math.max(8, (style.fontSize || 13) - 2) });
      case 'style-normal': return this.format({ fill:'#ffffff', color:'#293b32', bold:false, italic:false, border:false, fontSize:13 });
      case 'style-good': return this.format({ fill:'#e0efdf', color:'#37683d' });
      case 'style-heading': return this.format({ fill:'#176b4a', color:'#ffffff', bold:true });
      case 'style-warning': return this.format({ fill:'#fff0d4', color:'#8d6224' });
      case 'merge': return this.mergeSelection();
      case 'text-to-columns': return this.showTextToColumns();
      case 'hide-columns': case 'unhide-columns': case 'unhide-all-columns':
        if (!this.editable()) return;
        if (action === 'hide-columns' && new Set([...this.sheet.hiddenCols, ...Array.from({length:q.c2-q.c1+1},(_,i)=>q.c1+i)]).size === MAX_COLS) throw new Error('Keep at least one column visible.');
        this.workbook.mutate(action === 'hide-columns' ? 'Hide columns' : 'Unhide columns', () => {
          if (action === 'unhide-all-columns') this.sheet.hiddenCols.clear();
          else for (let c=q.c1;c<=q.c2;c++) this.sheet.hiddenCols[action === 'hide-columns' ? 'add' : 'delete'](c);
        });
        if (this.sheet.hiddenCols.has(this.active.c)) { let c=this.active.c; while(c<MAX_COLS-1 && this.sheet.hiddenCols.has(c)) c++; while(this.sheet.hiddenCols.has(c)) c--; this.goto(this.active.r,c); }
        return;
      case 'auto-fit': return this.autoFit();
      case 'format-table': return this.formatTable();
      case 'conditional': return this.showConditional();
      case 'insert-menu': return this.menuAt(element, [['insert-row','Insert row above'],['insert-column','Insert column to the left'],['add-sheet','Insert worksheet']]);
      case 'delete-menu': return this.menuAt(element, [['delete-row','Delete selected row'],['delete-column','Delete selected column'],null,['clear','Clear contents'],['delete-sheet','Delete worksheet']]);
      case 'clear-menu': return this.menuAt(element, [['clear','Clear contents'],['clear-format','Clear formatting'],['clear-all','Clear all']]);
      case 'clear': if (this.editable()) this.workbook.clear(this.sheet, this.selection); return;
      case 'clear-all': if (this.editable()) this.workbook.clear(this.sheet, this.selection, true); return;
      case 'clear-format': return this.format({ format:'general', bold:false, italic:false, underline:false, wrap:false, align:null, fill:null, color:null, border:false, borderColor:null, bottomBorder:null, fontFamily:'Aptos', fontSize:13 });
      case 'insert-column': case 'delete-column': if (this.editable()) { const {c1,c2} = this.selection; this.workbook.structuralEdit(this.sheet,'column',c1,(c2-c1+1)*(action==='insert-column'?1:-1)); } return;
      case 'insert-row': case 'delete-row': if (this.editable()) { const axis = action.endsWith('row') ? 'row' : 'column'; this.workbook.structuralEdit(this.sheet, axis, axis === 'row' ? this.active.r : this.active.c, action.startsWith('insert') ? 1 : -1); this.toast('Structural edit applied. Chart, filter and conditional-rule metadata on this sheet was reset; Undo restores it.'); } return;
      case 'autosum': return this.autoSum();
      case 'fill-down': if (this.editable()) { const q = this.selection; const source = { ...q, r2: q.r1 }; if (q.r1 === q.r2 && q.r1 > 0) { source.r1 = source.r2 = q.r1 - 1; } this.workbook.fill(this.sheet, source, q); } return;
      case 'fill-right': if (this.editable()) { const q = this.selection, source = { ...q, c2: q.c1 }; if (q.c1 === q.c2 && q.c1 > 0) source.c1 = source.c2 = q.c1 - 1; this.workbook.fill(this.sheet, source, q); } return;
      case 'select-all': this.select({ r1:0,c1:0,r2:MAX_ROWS-1,c2:MAX_COLS-1 },this.active); return;
      case 'sort-filter': return this.menuAt(element, [['sort-asc','Sort A to Z'],['sort-desc','Sort Z to A'],['sort','Custom sort…'],null,['filter','Filter values…'],['clear-filter','Clear filters']]);
      case 'sort': return this.showSort();
      case 'sort-asc': case 'sort-desc': if (this.editable()) { const q = this.dataRange(); this.workbook.sort(this.sheet, q, Math.max(q.c1, Math.min(q.c2,this.active.c)), action === 'sort-desc', true); } return;
      case 'filter': return this.showFilter();
      case 'clear-filter': this.workbook.mutate('Clear filters', () => { this.sheet.hiddenRows.clear(); if (this.sheet.filters) this.sheet.filters.criteria = {}; }); return;
      case 'remove-duplicates': return this.removeDuplicates();
      case 'find': return this.showFind();
      case 'notes': return this.showNotes();
      case 'add-note': return this.addNote();
      case 'delete-note': if (this.editable()) this.workbook.setCell(this.sheet, +element.dataset.row, +element.dataset.col, { note: undefined }); return;
      case 'protect': this.workbook.mutate('Toggle read-only sheet', () => this.sheet.protected = !this.sheet.protected); this.renderRibbon(); this.toast(this.sheet.protected ? 'Sheet editing is disabled in this app. This is not encryption.' : 'Sheet editing enabled.'); return;
      case 'functions': return this.showFunctions();
      case 'commands': return this.showCommands();
      case 'name-manager': return this.showNames();
      case 'recalculate': this.workbook.engine.invalidate(); this.renderCharts(); this.updateUI(); this.toast('Workbook recalculated.'); return;
      case 'inspect-formula': return this.showCellInspector();
      case 'show-formulas': this.renderer.showFormulas = !this.renderer.showFormulas; this.renderer.requestFrame(); return;
      case 'chart': case 'chart-line': case 'chart-bar': case 'chart-donut': return this.showChart(action === 'chart' ? 'column' : action.slice(6));
      case 'remove-chart': if (!this.editable()) return; this.workbook.mutate('Remove chart', () => this.sheet.charts = this.sheet.charts.filter(c => c.id !== element.dataset.chart)); return;
      case 'freeze': case 'freeze-top': case 'freeze-first': case 'unfreeze': this.workbook.mutate('Freeze panes', () => { this.sheet.freezeRows = action === 'freeze' ? Math.min(1000,this.active.r) : action === 'freeze-top' ? 1 : 0; this.sheet.freezeCols = action === 'freeze' ? Math.min(100,this.active.c) : action === 'freeze-first' ? 1 : 0; }); this.renderer.scrollX = this.renderer.scrollY = 0; this.renderer.requestFrame(); return;
      case 'toggle-gridlines': this.workbook.mutate('Toggle gridlines', () => this.sheet.gridlines = !this.sheet.gridlines); return;
      case 'theme': document.body.classList.toggle('dark'); this.renderer.dark = document.body.classList.contains('dark'); this.renderer.requestFrame(); this.renderCharts(); return;
      case 'zoom-in': return this.setZoom(this.renderer.zoom + .1);
      case 'zoom-out': return this.setZoom(this.renderer.zoom - .1);
      case 'zoom-reset': return this.setZoom(1);
      case 'add-sheet': this.workbook.addSheet(); this.renderer.scrollX = this.renderer.scrollY = 0; this.goto(0,0); return;
      case 'duplicate-sheet': this.workbook.duplicateSheet(this.sheet); return;
      case 'rename-sheet': return this.renameSheet();
      case 'delete-sheet': return this.confirm(`Delete “${this.sheet.name}”?`, 'This worksheet will be removed. You can undo the deletion.', () => this.workbook.deleteSheet(this.sheet), 'Delete worksheet');
      case 'prev-sheet': case 'next-sheet': { const at = this.workbook.sheets.findIndex(s => s.id === this.sheet.id), next = this.workbook.sheets[at + (action === 'prev-sheet' ? -1 : 1)]; if (next) this.switchSheet(next.id); return; }
      case 'sheets': return this.showSheets();
      case 'performance': return this.showPerformance();
      case 'help': return this.showHelp();
      case 'about': return this.showAbout();
      case 'print': return this.printSheet(false);
      case 'print-selection': return this.printSheet(true);
      case 'cancel-edit': return this.cancelEdit();
      case 'commit-edit': this.commitEdit(false); this.commitFormula(); this.host.focus(); return;
      case 'expand-formula': $('.formula-row').classList.toggle('expanded'); return;
      case 'close-dialog': return this.closeDialog();
      case 'close-panel': return this.closePanel();
      default: throw new Error(`Unknown command: ${action}`);
    }
  }
  fileName(extension) { return this.workbook.title.replace(/[<>:"/\\|?*]/g, '-').slice(0, 100) + extension; }
  async openFile(file) {
    if (file.size > 32 * 1024 * 1024) throw new Error('Files must be 32 MB or smaller.'); this.toast('Opening ' + file.name + '…');
    const title = file.name.replace(/\.[^.]+$/, ''); let workbook, warnings;
    if (/\.xlsx$/i.test(file.name)) { const result = await importXLSX(await file.arrayBuffer(), title); workbook = result.workbook; warnings = result.warnings; }
    else if (/\.(gridline|json)$/i.test(file.name)) workbook = Workbook.fromJSON(JSON.parse(await file.text()));
    else workbook = workbookFromCSV(await file.text(), title);
    this.setWorkbook(workbook); this.toast(warnings?.[0] || `Opened ${file.name}.`);
  }
  clipboardPayload(cut = false) {
    const q = this.sheet.populatedRange(this.selection), rows = [], cells = [];
    for (const { r,c } of cellsIn(q)) {
      if (!rows[r-q.r1]) { rows[r-q.r1] = []; cells[r-q.r1] = []; }
      rows[r-q.r1][c-q.c1] = this.workbook.display(this.sheet,r,c); cells[r-q.r1][c-q.c1] = structuredClone({ ...(this.sheet.get(r,c) || { raw:'' }), style: this.sheet.style(r,c) });
    }
    const text = serializeDelimited(rows,'\t'); this.clipboard = { text, cells, source:{...q}, sheetId:this.sheet.id, cut }; this.renderer.copyRange = {...q}; this.renderer.requestFrame(); return text;
  }
  copyToEvent(e, cut) { this.errorBoundary(() => { e.clipboardData.setData('text/plain',this.clipboardPayload(cut)); e.preventDefault(); }); }
  async copy(cut) {
    const text = this.clipboardPayload(cut);
    try { await navigator.clipboard.writeText(text); this.toast(cut ? 'Cut selection. Paste to move its cells.' : 'Selection copied.'); }
    catch { this.toast('Selection copied inside Gridline. Use Paste, or Ctrl/⌘ C to copy to your system clipboard.'); }
  }
  async paste() { let text; try { text = await navigator.clipboard.readText(); } catch { text = this.clipboard?.text; } if (text === undefined) { this.toast('Use Ctrl/⌘ V to paste from your clipboard.'); this.host.focus(); return; } this.pasteText(text); }
  pasteText(text) {
    if (!this.editable()) return; const internal = this.clipboard?.text === text ? this.clipboard : null, values = internal ? internal.cells : parseDelimited(text,'\t');
    const r0 = this.active.r, c0 = this.active.c, h = values.length, w = Math.max(...values.map(r => r.length)); if (h*w > MAX_RANGE_CELLS || r0+h > MAX_ROWS || c0+w > MAX_COLS) throw new Error('The pasted range exceeds sheet or operation limits.');
    this.workbook.transaction(internal?.cut ? 'Move cells' : 'Paste cells', () => {
      if (internal?.cut) { const source = this.workbook.sheets.find(s => s.id === internal.sheetId); if (source?.protected) throw new Error('The source sheet is read-only.'); if (source) this.workbook.clear(source,internal.source,true); }
      values.forEach((row,ri) => row.forEach((value,ci) => {
        const cell = internal ? structuredClone(value) : {raw:String(value)};
        if (internal && !internal.cut) cell.raw = shiftFormula(cell.raw, r0-internal.source.r1, c0-internal.source.c1);
        if (internal) { this.workbook.setCell(this.sheet,r0+ri,c0+ci,null); this.workbook.setCell(this.sheet,r0+ri,c0+ci,cell); } else this.workbook.setRaw(this.sheet,r0+ri,c0+ci,cell.raw);
      }));
    });
    if (internal?.cut) { this.clipboard = null; this.renderer.copyRange = null; }
    this.select({r1:r0,c1:c0,r2:r0+h-1,c2:c0+w-1},{r:r0,c:c0}); this.host.focus();
  }
  autoSum() {
    if (!this.editable()) return; const q = this.selection;
    if (q.r1 !== q.r2 || q.c1 !== q.c2) {
      this.workbook.transaction('AutoSum', () => { for (let c=q.c1;c<=q.c2;c++) this.workbook.setRaw(this.sheet,q.r2+1,c,`=SUM(${address(q.r1,c)}:${address(q.r2,c)})`); }); this.goto(q.r2+1,q.c1); return;
    }
    const {r,c}=this.active; let start=r-1; while(start>=0&&typeof this.workbook.value(this.sheet,start,c)==='number')start--;
    if(start<r-1) { this.workbook.setRaw(this.sheet,r,c,`=SUM(${address(start+1,c)}:${address(r-1,c)})`); return; }
    let left=c-1; while(left>=0&&typeof this.workbook.value(this.sheet,r,left)==='number')left--;
    if(left<c-1)this.workbook.setRaw(this.sheet,r,c,`=SUM(${address(r,left+1)}:${address(r,c-1)})`); else this.startEdit('=SUM(');
  }
  mergeSelection() {
    if (!this.editable()) return; const q={...this.selection}, overlap=m=>!(m.r2<q.r1||m.r1>q.r2||m.c2<q.c1||m.c1>q.c2), existing=this.sheet.merges.some(overlap);
    if(existing) { this.workbook.mutate('Unmerge cells',()=>this.sheet.merges=this.sheet.merges.filter(m=>!overlap(m))); return; }
    if(q.r1===q.r2&&q.c1===q.c2)return;
    const apply=()=>this.workbook.mutate('Merge cells',()=> { this.sheet.merges.push(q); const first=this.sheet.get(q.r1,q.c1)||{raw:''}; first.style={...first.style,align:'center'}; this.sheet.cells.set(keyOf(q.r1,q.c1),first); for(const p of cellsIn(q))if(p.r!==q.r1||p.c!==q.c1){ const cell=this.sheet.get(p.r,p.c); if(cell)cell.raw=''; } });
    let nonempty=0;for(const p of cellsIn(q))if(this.sheet.raw(p.r,p.c))nonempty++;
    if(nonempty>1)this.confirm('Merge selected cells?','Only the upper-left value is kept. Other values in the selection are cleared. Undo restores them.',apply,'Merge cells');else apply();
  }
  formatTable() {
    if(!this.editable())return;const q=this.dataRange();if(q.r2<=q.r1){this.toast('Select a rectangular range with a header and at least one data row.');return;}
    [...cellsIn(q)];
    this.workbook.mutate('Format as table',()=>{for(const {r,c} of cellsIn(q)){const cell=this.sheet.get(r,c)||{raw:''};cell.style={...cell.style,fill:r===q.r1?'#176b4a':(r-q.r1)%2?'#ffffff':'#eff6f1',color:r===q.r1?'#ffffff':'#344f3e',bold:r===q.r1};this.sheet.cells.set(keyOf(r,c),cell);}this.sheet.filters={range:q,criteria:{}};});
  }
  removeDuplicates() {
    if(!this.editable())return;const q=this.dataRange();const rows=[],seen=new Set();let removed=0;
    for(let r=q.r1+1;r<=q.r2;r++){const cells=[];for(let c=q.c1;c<=q.c2;c++)cells.push(structuredClone(this.sheet.get(r,c)||{raw:''}));const signature=JSON.stringify(cells.map((_,i)=>this.workbook.value(this.sheet,r,q.c1+i)));if(seen.has(signature))removed++;else{seen.add(signature);rows.push({r,cells});}}
    if(!removed){this.toast('No duplicate data rows found.');return;}
    this.workbook.transaction('Remove duplicates',()=>{this.workbook.clear(this.sheet,{...q,r1:q.r1+1},true);rows.forEach((row,i)=>row.cells.forEach((cell,j)=>{cell.raw=shiftFormula(cell.raw,q.r1+1+i-row.r,0);this.workbook.setCell(this.sheet,q.r1+1+i,q.c1+j,cell);}));});this.toast(`${removed} duplicate row${removed===1?'':'s'} removed. Header preserved.`);
  }
  openDialog(title, body, width = 550) {
    this.dialog.style.width = width + 'px'; $('#dialog-content').innerHTML = `<div class="dialog-head"><h2>${escapeHTML(title)}</h2><button data-action="close-dialog" aria-label="Close dialog">×</button></div><div class="dialog-body">${body}</div>`;
    if (!this.dialog.open) this.dialog.showModal(); requestAnimationFrame(() => this.dialog.querySelector('[autofocus]')?.focus());
  }
  closeDialog() { if (this.dialog.open) this.dialog.close(); }
  confirm(title, message, callback, label = 'Continue') {
    this.openDialog(title, `<p class="help-text">${escapeHTML(message)}</p><div class="dialog-actions"><button class="secondary-btn" data-action="close-dialog">Cancel</button><button class="primary-btn" id="confirm-action">${escapeHTML(label)}</button></div>`);
    $('#confirm-action').onclick = () => this.errorBoundary(() => { this.closeDialog(); callback(); });
  }
  contextMenu(x, y, entries) {
    const menu = $('#context-menu'); menu.innerHTML = entries.map(entry => entry ? `<button data-action="${entry[0]}" class="${entry[0].startsWith('delete') ? 'danger' : ''}">${escapeHTML(entry[1])}${entry[2] ? `<kbd>${entry[2]}</kbd>` : ''}</button>` : '<hr>').join('');
    menu.hidden = false; menu.style.left = Math.max(5, Math.min(x, window.innerWidth - menu.offsetWidth - 8)) + 'px'; menu.style.top = Math.max(5, Math.min(y, window.innerHeight - menu.offsetHeight - 8)) + 'px';
    menu.onclick = () => menu.hidden = true;
  }
  menuAt(element, entries) { const box = element?.getBoundingClientRect(); this.contextMenu(box?.left ?? 200, box?.bottom ?? 200, entries); }
  cellContextMenu(x, y, header = null) { this.contextMenu(x,y,[['cut','Cut','Ctrl/⌘ X'],['copy','Copy','Ctrl/⌘ C'],['paste','Paste','Ctrl/⌘ V'],null,['insert-row','Insert row above'],['insert-column','Insert column left'],['delete-row','Delete row'],['delete-column','Delete column'],null,['clear','Clear contents','Delete'],['format-cells','Format cells…','Ctrl/⌘ 1'],['auto-fit','Auto-fit columns'],['hide-columns','Hide columns'],['unhide-columns','Unhide columns'],['unhide-all-columns','Unhide all columns'],['add-note','Add / edit note'],['inspect-formula','Inspect cell'],null,['filter','Filter values…']].filter(entry => !entry || (!header || !entry[0].startsWith('insert-') && !entry[0].startsWith('delete-') || entry[0].endsWith(header)) && (header !== 'row' || entry[0] !== 'auto-fit') && (!['hide-columns','unhide-columns','unhide-all-columns'].includes(entry[0]) || header === 'column'))); }
  showFile() {
    const card=(action,title,desc,image)=>`<button class="file-card" data-action="${action}">${icon(image)}<span><strong>${title}</strong><small>${desc}</small></span></button>`;
    this.openDialog('Your workspace', `<div class="file-hero"><div class="eyebrow">GRIDLINE / LOCAL FIRST</div><h3>Big ideas.<br>Beautifully organized.</h3><p>A spreadsheet that gives your numbers room to make sense.</p></div><div class="dialog-grid">${card('new','Blank workbook','Start with a clean sheet.','file')}${card('open','Open a workbook','Gridline, XLSX, CSV or TSV.','open')}${card('save','Save a copy','Preserve every Gridline feature.','save')}${card('export','Export your work','Excel workbook or CSV values.','export')}${card('sample','Explore the demo','A fictional revenue operations model.','table')}${card('help','Make yourself at home','Shortcuts, formulas and editing tips.','info')}</div><p><label><input id="reopen-last" type="checkbox" role="switch"${this.reopenLast ? ' checked' : ''}> Reopen last workbook on startup</label></p><p class="help-text" style="margin:20px 0 0">No account. No uploads. Workbook data is processed in your browser. Local autosave is specific to this browser and site.</p>`, 620);
    $('#reopen-last').onchange = e => this.errorBoundary(() => { localStorage.setItem(REOPEN_KEY,String(e.target.checked)); this.reopenLast = e.target.checked; });
  }
  showExport() {
    this.openDialog('Export workbook', `<div class="dialog-grid"><button class="file-card" data-action="save">${icon('save')}<span><strong>Gridline workbook</strong><small>.gridline · Full document fidelity, notes, chart definitions and rules.</small></span></button><button class="file-card" data-action="export-xlsx">${icon('table')}<span><strong>Excel workbook</strong><small>.xlsx · Cells, formulas, basic styles, merges, dimensions and frozen panes.</small></span></button><button class="file-card" data-action="export-csv">${icon('file')}<span><strong>CSV values</strong><small>.csv · Current sheet, values only. No formatting or formulas.</small></span></button><button class="file-card" data-action="print">${icon('print')}<span><strong>Print / Save as PDF</strong><small>Use your browser’s print dialog. Tabular sheet output.</small></span></button></div><p class="help-text">XLSX is a deliberately limited interoperability path, not a lossless Excel round-trip. Use .gridline to retain all features of this app. No files are uploaded.</p>`, 640);
  }
  showConditional() {
    if (!this.editable()) return;
    this.openDialog('Conditional formatting', `<p class="help-text">Apply a live rule to <b>${rangeAddress(this.selection)}</b>. Formatting updates when formula results change.</p><div class="dialog-grid"><button class="file-card" data-rule="bars">${icon('chart')}<span><strong>Data bars</strong><small>Compare the magnitude of each value.</small></span></button><button class="file-card" data-rule="scale">${icon('conditional')}<span><strong>Green color scale</strong><small>Shade cells from low to high.</small></span></button><button class="file-card" data-rule="positive">${icon('percent')}<span><strong>Positive / negative</strong><small>Green gains and red losses.</small></span></button><button class="file-card" data-rule="greater">${icon('check')}<span><strong>Greater than…</strong><small>Highlight values above a threshold.</small></span></button></div><label class="field-label">Threshold for “Greater than”</label><input id="cf-threshold" class="dialog-input" type="number" value="0"><div class="dialog-actions"><button id="clear-rules" class="secondary-btn">Clear sheet rules</button></div>`);
    $$('[data-rule]').forEach(button => button.onclick = () => this.errorBoundary(() => { const range = {...this.selection}; [...cellsIn(range)]; const rule={type:button.dataset.rule,range,value:Number($('#cf-threshold').value)}; this.workbook.mutate('Conditional formatting',()=>this.sheet.conditionalRules.push(rule)); this.closeDialog(); }));
    $('#clear-rules').onclick=()=>{this.workbook.mutate('Clear conditional rules',()=>this.sheet.conditionalRules=[]);this.closeDialog();};
  }
  showTextToColumns() {
    if (!this.editable()) return;
    const sheet = this.sheet, q = {...this.selection};
    if (q.c1 !== q.c2) throw new Error('Select one column to split.');
    if (q.r1 === q.r2) q.r2 = this.dataRange().r2;
    q.r2 = sheet.populatedRange(q).r2;
    if (q.r2 < q.r1 || q.r2-q.r1+1 > MAX_RANGE_CELLS) throw new Error('Select up to 200,000 source cells.');
    this.openDialog('Text to Columns', `<p class="help-text">Split ${rangeAddress(q)} using a delimiter. Dates split into month, day, and year with Slash, Hyphen, or Date parts. Text such as 8/3/???? keeps its unknown year. Formula results are converted to values.</p><label class="field-label" for="split-delimiter">Delimiter</label><select id="split-delimiter" class="dialog-input"><option value="/">Slash (/)</option><option value="-">Hyphen (-)</option><option value="date">Date parts (Month / Day / Year, including Excel numbers)</option><option value=",">Comma</option><option value="tab">Tab</option><option value=";">Semicolon</option><option value=" ">Space</option><option value="other">Other</option></select><input id="split-other" class="dialog-input" aria-label="Other delimiter" maxlength="1" placeholder="Custom delimiter" hidden><label class="field-label" for="split-destination">Destination</label><input id="split-destination" class="dialog-input" value="${address(q.r1,q.c1)}"><label class="field-label" for="split-format">Output format</label><select id="split-format" class="dialog-input"><option value="general">General</option><option value="text">Text (keep leading zeros)</option></select><label class="field-label">Preview (first 5 rows)</label><div id="split-preview" style="overflow:auto;max-height:180px"></div><div class="dialog-actions"><button class="secondary-btn" data-action="close-dialog">Cancel</button><button id="split-apply" class="primary-btn">Finish</button></div>`,620);
    const read = (limit = q.r2) => {
      const mode = $('#split-delimiter').value;
      const delimiter = mode === 'date' ? '/' : $('#split-delimiter').value === 'tab' ? '\t' : $('#split-delimiter').value === 'other' ? $('#split-other').value : $('#split-delimiter').value;
      if (delimiter.length !== 1) throw new Error('Enter one delimiter character.');
      const rows = [];
      for (let r=q.r1;r<=limit;r++) {
        const style = sheet.style(r,q.c1), fmt = numberFormatStyle(style.format).format, raw = this.workbook.value(sheet,r,q.c1);
        const date = typeof raw === 'number' && (mode === 'date' || fmt === 'date' && ['/', '-'].includes(delimiter));
        const value = date ? formatValue(raw, {format:'date', pattern:`m${delimiter}d${delimiter}yyyy`}) : ['date','time'].includes(fmt) ? this.workbook.display(sheet,r,q.c1) : raw;
        const parsed = parseDelimited(String(value ?? ''),delimiter);
        if (parsed.length !== 1) throw new Error('Line breaks inside a source cell must be enclosed in double quotes.');
        rows.push(parsed[0]);
      }
      return rows;
    };
    const preview = () => { $('#split-other').hidden = $('#split-delimiter').value !== 'other'; try { $('#split-preview').innerHTML='<table>'+read(Math.min(q.r2,q.r1+4)).map(row=>'<tr>'+row.map(v=>`<td style="border:1px solid #ccc;padding:6px">${escapeHTML(v)}</td>`).join('')+'</tr>').join('')+'</table>'; } catch(e) { $('#split-preview').textContent=e.message; } };
    const first = this.workbook.display(sheet,q.r1,q.c1);
    $('#split-delimiter').value = numberFormatStyle(sheet.style(q.r1,q.c1).format).format === 'date' || first.includes('/') ? '/' : first.includes('-') ? '-' : ',';
    $('#split-delimiter').onchange=preview; $('#split-other').oninput=preview; preview();
    $('#split-apply').onclick=()=>this.errorBoundary(()=>{
      const rows=read(), dest=parseAddress($('#split-destination').value.trim()), width=rows.reduce((n,row)=>Math.max(n,row.length),0), text=$('#split-format').value==='text';
      if (!dest || dest.r+rows.length>MAX_ROWS || dest.c+width>MAX_COLS || rows.length*width>MAX_RANGE_CELLS) throw new Error('Destination must fit within the sheet and the 200,000-cell limit.');
      const target={r1:dest.r,c1:dest.c,r2:dest.r+rows.length-1,c2:dest.c+width-1};
      if (sheet.merges.some(m=>m.r1<=target.r2 && m.r2>=target.r1 && m.c1<=target.c2 && m.c2>=target.c1)) throw new Error('Unmerge destination cells before splitting.');
      let overwrite=false;
      for(const {r,c} of cellsIn(target)) if(sheet.raw(r,c) && !(c===q.c1 && r>=q.r1 && r<=q.r2)) overwrite=true;
      const apply=()=>{this.workbook.transaction('Text to Columns',()=>rows.forEach((row,i)=>{for(let j=0;j<width;j++){const raw=row[j]??'';this.workbook.setCell(sheet,dest.r+i,dest.c+j,{raw:raw && (text || raw.startsWith('=') || raw.startsWith("'")) ? "'"+raw : raw,style:{format:text?'text':'general',decimals:null}});}}));this.closeDialog();this.select(target,dest,true);};
      if(overwrite) this.confirm('Replace destination data?', `Text to Columns will overwrite existing data in ${rangeAddress(target)}.`,apply,'Replace data'); else apply();
    });
  }
  showSort() {
    if(!this.editable())return;const q=this.dataRange();if(q.r2<=q.r1){this.toast('Select a table or range with data rows.');return;}
    const options=[];for(let c=q.c1;c<=q.c2;c++)options.push(`<option value="${c}"${c===this.active.c?' selected':''}>${colName(c)} — ${escapeHTML(this.workbook.display(this.sheet,q.r1,c))}</option>`);
    this.openDialog('Sort range',`<p class="help-text">Sort complete rows in <b>${rangeAddress(q)}</b>. Formula references move relative to their cells.</p><div id="sort-levels"><div class="sort-level"><label class="field-label">Sort by column</label><select id="sort-column" class="dialog-input">${options.join('')}</select><label class="field-label">Order</label><select id="sort-order" class="dialog-input"><option value="asc">A to Z / Smallest to largest</option><option value="desc">Z to A / Largest to smallest</option></select></div></div><button id="sort-add" class="secondary-btn">Add level</button><p><label><input type="checkbox" id="sort-header" checked> My data has headers</label></p><div class="dialog-actions"><button class="secondary-btn" data-action="close-dialog">Cancel</button><button id="sort-apply" class="primary-btn">Sort range</button></div>`);
    $('#sort-add').onclick=()=>{ const row=document.createElement('div'); row.className='sort-level'; row.innerHTML=`<label class="field-label">Then by</label><select class="dialog-input">${options.join('')}</select><select class="dialog-input"><option value="asc">A to Z / Smallest to largest</option><option value="desc">Z to A / Largest to smallest</option></select><button class="secondary-btn" type="button">Delete level</button>`; row.querySelector('button').onclick=()=>row.remove(); $('#sort-levels').append(row); };
    $('#sort-apply').onclick=()=>this.errorBoundary(()=>{this.workbook.sort(this.sheet,q,$$('.sort-level').map(row=>({column:+row.querySelectorAll('select')[0].value,descending:row.querySelectorAll('select')[1].value==='desc'})),false,$('#sort-header').checked);this.closeDialog();});
  }
  showFilter() {
    const q=this.dataRange();if(q.r2<=q.r1){this.toast('Select a table with a header row to filter.');return;}
    const c=Math.max(q.c1,Math.min(q.c2,this.active.c)), values=[...new Set(Array.from({length:Math.min(q.r2-q.r1,10000)},(_,i)=>this.workbook.display(this.sheet,q.r1+1+i,c)))].sort((a,b)=>a.localeCompare(b));
    if(values.length>2000){this.toast('This column has more than 2,000 distinct values. Select a smaller range.',true);return;}
    const current=this.sheet.filters?.criteria?.[c];
    this.openDialog(`Filter ${colName(c)} — ${this.workbook.display(this.sheet,q.r1,c)}`,`<input id="filter-search" class="dialog-input" placeholder="Search values…" autofocus><div class="panel-actions"><button class="secondary-btn" id="filter-all">Select all</button><button class="secondary-btn" id="filter-none">Select none</button></div><div class="filter-values">${values.map((v,i)=>`<label data-filter-value="${i}"><input type="checkbox" value="${i}" ${!current||current.includes(v)?'checked':''}><span>${escapeHTML(v||'(Blanks)')}</span></label>`).join('')}</div><p class="help-text">The first row remains a header. Multiple column filters combine with AND. SUM includes hidden rows.</p><div class="dialog-actions"><button class="secondary-btn" data-action="close-dialog">Cancel</button><button class="primary-btn" id="filter-apply">Apply filter</button></div>`);
    $('#filter-search').oninput=e=>$$('[data-filter-value]').forEach(label=>label.hidden=!values[+label.dataset.filterValue].toLowerCase().includes(e.target.value.toLowerCase()));
    $('#filter-all').onclick=()=>$$('.filter-values input').forEach(i=>i.checked=true);$('#filter-none').onclick=()=>$$('.filter-values input').forEach(i=>i.checked=false);
    $('#filter-apply').onclick=()=>this.errorBoundary(()=>{
      const selected=$$('.filter-values input:checked').map(i=>values[+i.value]);
      this.workbook.mutate('Filter rows',()=>{const criteria={...(this.sheet.filters?.criteria||{}),[c]:selected};this.sheet.filters={range:q,criteria};this.sheet.hiddenRows.clear();for(let r=q.r1+1;r<=q.r2;r++)if(Object.entries(criteria).some(([col,allowed])=>!allowed.includes(this.workbook.display(this.sheet,r,+col))))this.sheet.hiddenRows.add(r);});
      this.closeDialog();this.toast(`${q.r2-q.r1-this.sheet.hiddenRows.size} of ${q.r2-q.r1} rows visible.`);
    });
  }
  renameSheet() {
    this.openDialog('Rename worksheet',`<label class="field-label">Worksheet name</label><input id="sheet-name-input" class="dialog-input" value="${escapeHTML(this.sheet.name)}" maxlength="31" autofocus><div class="dialog-actions"><button class="secondary-btn" data-action="close-dialog">Cancel</button><button id="sheet-rename-apply" class="primary-btn">Rename</button></div>`);
    $('#sheet-rename-apply').onclick=()=>this.errorBoundary(()=>{this.workbook.renameSheet(this.sheet,$('#sheet-name-input').value);this.closeDialog();});
    $('#sheet-name-input').onkeydown=e=>{if(e.key==='Enter')$('#sheet-rename-apply').click();};
  }
  showSheets() {
    this.openDialog('Worksheets',`<div class="command-list">${this.workbook.sheets.map(s=>`<button data-go-sheet="${escapeHTML(s.id)}">${icon('table')}${escapeHTML(s.name)}<small>${s.cells.size.toLocaleString()} stored cells</small></button>`).join('')}</div><div class="dialog-actions"><button class="primary-btn" id="sheets-add">New worksheet</button></div>`);
    $$('[data-go-sheet]').forEach(b=>b.onclick=()=>{this.switchSheet(b.dataset.goSheet);this.closeDialog();});$('#sheets-add').onclick=()=>{this.closeDialog();this.run('add-sheet');};
  }
  showFunctions() {
    this.openDialog('Function library',`<input id="function-search" class="dialog-input" placeholder="Find a function — SUM, IF, XLOOKUP…" autofocus><p class="help-text">${FUNCTIONS.size} function names supported. Select one to start editing the active cell. Arguments use commas; ranges use A1:B10 notation.</p><div id="function-list" class="command-list"></div>`);
    const render=()=>{const query=$('#function-search').value.toLowerCase();$('#function-list').innerHTML=[...FUNCTIONS].filter(([name,fn])=>name.toLowerCase().includes(query)||fn.description.toLowerCase().includes(query)).map(([name,fn])=>`<button data-function="${name}">${icon('function')}<b style="min-width:100px;text-align:left">${name}</b><small>${escapeHTML(fn.description)}</small></button>`).join('');$$('[data-function]').forEach(b=>b.onclick=()=>{this.closeDialog();this.startEdit(`=${b.dataset.function}(`);});};
    $('#function-search').oninput=render;render();
  }
  showCommands() {
    this.openDialog('Find a command',`<input id="command-search" class="dialog-input" placeholder="What would you like to do?" autofocus><div id="command-list" class="command-list"></div>`,590);
    const render=()=>{const query=$('#command-search').value.toLowerCase();$('#command-list').innerHTML=COMMANDS.filter(c=>c[1].toLowerCase().includes(query)).map(c=>`<button data-command="${c[0]}">${icon(c[2])}${escapeHTML(c[1])}<small>${c[3]}</small></button>`).join('');$$('[data-command]').forEach(b=>b.onclick=()=>{this.closeDialog();this.errorBoundary(()=>this.run(b.dataset.command));});};$('#command-search').oninput=render;render();
  }
  showNames() {
    const qualified=`'${this.sheet.name.replaceAll("'","''")}'!${rangeAddress(this.selection)}`;
    this.openDialog('Named ranges',`<p class="help-text">Use names in formulas, for example <b>=RevenueTarget*1.2</b>.</p><div class="result-list">${Object.entries(this.workbook.names).map(([name,ref])=>`<div class="metric-line"><strong>${escapeHTML(name)}</strong><span>${escapeHTML(ref)}</span></div>`).join('')||'<p class="help-text">No named ranges yet.</p>'}</div><label class="field-label">Name</label><input id="range-name" class="dialog-input" placeholder="RevenueTarget" autofocus><label class="field-label">Refers to</label><input id="range-reference" class="dialog-input" value="${escapeHTML(qualified)}"><div class="dialog-actions"><button class="secondary-btn" data-action="close-dialog">Close</button><button class="primary-btn" id="name-define">Define name</button></div>`,600);
    $('#name-define').onclick=()=>this.errorBoundary(()=>{const name=$('#range-name').value.trim().toUpperCase(),ref=$('#range-reference').value.trim().replace(/^=/,'');if(!/^[A-Z_][A-Z0-9_]{0,63}$/.test(name)||parseAddress(name)||['TRUE','FALSE'].includes(name))throw new Error('Use a name beginning with a letter or underscore, not a cell address.');const match=/^(?:'((?:[^']|'')+)'|([^!]+))!([^!]+)$/.exec(ref);if(!match||!this.workbook.sheetByName((match[1]||match[2]).replaceAll("''","'"))||!parseRange(match[3]))throw new Error('Use a valid sheet-qualified reference, for example Sales!B2:B20.');this.workbook.mutate('Define name',()=>this.workbook.names[name]=ref);this.showNames();});
  }
  openPanel(type,title,body) {this.panelType=type;const panel=$('#side-panel');panel.hidden=false;panel.innerHTML=`<div class="panel-heading"><span>${title}</span><button data-action="close-panel" aria-label="Close panel">×</button></div>${body}`;this.renderer.resize();}
  closePanel() {this.panelType=null;$('#side-panel').hidden=true;this.renderer.resize();this.host.focus();}
  showFind() {
    const scope = {...this.selection}, column = this.active.c, sheetId = this.sheet.id;
    const includes = (r,c) => this.sheet.id === sheetId && ($('#find-scope').value === 'sheet' || ($('#find-scope').value === 'column' ? c === column : r >= scope.r1 && r <= scope.r2 && c >= scope.c1 && c <= scope.c2));
    this.openPanel('find','Find & replace',`<label class="field-label" for="find-scope">Within</label><select id="find-scope" class="panel-input"><option value="sheet">Sheet</option><option value="column">Column ${colName(column)}</option><option value="selection">Selection ${rangeAddress(scope)}</option></select><label class="field-label" for="find-look-in">Look in</label><select id="find-look-in" class="panel-input"><option value="display">Displayed values</option><option value="raw">Original values and formulas</option></select><label class="field-label">Find</label><input id="find-query" class="panel-input" placeholder="Search…"><label class="field-label">Replace with</label><input id="replace-query" class="panel-input" placeholder="Replacement text"><p><label><input id="find-match-case" type="checkbox"> Match Case</label></p><p><label><input id="find-entire" type="checkbox"> Find entire cells only</label></p><div class="panel-actions"><button class="primary-btn" id="find-next">Find next</button><button class="secondary-btn" id="replace-all">Replace all</button></div><p class="help-text">Search uses displayed values by default. Choose original values and formulas to search stored numbers or formula text. Replace in Displayed values writes the changed text, converting matching dates and formula results to text. Original values and formulas edits stored input. Match Case and Find entire cells only apply to both Find and Replace. The scope stays fixed while navigating results. Results are limited to 200; Replace all processes every match in scope.</p><div id="find-count" class="badge">Enter a search term</div><div class="result-list" id="find-results"></div>`);
    const pattern = () => { const query = $('#find-query').value.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); return new RegExp($('#find-entire').checked ? `^(?:${query})(?![\\s\\S])` : query,$('#find-match-case').checked ? '' : 'i'); };
    let matches=[],next=-1;
    const search=()=>{const query=$('#find-query').value,match=pattern();matches=[];if(query)for(const [key,cell]of this.sheet.cells){const [r,c]=key.split(',').map(Number);if(!includes(r,c))continue;const display=this.workbook.display(this.sheet,r,c);const text=$('#find-look-in').value==='raw'?cell.raw:display;if(match.test(text)){matches.push({r,c,display:text});if(matches.length>=200)break;}}matches.sort((a,b)=>a.r-b.r||a.c-b.c);next=-1;$('#find-count').textContent=query?`${matches.length}${matches.length===200?'+':''} matches`:'Enter a search term';$('#find-results').innerHTML=matches.map((m,i)=>`<button class="find-result" data-result="${i}"><b>${address(m.r,m.c)}</b><span>${escapeHTML(m.display.slice(0,120))}</span></button>`).join('');$$('[data-result]').forEach(b=>b.onclick=()=>{const m=matches[+b.dataset.result];this.goto(m.r,m.c);});};
    $('#find-scope').value = scope.r1 !== scope.r2 || scope.c1 !== scope.c2 ? 'selection' : 'sheet'; $('#find-scope').onchange=search; $('#find-look-in').onchange=search; $('#find-match-case').onchange=search; $('#find-entire').onchange=search;
    $('#find-query').oninput=search;$('#find-next').onclick=()=>{if(matches.length){const m=matches[++next%matches.length];this.goto(m.r,m.c);}};
    $('#find-query').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();$('#find-next').click();}};
    $('#replace-all').onclick=()=>this.errorBoundary(()=>{if(!this.editable())return;const query=$('#find-query').value,replace=$('#replace-query').value,match=pattern();if(!query)return;let count=0;this.workbook.transaction('Replace all',()=>{for(const [key,cell]of this.sheet.cells){const [r,c]=key.split(',').map(Number);if(!includes(r,c))continue;const displayed=$('#find-look-in').value==='display',text=displayed?this.workbook.display(this.sheet,r,c):cell.raw;if(!match.test(text))continue;const changed=text.replace(new RegExp(match.source,match.flags+'g'),()=>replace);this.workbook.setRaw(this.sheet,r,c,displayed&&changed ? "'"+changed : changed);count++;}});this.toast(`Replaced matches in ${count} cells.`);search();});$('#find-query').focus();
  }
  showNotes() {
    const notes=[];for(const [key,cell]of this.sheet.cells)if(cell.note){const[r,c]=key.split(',').map(Number);notes.push({r,c,text:cell.note});}
    this.openPanel('notes','Worksheet notes',`<button class="primary-btn" data-action="add-note">${icon('plus')} Add note at ${address(this.active.r,this.active.c)}</button><p class="help-text">Local cell annotations. Notes are stored in .gridline workbooks, not shared with other people.</p>${notes.map(note=>`<div class="note-card"><div class="note-card-header"><button data-note-location="${address(note.r,note.c)}">${address(note.r,note.c)}</button><button data-action="delete-note" data-row="${note.r}" data-col="${note.c}" title="Delete note">×</button></div><p>${escapeHTML(note.text)}</p></div>`).join('')||'<div class="empty-panel">A little context goes a long way.<br>Add a note to any cell.</div>'}`);
    $$('[data-note-location]').forEach(b=>b.onclick=()=>{const p=parseAddress(b.dataset.noteLocation);this.goto(p.r,p.c);});
  }
  addNote() {
    if(!this.editable())return;const {r,c}=this.active;
    this.openDialog(`Note at ${address(r,c)}`,`<textarea id="note-text" class="dialog-input" style="height:155px;resize:vertical" maxlength="10000" autofocus placeholder="Add context, a reminder, or a question…">${escapeHTML(this.sheet.get(r,c)?.note||'')}</textarea><p class="help-text">This note is local to your workbook.</p><div class="dialog-actions"><button class="secondary-btn" data-action="close-dialog">Cancel</button><button class="primary-btn" id="note-save">Save note</button></div>`);
    $('#note-save').onclick=()=>{this.workbook.setCell(this.sheet,r,c,{note:$('#note-text').value.trim()||undefined});this.closeDialog();this.showNotes();};
  }
  showCellInspector() {
    const cell=this.sheet.get(this.active.r,this.active.c),value=this.workbook.value(this.sheet,this.active.r,this.active.c),id=this.workbook.engine.id(this.sheet,this.active.r,this.active.c),deps=[...(this.workbook.engine.dependencies.get(id)||[])];
    this.openPanel('inspect','Cell inspector',`<div class="eyebrow">${escapeHTML(this.sheet.name)} / ${address(this.active.r,this.active.c)}</div><label class="field-label">Original input</label><pre style="white-space:pre-wrap;overflow-wrap:anywhere;font-size:12px">${escapeHTML(cell?.raw||'(empty)')}</pre><label class="field-label">Calculated value</label><pre style="white-space:pre-wrap">${escapeHTML(value instanceof FormulaError?value.code:value??'(blank)')}</pre>${value instanceof FormulaError?`<p class="help-text">${escapeHTML(value.message)}</p>`:''}<div class="metric-line"><span>Value type</span><strong>${value instanceof FormulaError?'error':value===null?'blank':typeof value}</strong></div><div class="metric-line"><span>Direct dependencies</span><strong>${deps.length}</strong></div><label class="field-label">References read during evaluation</label><div class="help-text">${deps.slice(0,100).map(d=>{const [sid,key]=d.split('!'),[r,c]=key.split(',').map(Number);return escapeHTML((this.workbook.sheets.find(s=>s.id===sid)?.name||sid)+'!'+address(r,c));}).join('<br>')||'No cell dependencies.'}</div>`);
  }
  showPerformance() {
    this.openPanel('performance','Engine diagnostics',`<div class="badge">${this.renderer.backend==='webgpu'?'WEBGPU / INSTANCED QUADS':'CANVAS2D / FALLBACK'}</div><p class="help-text">${this.renderer.backend==='webgpu'?'Cells and glyph masks are rendered through one ordered instanced draw call. Only the visible grid is drawn.':'The same viewport display list is rendered through Canvas2D because WebGPU is unavailable or was disabled.'}</p><div id="perf-metrics"></div><p class="help-text">Frame time below is CPU preparation and command submission, not GPU completion time or an FPS benchmark. Logical grid capacity does not imply tested million-row calculation throughput.</p><button class="secondary-btn" id="perf-test">Create a 10,000-row test sheet</button><p class="help-text">Adds 80,008 stored cells with formulas. Your existing worksheets are kept.</p>${this.renderer.fallbackReason?`<p class="help-text">Fallback reason: ${escapeHTML(this.renderer.fallbackReason)}</p>`:''}`);
    $('#perf-test').onclick=()=>this.errorBoundary(()=>{const before=performance.now();this.workbook.mutate('Create performance sheet',()=>{const name=this.workbook.uniqueSheetName('Performance test');const template=new Workbook().activeSheet;template.name=name;template.freezeRows=1;this.workbook.sheets.push(template);this.workbook.activeSheetId=template.id;['Index','Units','Price','Revenue','Cost','Profit','Margin','Contribution'].forEach((v,c)=>template.cells.set(keyOf(0,c),{raw:v,style:{bold:true,fill:'#176b4a',color:'#ffffff'}}));for(let r=1;r<=10000;r++){const n=r+1;const row=[r,1+r%100,25+r%75,`=B${n}*C${n}`,`=D${n}*0.62`,`=D${n}-E${n}`,`=IFERROR(F${n}/D${n},0)`,`=D${n}*0.1`];row.forEach((v,c)=>template.cells.set(keyOf(r,c),{raw:String(v),style:{format:c===6?'percent':c>=3?'number':'general'}}));}});this.renderer.scrollX=this.renderer.scrollY=0;this.goto(0,0);this.toast(`Created 10,000 rows in ${(performance.now()-before).toFixed(0)} ms on this browser.`);});
    this.refreshPerformance();
  }
  refreshPerformance() {
    const metrics=$('#perf-metrics');if(!metrics)return;const formulaCount=[...this.sheet.cells.values()].filter(c=>c.raw.startsWith('=')).length;
    const lines=[['Backend',this.renderer.backend],['Logical sheet','1,048,576 × 16,384'],['Stored cells',this.sheet.cells.size.toLocaleString()],['Formula cells',formulaCount.toLocaleString()],['Visible cells',this.renderer.visibleCellCount],['Last frame CPU',this.renderer.lastFrameMs.toFixed(2)+' ms'],['Instanced quads',this.renderer.backend==='webgpu'?this.renderer.instanceCount.toLocaleString():'—'],['Glyphs cached',this.renderer.atlas.map.size],['Cached formula results',this.workbook.engine.cache.size],['Cell evaluations',this.workbook.engine.evaluations.toLocaleString()]];
    metrics.innerHTML=lines.map(([label,value])=>`<div class="metric-line"><span>${label}</span><strong>${escapeHTML(value)}</strong></div>`).join('');
  }
  showHelp() {
    const shortcuts=[['Navigate / extend selection','Arrow keys / Shift + arrows'],['Jump to data edge','Ctrl/⌘ + arrow'],['First cell / last used cell','Ctrl/⌘ + Home / End'],['Edit active cell','F2 or start typing'],['Apply edit / cancel','Enter / Escape'],['Move across cells','Tab / Shift + Tab'],['Insert line break while editing','Alt + Enter'],['Cycle absolute formula reference','F4 while editing'],['Copy / cut / paste','Ctrl/⌘ + C / X / V'],['Undo / redo','Ctrl/⌘ + Z / Shift + Z'],['Fill down / right','Ctrl/⌘ + D / R'],['Bold / italic / underline','Ctrl/⌘ + B / I / U'],['Find and replace','Ctrl/⌘ + F'],['AutoSum','Alt + ='],['Command search','Ctrl/⌘ + K'],['Export full-fidelity workbook','Ctrl/⌘ + S'],['Open a workbook','Ctrl/⌘ + O'],['Print','Ctrl/⌘ + P'],['Show formula input','Ctrl/⌘ + `']];
    this.openDialog('Make yourself at home',`<p class="help-text">Double-click a cell to edit. Drag to select a range. Drag the small green square at the selection’s lower-right corner to fill. Two numeric seed cells create a sequence. Resize columns and rows by dragging header boundaries; double-click a column boundary to auto-fit.</p><table class="help-table">${shortcuts.map(([label,key])=>`<tr><td>${label}</td><td><kbd>${key}</kbd></td></tr>`).join('')}</table><h3>A few formulas to try</h3><p class="help-text"><code>=SUM(D12:F12)</code><br><code>=IF(G12&gt;=H12,"Above target","Needs focus")</code><br><code>=XLOOKUP("Enterprise",B12:B19,G12:G19)</code><br><code>='Sales data'!H2 * Assumptions!$B$3</code></p>`,680);
  }
  showAbout() {
    this.openDialog('Gridline',`<div class="file-hero"><div class="eyebrow">VERSION 0.1 / ENGINEERING PREVIEW</div><h3>A clearer way to work.</h3><p>A working, local-first spreadsheet built with plain JavaScript, HTML and CSS. No application framework, cloud backend, runtime dependencies, or evaluation of formula strings as JavaScript.</p></div><h3>Under the hood</h3><p class="help-text">Sparse cell storage, ${FUNCTIONS.size} registered function names, a Pratt-parser formula AST, dependency invalidation, transactional cell edits, undo/redo, viewport culling, frozen panes, and an instanced WebGPU renderer with a cached glyph-mask atlas. Native text inputs handle editing.</p><h3>Boundaries of this build</h3><p class="help-text">This is not feature-equivalent to Microsoft Excel. VBA, Power Query, pivot tables, dynamic-array spills, collaboration, advanced print layout, rich cell text, full international text shaping and complete Excel formula/file compatibility are not implemented. The grid has a 200,000-cell range-operation limit and a 256-cell calculation-stack limit. Large snapshots and calculation run on the main thread. Charts use SVG overlays. Basic XLSX exchange is not lossless; use .gridline for full app fidelity.</p><p class="help-text">Sheet read-only mode is a UI editing guard, not security or encryption. Cut/paste moves values and formulas but does not retarget references from other cells. Structural row/column edits clear charts and conditional rules on that sheet to avoid stale range metadata.</p><p class="help-text">Gridline is an independent implementation with an Excel-inspired interface. It is not affiliated with Microsoft. All demo business data is fictional.</p>`,670);
  }
  showChart(type='column') {
    if(!this.editable())return;const q=this.dataRange();
    this.openDialog('Insert a chart',`<label class="field-label">Chart title</label><input id="chart-title-input" class="dialog-input" value="${type==='line'?'A view of the trend':'Your data, visualized'}" autofocus><label class="field-label">Chart type</label><select id="chart-type-input" class="dialog-input">${[['column','Column'],['line','Line'],['bar','Horizontal bar'],['donut','Doughnut']].map(([v,t])=>`<option value="${v}"${v===type?' selected':''}>${t}</option>`).join('')}</select><label class="field-label">Source range</label><input id="chart-range-input" class="dialog-input" value="${rangeAddress(q)}"><p class="help-text">The first row supplies headers. The first column supplies category labels; the second supplies numeric values. Up to 100 data points are shown. Changes to source cells update the chart.</p><div class="dialog-actions"><button class="secondary-btn" data-action="close-dialog">Cancel</button><button id="chart-create" class="primary-btn">Insert chart</button></div>`);
    $('#chart-create').onclick=()=>this.errorBoundary(()=>{const range=parseRange($('#chart-range-input').value);if(!range||range.r2<=range.r1||range.c2<=range.c1)throw new Error('Choose a range with at least two columns and a header plus data row.');const chart={id:'chart-'+Date.now().toString(36),title:$('#chart-title-input').value.trim()||'Chart',subtitle:rangeAddress(range),type:$('#chart-type-input').value,range,row:Math.min(MAX_ROWS-20,range.r2+3),col:range.c1,width:550,height:300};this.workbook.mutate('Insert chart',()=>this.sheet.charts.push(chart));this.closeDialog();this.goto(chart.row,chart.col);});
  }
  chartData(chart) {
    const q=chart.range;
    if(chart.aggregate){const labels=[],values=[];for(let c=q.c1;c<=q.c2;c++){labels.push(this.workbook.display(this.sheet,q.r1,c));let total=0;for(let r=q.r1+1;r<=Math.min(q.r2,q.r1+10000);r++){const value=this.workbook.value(this.sheet,r,c);if(typeof value==='number')total+=value;}values.push(total);}return{labels,values};}
    const labels=[],values=[];for(let r=q.r1+1;r<=Math.min(q.r2,q.r1+100);r++){const value=this.workbook.value(this.sheet,r,chart.valueColumn??q.c1+1);if(typeof value==='number'){labels.push(this.workbook.display(this.sheet,r,chart.labelColumn??q.c1));values.push(value);}}return{labels,values};
  }
  chartSVG(chart) {
    const {labels,values}=this.chartData(chart),W=550,H=Math.max(130,chart.height-60),muted=this.renderer.dark?'#97afa0':'#8ca093',text=this.renderer.dark?'#d6e5da':'#597363';
    if(!values.length)return `<svg class="chart-svg" viewBox="0 0 ${W} ${H}" role="img" aria-label="No numeric chart data"><text x="${W/2}" y="${H/2}" text-anchor="middle" fill="${muted}" font-size="12">Select a label column and a numeric value column.</text></svg>`;
    const short=n=>Math.abs(n)>=1000000?(n/1000000).toFixed(1)+'m':Math.abs(n)>=1000?(n/1000).toFixed(0)+'k':String(Math.round(n));
    const palette=['#247651','#559b73','#8dbaa0','#bad5c0','#d9e7ce','#9ba77c','#cfb572','#a9c2b0'];let body='';
    if(chart.type==='bar'){
      const n=Math.min(8,values.length),max=Math.max(...values.map(Math.abs),1),rowH=(H-20)/n;
      for(let i=0;i<n;i++){const y=9+i*rowH;body+=`<text x="20" y="${y+rowH*.61}" font-size="10" fill="${text}">${escapeHTML(labels[i].slice(0,18))}</text><rect x="138" y="${y+rowH*.22}" width="${330*Math.abs(values[i])/max}" height="${rowH*.48}" rx="2" fill="${palette[i%palette.length]}"/><text x="520" y="${y+rowH*.61}" text-anchor="end" font-size="10" fill="${text}">${short(values[i])}</text>`;}
    }else if(chart.type==='donut'){
      const positive=values.map(v=>Math.max(0,v)),total=positive.reduce((a,b)=>a+b,0)||1;let angle=-Math.PI/2;const cx=H*.6,cy=H/2,r=Math.min(74,H*.4);
      positive.forEach((v,i)=>{const end=angle+v/total*Math.PI*2,x1=cx+r*Math.cos(angle),y1=cy+r*Math.sin(angle),x2=cx+r*Math.cos(end),y2=cy+r*Math.sin(end);if(v/total>.999)body+=`<circle cx="${cx}" cy="${cy}" r="${r}" stroke="${palette[i%palette.length]}" stroke-width="27" fill="none"/>`;else if(v)body+=`<path d="M${x1},${y1} A${r},${r} 0 ${end-angle>Math.PI?1:0} 1 ${x2},${y2}" fill="none" stroke="${palette[i%palette.length]}" stroke-width="27"/>`;angle=end;if(i<8)body+=`<rect x="260" y="${20+i*21}" width="8" height="8" rx="2" fill="${palette[i%palette.length]}"/><text x="280" y="${28+i*21}" fill="${text}" font-size="10">${escapeHTML(labels[i].slice(0,24))}</text><text x="520" y="${28+i*21}" fill="${text}" text-anchor="end" font-size="10">${(v/total*100).toFixed(1)}%</text>`;});body+=`<text x="${cx}" y="${cy+5}" text-anchor="middle" font-size="22" font-weight="600" fill="${text}">${short(total)}</text>`;
    }else{
      const x0=58,y0=H-28,top=21,plotW=467,plotH=y0-top,lo=Math.min(0,...values),hi=Math.max(1,...values)*1.12,span=hi-lo,y=v=>y0-(v-lo)/span*plotH,zero=y(0),step=plotW/values.length;
      for(let i=0;i<=3;i++){const value=lo+span*i/3,py=y(value);body+=`<line x1="${x0}" x2="${W-20}" y1="${py}" y2="${py}" stroke="${this.renderer.dark?'#344c3d':'#eaf0eb'}"/><text x="${x0-10}" y="${py+3}" text-anchor="end" font-size="9" fill="${muted}">${short(value)}</text>`;}
      if(chart.type==='line'){const points=values.map((v,i)=>`${x0+step*(i+.5)},${y(v)}`);body+=`<path d="M${x0+step*.5},${zero} L${points.join(' L')} L${x0+step*(values.length-.5)},${zero} Z" fill="#39875e" opacity=".08"/><polyline points="${points.join(' ')}" fill="none" stroke="#36845c" stroke-width="2.5"/>`;values.forEach((v,i)=>body+=`<circle cx="${x0+step*(i+.5)}" cy="${y(v)}" r="3.5" fill="#36845c" stroke="white" stroke-width="2"/>`);}
      else values.forEach((v,i)=>{const width=Math.min(78,step*.54);body+=`<rect x="${x0+step*(i+.5)-width/2}" y="${Math.min(zero,y(v))}" width="${width}" height="${Math.abs(zero-y(v))}" rx="3" fill="${i===values.length-1?'#247651':'#b3cdbb'}"/>`;if(values.length<=12)body+=`<text x="${x0+step*(i+.5)}" y="${y(v)-8}" text-anchor="middle" font-size="11" font-weight="600" fill="${text}">${short(v)}</text>`;});
      labels.forEach((label,i)=>{if(values.length<=12||i%Math.ceil(values.length/12)===0)body+=`<text x="${x0+step*(i+.5)}" y="${H-8}" text-anchor="middle" font-size="10" fill="${muted}">${escapeHTML(label.slice(0,13))}</text>`;});
    }
    return `<svg class="chart-svg" viewBox="0 0 ${W} ${H}" role="img" aria-label="${escapeHTML(chart.title+': '+labels.map((l,i)=>l+' '+values[i]).join(', '))}">${body}</svg>`;
  }
  renderCharts() {
    const layer=$('#chart-layer');layer.innerHTML=this.sheet.charts.map(chart=>`<div class="chart-card" id="${escapeHTML(chart.id)}" data-chart-id="${escapeHTML(chart.id)}"><div class="chart-header"><div><div class="chart-title">${escapeHTML(chart.title)}</div><div class="chart-subtitle">${escapeHTML(chart.subtitle||rangeAddress(chart.range))}</div></div><button data-action="remove-chart" data-chart="${escapeHTML(chart.id)}" title="Remove chart" aria-label="Remove chart">×</button></div>${this.chartSVG(chart)}</div>`).join('');
    layer.querySelectorAll('.chart-header').forEach(header=>header.onpointerdown=e=>{
      if(e.target.closest('button')||!this.editable())return;e.preventDefault();e.stopPropagation();const card=header.closest('[data-chart-id]'),chart=this.sheet.charts.find(c=>c.id===card.dataset.chartId),startX=e.clientX,startY=e.clientY,ox=chart.offsetX||0,oy=chart.offsetY||0;let nx=ox,ny=oy;header.setPointerCapture(e.pointerId);
      const move=event=>{nx=ox+(event.clientX-startX)/this.renderer.zoom;ny=oy+(event.clientY-startY)/this.renderer.zoom;chart.offsetX=nx;chart.offsetY=ny;this.positionCharts();};
      const up=()=>{header.removeEventListener('pointermove',move);header.removeEventListener('pointerup',up);header.removeEventListener('pointercancel',up);chart.offsetX=ox;chart.offsetY=oy;this.workbook.mutate('Move chart',()=>{chart.offsetX=nx;chart.offsetY=ny;});};header.addEventListener('pointermove',move);header.addEventListener('pointerup',up);header.addEventListener('pointercancel',up);
    });this.positionCharts();
  }
  positionCharts() {
    const z=this.renderer.zoom, frozen=this.renderer.frozenSize(); $('#chart-layer').style.clipPath=`inset(${frozen.y}px 0 0 ${frozen.x}px)`;
    for(const chart of this.sheet.charts){const el=document.getElementById(chart.id);if(!el)continue;const rect=this.renderer.cellRect(chart.row,chart.col,false),x=rect.x-this.renderer.headerW+(chart.offsetX||0)*z,y=rect.y-this.renderer.headerH+(chart.offsetY||0)*z,w=chart.width*z,h=chart.height*z;el.style.left=x+'px';el.style.top=y+'px';el.style.width=w+'px';el.style.height=h+'px';el.style.display=x+w<0||y+h<0||x>this.renderer.width||y>this.renderer.height?'none':'block';}
  }
  printSheet(selected=false) {
    this.closeDialog();const q=selected?this.selection:this.sheet.usedRange();if((q.r2-q.r1+1)*(q.c2-q.c1+1)>10000)throw new Error('Print supports at most 10,000 cells. Select a smaller range.');
    let rows='';for(let r=q.r1;r<=q.r2;r++){if(this.sheet.hiddenRows.has(r))continue;let cells='';for(let c=q.c1;c<=q.c2;c++){if(this.sheet.hiddenCols.has(c))continue;const merge=this.sheet.mergeAt(r,c);if(merge&&(r!==merge.r1||c!==merge.c1))continue;const style=this.sheet.style(r,c);const safeColor=color=>/^#[0-9a-f]{6}$/i.test(color)?color:'inherit';cells+=`<td${merge?` rowspan="${Math.min(merge.r2,q.r2)-r+1}" colspan="${Math.min(merge.c2,q.c2)-c+1}"`:''} style="background:${safeColor(style.fill)};color:${safeColor(style.color)};font-weight:${style.bold?'bold':'normal'};text-align:${['left','center','right'].includes(style.align)?style.align:typeof this.workbook.value(this.sheet,r,c)==='number'?'right':'left'}">${escapeHTML(this.workbook.display(this.sheet,r,c))}</td>`;}rows+='<tr>'+cells+'</tr>';}
    $('#print-area').innerHTML=`<h1>${escapeHTML(this.workbook.title)}</h1><p class="print-meta">${escapeHTML(this.sheet.name)} · ${rangeAddress(q)} · Printed from Gridline</p><table>${rows}</table>`;window.print();
  }
}
const app = new GridlineApp();
// Intentional diagnostics/embedding API. The model remains independent of the view.
window.gridline = app;

window.Gridline = Object.freeze({ version: '0.1.0', Workbook, GridRenderer, exportXLSX, importXLSX, parseDelimited, serializeDelimited, exportCSV, createSampleWorkbook });
