/** Instanced WebGPU renderer: solid quads + cached glyph-mask atlas, one ordered draw call.
 * Canvas2D fallback consumes the same retained display list. Coordinates are CSS pixels.
 */
import { MAX_ROWS, MAX_COLS, colName, FormulaError, formatValue } from './engine.js';
const rgbaCache = new Map();
function rgba(color, alpha = 1) {
  const id = `${color}/${alpha}`; if (rgbaCache.has(id)) return rgbaCache.get(id);
  let h = (color || '#ffffff').replace('#', ''); if (h.length === 3) h = h.split('').map(x => x + x).join('');
  if (!/^[0-9a-f]{6}$/i.test(h)) h = 'ffffff';
  const v = [parseInt(h.slice(0, 2), 16) / 255, parseInt(h.slice(2, 4), 16) / 255, parseInt(h.slice(4, 6), 16) / 255, alpha]; rgbaCache.set(id, v); return v;
}
export class AxisLayout {
  constructor(count, defaultSize, overrides = new Map(), hidden = new Set()) {
    this.count = count; this.defaultSize = defaultSize;
    const map = new Map(overrides); for (const i of hidden) map.set(i, 0);
    this.entries = [...map].filter(([i]) => i >= 0 && i < count).sort((a, b) => a[0] - b[0]);
    this.indexes = this.entries.map(e => e[0]); this.sizes = new Map(this.entries); this.prefix = [0];
    for (const [, size] of this.entries) this.prefix.push(this.prefix.at(-1) + size - defaultSize);
  }
  lowerBound(i) { let lo = 0, hi = this.indexes.length; while (lo < hi) { const m = (lo + hi) >> 1; if (this.indexes[m] < i) lo = m + 1; else hi = m; } return lo; }
  offset(i) { return i * this.defaultSize + this.prefix[this.lowerBound(i)]; }
  size(i) { return this.sizes.get(i) ?? this.defaultSize; }
  find(pixel) { let lo = 0, hi = this.count - 1; while (lo < hi) { const m = Math.ceil((lo + hi) / 2); if (this.offset(m) <= pixel) lo = m; else hi = m - 1; } while (lo < this.count - 1 && this.size(lo) === 0) lo++; return lo; }
}
class GlyphAtlas {
  constructor(scale = 2) {
    this.size = 2048; this.scale = Math.min(3, Math.max(1, scale)); this.canvas = document.createElement('canvas'); this.canvas.width = this.canvas.height = this.size;
    this.ctx = this.canvas.getContext('2d', { willReadFrequently: true }); this.ctx.textBaseline = 'alphabetic'; this.map = new Map(); this.x = 2; this.y = 2; this.rowHeight = 0; this.dirty = false; this.overflow = false;
  }
  reset(scale) { this.scale = Math.min(3, Math.max(1, scale)); this.map.clear(); this.ctx.clearRect(0, 0, this.size, this.size); this.x = this.y = 2; this.rowHeight = 0; this.dirty = true; this.overflow = false; }
  glyph(char, font) {
    const id = font + '\0' + char; if (this.map.has(id)) return this.map.get(id);
    const ctx = this.ctx; ctx.font = font; const metrics = ctx.measureText(char), advance = metrics.width;
    const left = Math.ceil(Math.max(0, metrics.actualBoundingBoxLeft || 0)) + 2;
    const right = Math.ceil(Math.max(advance, metrics.actualBoundingBoxRight || advance)) + 2;
    const ascent = Math.ceil(metrics.actualBoundingBoxAscent || parseFloat(font.match(/(\d+(?:\.\d+)?)px/)?.[1]) * 0.8 || 12) + 2;
    const descent = Math.ceil(metrics.actualBoundingBoxDescent || 3) + 2;
    const w = Math.ceil((left + right) * this.scale), h = Math.ceil((ascent + descent) * this.scale);
    if (this.x + w + 2 > this.size) { this.x = 2; this.y += this.rowHeight + 2; this.rowHeight = 0; }
    if (this.y + h + 2 > this.size) { this.overflow = true; return null; }
    ctx.save(); ctx.translate(this.x, this.y); ctx.scale(this.scale, this.scale); ctx.font = font; ctx.fillStyle = '#ffffff'; ctx.textBaseline = 'alphabetic'; ctx.fillText(char, left, ascent); ctx.restore();
    const glyph = { u: this.x / this.size, v: this.y / this.size, uw: w / this.size, vh: h / this.size, w: w / this.scale, h: h / this.scale, left, ascent, advance };
    this.map.set(id, glyph); this.x += w + 2; this.rowHeight = Math.max(this.rowHeight, h); this.dirty = true; return glyph;
  }
}
const SHADER = `
struct Globals { viewport: vec2f, padding: vec2f };
@group(0) @binding(0) var<uniform> globals: Globals;
@group(0) @binding(1) var atlas: texture_2d<f32>;
@group(0) @binding(2) var atlasSampler: sampler;
struct In {
  @location(0) rect: vec4f,
  @location(1) uv: vec4f,
  @location(2) color: vec4f,
  @location(3) clip: vec4f,
};
struct Out {
  @builtin(position) position: vec4f,
  @location(0) uv: vec2f,
  @location(1) color: vec4f,
  @location(2) point: vec2f,
  @location(3) @interpolate(flat) clip: vec4f,
  @location(4) @interpolate(flat) textured: f32,
};
@vertex fn vs(input: In, @builtin(vertex_index) index: u32) -> Out {
  let vertices = array<vec2f, 6>(vec2f(0,0),vec2f(1,0),vec2f(0,1),vec2f(0,1),vec2f(1,0),vec2f(1,1));
  let v = vertices[index]; let p = input.rect.xy + v * input.rect.zw;
  var o: Out; o.position = vec4f(p.x / globals.viewport.x * 2.0 - 1.0, 1.0 - p.y / globals.viewport.y * 2.0, 0, 1);
  o.uv = input.uv.xy + v * input.uv.zw; o.color = input.color; o.point = p; o.clip = input.clip;
  o.textured = select(0.0, 1.0, input.uv.z > 0); return o;
}
@fragment fn fs(input: Out) -> @location(0) vec4f {
  if (input.point.x < input.clip.x || input.point.y < input.clip.y || input.point.x >= input.clip.z || input.point.y >= input.clip.w) { discard; }
  let mask = textureSampleLevel(atlas, atlasSampler, input.uv, 0.0).a;
  let alpha = input.color.a * select(1.0, mask, input.textured > 0);
  return vec4f(input.color.rgb * alpha, alpha);
}`;
export class GridRenderer {
  constructor(host, workbook, onBackend) {
    this.host = host; this.workbook = workbook; this.onBackend = onBackend; this.canvas = document.createElement('canvas'); this.canvas.className = 'grid-canvas'; this.canvas.setAttribute('aria-hidden', 'true'); host.prepend(this.canvas);
    this.backend = 'initializing'; this.width = 1; this.height = 1; this.zoom = 1; this.scrollX = 0; this.scrollY = 0;
    this.headerW = 46; this.headerH = 27; this.selection = { r1: 12, c1: 6, r2: 12, c2: 6 }; this.active = { r: 12, c: 6 }; this.showFormulas = false; this.dark = false;
    this.atlas = new GlyphAtlas(Math.max(2, window.devicePixelRatio || 1)); this.measureCanvas = document.createElement('canvas'); this.measure = this.measureCanvas.getContext('2d'); this.textWidths = new Map();
    this.commands = []; this.instances = new Float32Array(65536); this.instanceCount = 0; this.frameRequested = false; this.lastFrameMs = 0; this.visibleCellCount = 0; this.onFrame = null; this.disposed = false;
    this.resizeObserver = new ResizeObserver(() => this.resize()); this.resizeObserver.observe(host); this.syncLayout();
  }
  async initialize(forceCanvas = false) {
    if (!forceCanvas && navigator.gpu) {
      try {
        const adapter = await navigator.gpu.requestAdapter({ powerPreference: 'high-performance' }); if (!adapter) throw new Error('No WebGPU adapter.');
        this.device = await adapter.requestDevice(); this.device.addEventListener('uncapturederror', e => { console.error('WebGPU:', e.error.message); this.fallback('WebGPU validation failure'); });
        const module = this.device.createShaderModule({ label: 'Gridline instanced sheet shader', code: SHADER });
        const info = await module.getCompilationInfo(); if (info.messages.some(m => m.type === 'error')) throw new Error(info.messages.filter(m => m.type === 'error').map(m => m.message).join('\n'));
        const format = navigator.gpu.getPreferredCanvasFormat(); this.context = this.canvas.getContext('webgpu'); if (!this.context) throw new Error('WebGPU canvas context unavailable.');
        this.context.configure({ device: this.device, format, alphaMode: 'opaque' });
        this.pipeline = await this.device.createRenderPipelineAsync({ label: 'Gridline sheet pipeline', layout: 'auto',
          vertex: { module, entryPoint: 'vs', buffers: [{ arrayStride: 64, stepMode: 'instance', attributes: [0, 1, 2, 3].map(i => ({ shaderLocation: i, offset: i * 16, format: 'float32x4' })) }] },
          fragment: { module, entryPoint: 'fs', targets: [{ format, blend: { color: { srcFactor: 'one', dstFactor: 'one-minus-src-alpha', operation: 'add' }, alpha: { srcFactor: 'one', dstFactor: 'one-minus-src-alpha', operation: 'add' } } }] },
          primitive: { topology: 'triangle-list' }
        });
        this.uniform = this.device.createBuffer({ size: 16, usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST });
        this.texture = this.device.createTexture({ label: 'Glyph mask atlas', size: [this.atlas.size, this.atlas.size], format: 'rgba8unorm', usage: GPUTextureUsage.TEXTURE_BINDING | GPUTextureUsage.COPY_DST | GPUTextureUsage.RENDER_ATTACHMENT });
        this.sampler = this.device.createSampler({ magFilter: 'linear', minFilter: 'linear' });
        this.bindGroup = this.device.createBindGroup({ layout: this.pipeline.getBindGroupLayout(0), entries: [{ binding: 0, resource: { buffer: this.uniform } }, { binding: 1, resource: this.texture.createView() }, { binding: 2, resource: this.sampler }] });
        this.backend = 'webgpu'; this.adapterInfo = adapter.info?.description || adapter.info?.architecture || 'WebGPU';
        this.device.lost.then(info => { if (!this.disposed && this.backend === 'webgpu') this.fallback(info.message || 'GPU device lost'); });
      } catch (error) { console.warn('Gridline WebGPU fallback:', error.message); this.fallbackReason = error.message; this.createFallback(); }
    } else { this.fallbackReason = forceCanvas ? 'Canvas renderer requested.' : 'WebGPU is not exposed by this browser.'; this.createFallback(); }
    this.resize(); this.onBackend?.(this.backend, this.fallbackReason); this.requestFrame(); return this.backend;
  }
  createFallback() {
    if (this.context) { const replacement = this.canvas.cloneNode(false); this.canvas.replaceWith(replacement); this.canvas = replacement; this.context = null; }
    this.ctx = this.canvas.getContext('2d', { alpha: false }); this.backend = 'canvas2d';
  }
  fallback(reason) { if (this.backend === 'canvas2d') return; this.fallbackReason = reason; this.createFallback(); this.onBackend?.(this.backend, reason); this.resize(); this.requestFrame(); }
  resize() {
    const box = this.host.getBoundingClientRect(); this.width = Math.max(1, Math.floor(box.width)); this.height = Math.max(1, Math.floor(box.height));
    this.dpr = Math.min(window.devicePixelRatio || 1, 3); this.canvas.width = Math.round(this.width * this.dpr); this.canvas.height = Math.round(this.height * this.dpr);
    this.canvas.style.width = this.width + 'px'; this.canvas.style.height = this.height + 'px'; this.requestFrame();
  }
  syncLayout() {
    const sheet = this.workbook.activeSheet; this.populatedRows = sheet.populatedRange({r1:0,c1:0,r2:MAX_ROWS-1,c2:MAX_COLS-1}).r2 + 1; this.cols = new AxisLayout(MAX_COLS, 106, sheet.colWidths, sheet.hiddenCols); this.rows = new AxisLayout(MAX_ROWS, 27, sheet.rowHeights, sheet.hiddenRows); this.requestFrame();
  }
  setZoom(zoom) {
    const old = this.zoom; this.zoom = Math.max(0.5, Math.min(2, zoom)); this.scrollX *= this.zoom / old; this.scrollY *= this.zoom / old;
    this.atlas.reset(Math.max(2, this.dpr || 1)); this.textWidths.clear(); this.requestFrame();
  }
  frozenSize() { const sheet = this.workbook.activeSheet; return { x: this.cols.offset(sheet.freezeCols) * this.zoom, y: this.rows.offset(sheet.freezeRows) * this.zoom }; }
  cellRect(r, c, merge = true) {
    const sheet = this.workbook.activeSheet, q = merge ? sheet.mergeAt(r, c) : null;
    if (q) { r = q.r1; c = q.c1; }
    const z = this.zoom;
    return { x: this.headerW + this.cols.offset(c) * z - (c < sheet.freezeCols ? 0 : this.scrollX), y: this.headerH + this.rows.offset(r) * z - (r < sheet.freezeRows ? 0 : this.scrollY), w: (this.cols.offset(q ? q.c2 + 1 : c + 1) - this.cols.offset(c)) * z, h: (this.rows.offset(q ? q.r2 + 1 : r + 1) - this.rows.offset(r)) * z };
  }
  hitTest(x, y) {
    const sheet = this.workbook.activeSheet, frozen = this.frozenSize();
    const px = Math.max(0, x - this.headerW), py = Math.max(0, y - this.headerH);
    let c = this.cols.find((px + (px >= frozen.x ? this.scrollX : 0)) / this.zoom), r = this.rows.find((py + (py >= frozen.y ? this.scrollY : 0)) / this.zoom);
    if (y < this.headerH && x >= this.headerW && c > 0 && Math.abs(x - this.cellRect(0,c,false).x) < 6) { do { c--; } while (c > 0 && this.cols.size(c) === 0); }
    const merge = y >= this.headerH && x >= this.headerW ? sheet.mergeAt(r, c) : null; if (merge) { r = merge.r1; c = merge.c1; }
    return { r, c, rowHeader: x < this.headerW, colHeader: y < this.headerH };
  }
  ensureVisible(r, c) {
    const sheet = this.workbook.activeSheet, frozen = this.frozenSize(), rect = this.cellRect(r, c);
    if (c >= sheet.freezeCols) { if (rect.x < this.headerW + frozen.x) this.scrollX -= this.headerW + frozen.x - rect.x; else if (rect.x + rect.w > this.width - 14) this.scrollX += rect.x + rect.w - this.width + 14; }
    if (r >= sheet.freezeRows) { if (rect.y < this.headerH + frozen.y) this.scrollY -= this.headerH + frozen.y - rect.y; else if (rect.y + rect.h > this.height - 14) this.scrollY += rect.y + rect.h - this.height + 14; }
    this.clampScroll(); this.requestFrame();
  }
  clampScroll() { this.scrollX = Math.max(0, Math.min(this.scrollX, this.cols.offset(MAX_COLS) * this.zoom - this.width + this.headerW)); this.scrollY = Math.max(0, Math.min(this.scrollY, this.rows.offset(MAX_ROWS) * this.zoom - this.height + this.headerH)); }
  requestFrame() { if (!this.frameRequested && !this.disposed) { this.frameRequested = true; requestAnimationFrame(() => { this.frameRequested = false; this.draw(); }); } }
  rect(x, y, w, h, color, clip = [0, 0, this.width, this.height], alpha = 1) {
    if (w <= 0 || h <= 0 || x > clip[2] || y > clip[3] || x + w < clip[0] || y + h < clip[1]) return;
    this.commands.push({ kind: 'rect', x, y, w, h, color, clip, alpha });
  }
  outline(x, y, w, h, color, thickness = 1, clip) { this.rect(x, y, w, thickness, color, clip); this.rect(x, y + h - thickness, w, thickness, color, clip); this.rect(x, y, thickness, h, color, clip); this.rect(x + w - thickness, y, thickness, h, color, clip); }
  font(style = {}) { return `${style.italic ? 'italic ' : ''}${style.bold ? '600' : '400'} ${(style.fontSize ?? 13) * this.zoom}px ${style.fontFamily || 'Aptos, "Segoe UI", Arial, sans-serif'}`; }
  measureText(text, font) { const key = font + '|' + text; if (this.textWidths.has(key)) return this.textWidths.get(key); this.measure.font = font; const width = this.measure.measureText(text).width; if (this.textWidths.size > 20000) this.textWidths.clear(); this.textWidths.set(key, width); return width; }
  text(text, x, baseline, color, style = {}, clip = [0, 0, this.width, this.height]) {
    text = String(text); if (!text) return; const font = this.font(style);
    this.commands.push({ kind: 'text', text, x, baseline, color, font, clip, underline: style.underline });
  }
  cellText(text, value, style, rect, clip) {
    const pad = 9 * this.zoom, font = this.font(style), size = (style.fontSize ?? 13) * this.zoom, avail = Math.max(0, rect.w - pad * 2);
    const align = style.align || (typeof value === 'number' ? 'right' : 'left');
    const baseClip = [Math.max(rect.x + 1, clip[0]), Math.max(rect.y + 1, clip[1]), Math.min(rect.x + rect.w - 2, clip[2]), Math.min(rect.y + rect.h - 1, clip[3])];
    const color = value instanceof FormulaError ? '#c7473d' : style.color || (this.dark ? '#dfebe5' : '#293b32');
    let lines = [text];
    if (style.wrap) {
      lines = []; const paragraphs = text.split('\n');
      for (const paragraph of paragraphs) {
        let line = ''; for (const word of paragraph.split(' ')) { const candidate = line ? `${line} ${word}` : word; if (line && this.measureText(candidate, font) > avail) { lines.push(line); line = word; } else line = candidate; } lines.push(line);
      }
      lines = lines.slice(0, Math.max(1, Math.floor(rect.h / (size * 1.3))));
    }
    const lineH = size * 1.3, totalH = lines.length * lineH;
    lines.forEach((line, i) => {
      let width = this.measureText(line, font);
      if (typeof value === 'number' && !style.wrap && width > avail) { line = '#'.repeat(Math.max(1, Math.floor(avail / this.measureText('#', font)))); width = this.measureText(line, font); }
      const x = align === 'right' ? rect.x + rect.w - pad - width : align === 'center' ? rect.x + (rect.w - width) / 2 : rect.x + pad;
      const baseline = rect.y + (rect.h - totalH) / 2 + lineH * i + size * 0.99;
      this.text(line, x, baseline, color, style, baseClip);
    });
  }
  visibleIndices(axis, scroll, size, frozenCount, header) {
    const frozen = axis.offset(frozenCount) * this.zoom, list = [];
    for (let i = 0; i < frozenCount && axis.offset(i) * this.zoom < size - header; i++) if (axis.size(i) > 0) list.push(i);
    const first = Math.max(frozenCount, axis.find((scroll + frozen) / this.zoom)), end = Math.min(axis.count - 1, axis.find((scroll + size - header) / this.zoom) + 1);
    for (let i = first; i <= end; i++) if (axis.size(i) > 0) list.push(i);
    return list;
  }
  draw() {
    if (this.backend === 'initializing' || this.disposed) return;
    const started = performance.now(), sheet = this.workbook.activeSheet;
    this.commands = []; this.instanceCount = 0; this.visibleCellCount = 0;
    const bg = this.dark ? '#19251f' : '#ffffff', grid = this.dark ? '#2a3b31' : '#e6ebe8', head = this.dark ? '#21342a' : '#f4f7f5';
    this.rect(0, 0, this.width, this.height, bg);
    const rows = this.visibleIndices(this.rows, this.scrollY, this.height, sheet.freezeRows, this.headerH), cols = this.visibleIndices(this.cols, this.scrollX, this.width, sheet.freezeCols, this.headerW), frozen = this.frozenSize();
    this.visibleRows = rows; this.visibleCols = cols;
    const ruleStats = new Map();
    for (const rule of sheet.conditionalRules) {
      if (rule.type === 'bars' || rule.type === 'scale') {
        let min = Infinity, max = -Infinity;
        for (let r = rule.range.r1; r <= rule.range.r2 && r < rule.range.r1 + 10000; r++) for (let c = rule.range.c1; c <= rule.range.c2; c++) { const v = this.workbook.value(sheet, r, c); if (typeof v === 'number') { min = Math.min(min, v); max = Math.max(max, v); } }
        ruleStats.set(rule, { min, max });
      }
    }
    const drawnMerges = new Set();
    for (const r of rows) for (const c of cols) {
      let cellR = r, cellC = c; const merge = sheet.mergeAt(r, c);
      if (merge) { const id = `${merge.r1},${merge.c1}`; if (drawnMerges.has(id)) continue; drawnMerges.add(id); cellR = merge.r1; cellC = merge.c1; }
      const rect = this.cellRect(cellR, cellC), { x, y, w, h } = rect;
      const clip = [this.headerW + (c < sheet.freezeCols ? 0 : frozen.x), this.headerH + (r < sheet.freezeRows ? 0 : frozen.y), c < sheet.freezeCols ? this.headerW + frozen.x : this.width, r < sheet.freezeRows ? this.headerH + frozen.y : this.height];
      if (x + w <= clip[0] || y + h <= clip[1]) continue;
      const cell = sheet.get(cellR, cellC), style = sheet.style(cellR, cellC), value = this.workbook.value(sheet, cellR, cellC); this.visibleCellCount++;
      for (const rule of sheet.conditionalRules) {
        const q = rule.range; if (r < q.r1 || r > q.r2 || c < q.c1 || c > q.c2) continue;
        if (rule.type === 'positive' && typeof value === 'number') { style.color = value >= 0 ? '#16815c' : '#c14944'; style.bold = true; }
        if (rule.type === 'greater' && typeof value === 'number' && value > rule.value) { style.fill = '#d9eee3'; style.color = '#166444'; }
        if (rule.type === 'scale' && typeof value === 'number') {
          const stats = ruleStats.get(rule), t = stats.max === stats.min ? 0.5 : (value - stats.min) / (stats.max - stats.min);
          style.fill = '#' + [Math.round(243 - t * 74), Math.round(248 - t * 35), Math.round(241 - t * 52)].map(x => x.toString(16).padStart(2, '0')).join('');
        }
      }
      if (style.fill) this.rect(x, y, w, h, style.fill, clip);
      if (sheet.gridlines) { this.rect(x + w - 1, y, 1, h, grid, clip); this.rect(x, y + h - 1, w, 1, grid, clip); }
      if (style.border) this.outline(x, y, w, h, style.borderColor || '#cbd8d0', 1, clip);
      if (style.bottomBorder) this.rect(x, y + h - 2, w, 2, style.bottomBorder, clip);
      for (const rule of sheet.conditionalRules) {
        if (rule.type !== 'bars' || typeof value !== 'number') continue; const q = rule.range;
        if (r < q.r1 || r > q.r2 || c < q.c1 || c > q.c2) continue;
        const stats = ruleStats.get(rule), max = Math.max(Math.abs(stats.min), Math.abs(stats.max), 1);
        this.rect(x + 4, y + 5, Math.max(0, (w - 8) * Math.abs(value) / max), h - 10, value < 0 ? '#f4ceca' : '#c5e9d7', clip, 0.8);
      }
      const text = this.showFormulas && cell?.raw.startsWith('=') ? cell.raw : this.workbook.display(sheet, cellR, cellC);
      if (text) this.cellText(text, value, style, rect, clip);
      if (cell?.note) this.rect(x + w - 7, y + 2, 5, 5, '#b691cc', clip);
      if (sheet.filters && r === sheet.filters.range.r1 && c >= sheet.filters.range.c1 && c <= sheet.filters.range.c2) this.text('⌄', x + w - 12, y + h / 2 + 4, '#789185', { fontSize: 11 }, clip);
    }
    // Header strips occlude all scrolled cells, independently of selection and frozen panes.
    this.rect(0, 0, this.width, this.headerH, head); this.rect(0, 0, this.headerW, this.height, head);
    const q = this.selection, selectedHead = this.dark ? '#295b42' : '#e0f0e6';
    for (const c of cols) {
      const rect = this.cellRect(0, c, false), clip = [this.headerW + (c < sheet.freezeCols ? 0 : frozen.x), 0, c < sheet.freezeCols ? this.headerW + frozen.x : this.width, this.headerH];
      if (c >= q.c1 && c <= q.c2) { this.rect(rect.x, 0, rect.w, this.headerH, selectedHead, clip); this.rect(rect.x, this.headerH - 2, rect.w, 2, '#18835a', clip); }
      this.rect(rect.x + rect.w - 1, 0, 1, this.headerH, grid, clip);
      const label = colName(c), font = { fontSize: 11, bold: c >= q.c1 && c <= q.c2 };
      this.text(label, rect.x + (rect.w - this.measureText(label, this.font(font))) / 2, 18, this.dark ? '#a8bdb0' : '#64776c', font, clip);
    }
    for (const r of rows) {
      const rect = this.cellRect(r, 0, false), clip = [0, this.headerH + (r < sheet.freezeRows ? 0 : frozen.y), this.headerW, r < sheet.freezeRows ? this.headerH + frozen.y : this.height];
      if (r >= q.r1 && r <= q.r2) { this.rect(0, rect.y, this.headerW, rect.h, selectedHead, clip); this.rect(this.headerW - 2, rect.y, 2, rect.h, '#18835a', clip); }
      this.rect(0, rect.y + rect.h - 1, this.headerW, 1, grid, clip);
      const label = String(r + 1), font = { fontSize: 11, bold: r >= q.r1 && r <= q.r2 };
      this.text(label, (this.headerW - this.measureText(label, this.font(font))) / 2, rect.y + rect.h / 2 + 4, this.dark ? '#a8bdb0' : '#64776c', font, clip);
    }
    this.rect(0, 0, this.headerW, this.headerH, head); this.text('◢', 28, 19, '#c3d1c8', { fontSize: 15 });
    if (frozen.y) this.rect(this.headerW, this.headerH + frozen.y - 1, this.width, 2, '#b1c8ba');
    if (frozen.x) this.rect(this.headerW + frozen.x - 1, this.headerH, 2, this.height, '#b1c8ba');
    this.drawSelection(q, '#18835a', true, this.copyRange && ['r1','r2','c1','c2'].every(key=>q[key]===this.copyRange[key]));
    if (this.copyRange) this.drawSelection(this.copyRange, '#176b4a', false, true);
    if (this.fillPreview) this.drawSelection(this.fillPreview, '#18835a', false);
    if (this.backend === 'webgpu') this.flushGPU(); else this.flushCanvas();
    this.lastFrameMs = performance.now() - started; this.onFrame?.();
    if (this.atlas.overflow && this.backend === 'webgpu') this.fallback('Visible glyph set exceeds the bounded atlas capacity.');
  }
  drawSelection(q, color, handle, dashed = false) {
    const sheet = this.workbook.activeSheet, frozen = this.frozenSize();
    const regions = [
      [0, sheet.freezeRows - 1, 0, sheet.freezeCols - 1, [this.headerW, this.headerH, this.headerW + frozen.x, this.headerH + frozen.y]],
      [0, sheet.freezeRows - 1, sheet.freezeCols, MAX_COLS - 1, [this.headerW + frozen.x, this.headerH, this.width, this.headerH + frozen.y]],
      [sheet.freezeRows, MAX_ROWS - 1, 0, sheet.freezeCols - 1, [this.headerW, this.headerH + frozen.y, this.headerW + frozen.x, this.height]],
      [sheet.freezeRows, MAX_ROWS - 1, sheet.freezeCols, MAX_COLS - 1, [this.headerW + frozen.x, this.headerH + frozen.y, this.width, this.height]]
    ];
    for (const [r1, r2, c1, c2, clip] of regions) {
      const a = { r: Math.max(q.r1, r1), c: Math.max(q.c1, c1) }, b = { r: Math.min(q.r2, r2), c: Math.min(q.c2, c2) }; if (a.r > b.r || a.c > b.c) continue;
      const p = this.cellRect(a.r, a.c, false), end = this.cellRect(b.r, b.c, false), w = end.x + end.w - p.x, h = end.y + end.h - p.y;
      if (handle && (q.r1 !== q.r2 || q.c1 !== q.c2)) this.rect(p.x, p.y, w, h, color, clip, 0.07);
      if (dashed) {
        for (let x=Math.max(p.x,clip[0]);x<Math.min(p.x+w,clip[2]);x+=8) { this.rect(x,p.y,Math.min(4,p.x+w-x),2,color,clip); this.rect(x,p.y+h-2,Math.min(4,p.x+w-x),2,color,clip); }
        for (let y=Math.max(p.y,clip[1]);y<Math.min(p.y+h,clip[3]);y+=8) { this.rect(p.x,y,2,Math.min(4,p.y+h-y),color,clip); this.rect(p.x+w-2,y,2,Math.min(4,p.y+h-y),color,clip); }
      } else this.outline(p.x, p.y, w, h, color, handle ? 2 : 1, clip);
      if (handle && b.r === q.r2 && b.c === q.c2) { this.rect(p.x + w - 4, p.y + h - 4, 8, 8, '#ffffff', clip); this.rect(p.x + w - 3, p.y + h - 3, 6, 6, color, clip); }
    }
  }
  addInstance(x, y, w, h, uv, color, clip) {
    if ((this.instanceCount + 1) * 16 > this.instances.length) { const next = new Float32Array(this.instances.length * 2); next.set(this.instances); this.instances = next; }
    const i = this.instanceCount++ * 16; this.instances.set([x, y, w, h, ...uv, ...color, ...clip], i);
  }
  flushGPU() {
    for (const cmd of this.commands) {
      if (cmd.kind === 'rect') this.addInstance(cmd.x, cmd.y, cmd.w, cmd.h, [0, 0, 0, 0], rgba(cmd.color, cmd.alpha), cmd.clip);
      else {
        let x = cmd.x;
        for (const char of cmd.text) {
          const glyph = this.atlas.glyph(char, cmd.font); if (!glyph) continue;
          if (x + glyph.advance >= cmd.clip[0] && x <= cmd.clip[2]) this.addInstance(x - glyph.left, cmd.baseline - glyph.ascent, glyph.w, glyph.h, [glyph.u, glyph.v, glyph.uw, glyph.vh], rgba(cmd.color), cmd.clip);
          x += glyph.advance; if (x > cmd.clip[2] + 30) break;
        }
        if (cmd.underline) this.addInstance(cmd.x, cmd.baseline + 2, x - cmd.x, 1, [0, 0, 0, 0], rgba(cmd.color), cmd.clip);
      }
    }
    if (this.atlas.dirty) { this.device.queue.copyExternalImageToTexture({ source: this.atlas.canvas }, { texture: this.texture, premultipliedAlpha: false }, [this.atlas.size, this.atlas.size]); this.atlas.dirty = false; }
    const bytes = this.instanceCount * 64;
    if (!this.vertexBuffer || this.bufferSize < bytes) { this.vertexBuffer?.destroy(); this.bufferSize = Math.max(65536, 2 ** Math.ceil(Math.log2(bytes || 1))); this.vertexBuffer = this.device.createBuffer({ label: 'Visible cell instances', size: this.bufferSize, usage: GPUBufferUsage.VERTEX | GPUBufferUsage.COPY_DST }); }
    if (bytes) this.device.queue.writeBuffer(this.vertexBuffer, 0, this.instances.buffer, 0, bytes);
    this.device.queue.writeBuffer(this.uniform, 0, new Float32Array([this.width, this.height, 0, 0]));
    const encoder = this.device.createCommandEncoder(); const pass = encoder.beginRenderPass({ colorAttachments: [{ view: this.context.getCurrentTexture().createView(), clearValue: { r: 1, g: 1, b: 1, a: 1 }, loadOp: 'clear', storeOp: 'store' }] });
    pass.setPipeline(this.pipeline); pass.setBindGroup(0, this.bindGroup); pass.setVertexBuffer(0, this.vertexBuffer); pass.draw(6, this.instanceCount); pass.end(); this.device.queue.submit([encoder.finish()]);
  }
  flushCanvas() {
    const ctx = this.ctx; ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    for (const cmd of this.commands) {
      ctx.save(); ctx.beginPath(); ctx.rect(cmd.clip[0], cmd.clip[1], cmd.clip[2] - cmd.clip[0], cmd.clip[3] - cmd.clip[1]); ctx.clip(); ctx.fillStyle = cmd.color;
      if (cmd.kind === 'rect') { ctx.globalAlpha = cmd.alpha; ctx.fillRect(cmd.x, cmd.y, cmd.w, cmd.h); }
      else { ctx.font = cmd.font; ctx.textBaseline = 'alphabetic'; ctx.fillText(cmd.text, cmd.x, cmd.baseline); if (cmd.underline) ctx.fillRect(cmd.x, cmd.baseline + 2, ctx.measureText(cmd.text).width, 1); }
      ctx.restore();
    }
  }
  dispose() { this.disposed = true; this.resizeObserver.disconnect(); this.vertexBuffer?.destroy(); this.uniform?.destroy(); this.texture?.destroy(); this.device?.destroy(); }
}
