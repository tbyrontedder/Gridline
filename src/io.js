/** Local-only CSV / Gridline JSON / basic OOXML interoperability. No third-party libraries. */
import { Workbook, Sheet, keyOf, address, parseAddress, parseRange, shiftFormula, FormulaError, rawValue, numberFormatCode, numberFormatStyle } from './engine.js';
export function parseDelimited(text, delimiter = null) {
  text = text.replace(/^\uFEFF/, '');
  if (!delimiter) { const first = text.split(/\r?\n/, 1)[0]; delimiter = first.includes('\t') ? '\t' : first.split(';').length > first.split(',').length ? ';' : ','; }
  const rows = []; let row = [], field = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) { if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false; } else field += ch; }
    else if (ch === '"' && field === '') quoted = true;
    else if (ch === delimiter) { row.push(field); field = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(field); rows.push(row); row = []; field = ''; }
    else field += ch;
  }
  if (quoted) throw new Error('The file contains an unterminated quoted field.');
  if (field !== '' || row.length || !rows.length) { row.push(field); rows.push(row); }
  if (rows.reduce((n, r) => n + r.length, 0) > 200000) throw new Error('Import is limited to 200,000 cells per delimited file.');
  return rows;
}
export function serializeDelimited(rows, delimiter = ',') {
  return rows.map(row => row.map(v => { const text = String(v ?? ''); return text.includes(delimiter) || /["\n\r]/.test(text) ? '"' + text.replaceAll('"', '""') + '"' : text; }).join(delimiter)).join('\r\n');
}
export function exportCSV(workbook, sheet = workbook.activeSheet, formulas = false) {
  const q = sheet.usedRange(); if ((q.r2 + 1) * (q.c2 + 1) > 200000) throw new Error('CSV export is limited to a 200,000-cell rectangular used range.');
  const rows = []; for (let r = 0; r <= q.r2; r++) {
    const row = []; for (let c = 0; c <= q.c2; c++) {
      let value = formulas ? sheet.raw(r, c) : workbook.value(sheet, r, c);
      if (value instanceof FormulaError) value = value.code;
      // Neutralize spreadsheet-formula injection when exporting literal strings.
      if (typeof value === 'string' && /^[=+\-@\t\r]/.test(value) && !formulas) value = "'" + value;
      row.push(value ?? '');
    } rows.push(row);
  }
  return '\uFEFF' + serializeDelimited(rows);
}
export function workbookFromCSV(text, title = 'Imported workbook') {
  const rows = parseDelimited(text), wb = new Workbook(); wb.title = title; wb.activeSheet.name = 'Imported data';
  rows.forEach((row, r) => row.forEach((raw, c) => { if (raw !== '') wb.activeSheet.cells.set(keyOf(r, c), { raw: raw.startsWith('=') ? "'" + raw : raw, ...(r === 0 ? { style: { bold: true, fill: '#e4efe8', color: '#18694a' } } : {}) }); }));
  wb.activeSheet.freezeRows = 1; return wb;
}
export function downloadFile(name, data, type = 'application/octet-stream') {
  const blob = data instanceof Blob ? data : new Blob([data], { type }); const url = URL.createObjectURL(blob), a = document.createElement('a'); a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url), 10000);
}
const encoder = new TextEncoder(), decoder = new TextDecoder();
const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; crcTable[n] = c >>> 0; }
function crc32(bytes) { let crc = 0xffffffff; for (const b of bytes) crc = crcTable[(crc ^ b) & 255] ^ (crc >>> 8); return (crc ^ 0xffffffff) >>> 0; }
export function zipStore(files, compressed = {}) {
  const local = [], central = []; let offset = 0;
  for (const [name, source] of Object.entries(files)) {
    const path = encoder.encode(name), data = typeof source === 'string' ? encoder.encode(source) : source, crc = crc32(data), packed = compressed[name] || data, method = compressed[name] ? 8 : 0;
    const head = new Uint8Array(30 + path.length), v = new DataView(head.buffer);
    v.setUint32(0, 0x04034b50, true); v.setUint16(4, 20, true); v.setUint16(6, 0x800, true); v.setUint16(8, method, true); v.setUint16(12, 0x21, true); v.setUint32(14, crc, true); v.setUint32(18, packed.length, true); v.setUint32(22, data.length, true); v.setUint16(26, path.length, true); head.set(path, 30);
    const cd = new Uint8Array(46 + path.length), d = new DataView(cd.buffer);
    d.setUint32(0, 0x02014b50, true); d.setUint16(4, 20, true); d.setUint16(6, 20, true); d.setUint16(8, 0x800, true); d.setUint16(10, method, true); d.setUint16(14, 0x21, true); d.setUint32(16, crc, true); d.setUint32(20, packed.length, true); d.setUint32(24, data.length, true); d.setUint16(28, path.length, true); d.setUint32(42, offset, true); cd.set(path, 46);
    local.push(head, packed); central.push(cd); offset += head.length + packed.length;
  }
  const centralLength = central.reduce((n, b) => n + b.length, 0), end = new Uint8Array(22), ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, central.length, true); ev.setUint16(10, central.length, true); ev.setUint32(12, centralLength, true); ev.setUint32(16, offset, true);
  const out = new Uint8Array(offset + centralLength + 22); let at = 0; for (const b of [...local, ...central, end]) { out.set(b, at); at += b.length; } return out;
}
export async function unzip(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer), v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length > 32 * 1024 * 1024) throw new Error('XLSX import is limited to 32 MB compressed.');
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) if (v.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('Not a valid ZIP-based XLSX file.');
  if (v.getUint16(eocd + 4, true) || v.getUint16(eocd + 6, true)) throw new Error('Multi-volume ZIP archives are not supported.');
  const count = v.getUint16(eocd + 10, true); if (count > 4096) throw new Error('Too many XLSX parts.');
  let pos = v.getUint32(eocd + 16, true), total = 0; const result = new Map();
  for (let i = 0; i < count; i++) {
    if (pos + 46 > bytes.length || v.getUint32(pos, true) !== 0x02014b50) throw new Error('Invalid ZIP directory.');
    const flags = v.getUint16(pos + 8, true), method = v.getUint16(pos + 10, true), crc = v.getUint32(pos + 16, true), compressed = v.getUint32(pos + 20, true), size = v.getUint32(pos + 24, true), nlen = v.getUint16(pos + 28, true), extra = v.getUint16(pos + 30, true), comment = v.getUint16(pos + 32, true), local = v.getUint32(pos + 42, true);
    total += size; if (total > 64 * 1024 * 1024 || size > 32 * 1024 * 1024) throw new Error('Expanded workbook exceeds the 64 MB limit.');
    const name = decoder.decode(bytes.subarray(pos + 46, pos + 46 + nlen)); if (flags & 1) throw new Error('Encrypted workbooks are not supported.');
    if (local + 30 > bytes.length || v.getUint32(local, true) !== 0x04034b50) throw new Error('Invalid ZIP entry.');
    const dataStart = local + 30 + v.getUint16(local + 26, true) + v.getUint16(local + 28, true);
    if (dataStart + compressed > bytes.length) throw new Error('Truncated ZIP entry.');
    let data = bytes.slice(dataStart, dataStart + compressed);
    if (method === 8) {
      if (typeof DecompressionStream === 'undefined') throw new Error('This browser does not provide the native decompressor required to open XLSX files.');
      const reader = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader(); const chunks = []; let read = 0;
      while (true) { const { done, value } = await reader.read(); if (done) break; read += value.length; if (read > size || read > 32 * 1024 * 1024) { await reader.cancel(); throw new Error('ZIP decompression limit exceeded.'); } chunks.push(value); }
      data = new Uint8Array(read); let offset = 0; for (const chunk of chunks) { data.set(chunk, offset); offset += chunk.length; }
    } else if (method !== 0) throw new Error(`Unsupported ZIP compression method: ${method}.`);
    if (data.length !== size || crc32(data) !== crc) throw new Error('Workbook ZIP integrity check failed.');
    if (!name.split('/').includes('..')) result.set(name.replace(/^\//, ''), data);
    pos += 46 + nlen + extra + comment;
  }
  return result;
}
const escapeXML = s => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;');
const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const spreadsheetNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relationshipNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelNS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const colorARGB = color => 'FF' + (color || '#000000').replace('#', '').toUpperCase();
function styleTable(workbook) {
  const styles = [{}], ids = new Map([['{}', 0]]);
  const idFor = style => { const normalized = Object.fromEntries(Object.entries(style ?? {}).filter(([, v]) => v !== undefined).sort(([a], [b]) => a.localeCompare(b))); const key = JSON.stringify(normalized); if (!ids.has(key)) { ids.set(key, styles.length); styles.push(normalized); } return ids.get(key); };
  for (const sheet of workbook.sheets) { for (const style of [...sheet.colStyles.values(), ...sheet.rowStyles.values()]) idFor(style); for (const [key] of sheet.cells) { const [r,c] = key.split(',').map(Number); idFor(sheet.style(r,c)); } }
  const formats = new Map(), fonts = [], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], xfs = [];
  styles.forEach((s, i) => {
    const size = (s.fontSize ?? 13) * 0.75; fonts.push(`<font>${s.bold ? '<b/>' : ''}${s.italic ? '<i/>' : ''}${s.underline ? '<u/>' : ''}<sz val="${size}"/><color rgb="${colorARGB(s.color || '#293b32')}"/><name val="${escapeXML(s.fontFamily?.split(',')[0] || 'Aptos')}"/></font>`);
    let fillId = 0; if (s.fill) { fillId = fills.length; fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${colorARGB(s.fill)}"/><bgColor indexed="64"/></patternFill></fill>`); }
    const fmt = numberFormatCode(s);
    let numFmtId = 0; if (fmt !== 'General') { if (!formats.has(fmt)) formats.set(fmt, 164 + formats.size); numFmtId = formats.get(fmt); }
    xfs.push(`<xf numFmtId="${numFmtId}" fontId="${i}" fillId="${fillId}" borderId="${s.border ? 1 : 0}" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="${s.format || s.decimals != null ? 1 : 0}" applyAlignment="1"><alignment vertical="center"${s.align ? ` horizontal="${s.align}"` : ''}${s.wrap ? ' wrapText="1"' : ''}/></xf>`);
  });
  const xml = xmlHeader + `<styleSheet xmlns="${spreadsheetNS}"><numFmts count="${formats.size}">${[...formats].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${escapeXML(code)}"/>`).join('')}</numFmts><fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border>${['left','right','top','bottom'].map(x => `<${x} style="thin"><color rgb="FFD4DFD8"/></${x}>`).join('')}<diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  return { xml, idFor };
}
export async function exportXLSX(workbook) {
  const files = {}, styles = styleTable(workbook);
  files['[Content_Types].xml'] = xmlHeader + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${workbook.sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`;
  files['_rels/.rels'] = xmlHeader + `<Relationships xmlns="${packageRelNS}"><Relationship Id="rId1" Type="${relationshipNS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  files['xl/workbook.xml'] = xmlHeader + `<workbook xmlns="${spreadsheetNS}" xmlns:r="${relationshipNS}"><bookViews><workbookView activeTab="${Math.max(0, workbook.sheets.findIndex(s => s.id === workbook.activeSheetId))}"/></bookViews><sheets>${workbook.sheets.map((s, i) => `<sheet name="${escapeXML(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${Object.keys(workbook.names).length ? '<definedNames>' + Object.entries(workbook.names).map(([n, v]) => `<definedName name="${escapeXML(n)}">${escapeXML(v)}</definedName>`).join('') + '</definedNames>' : ''}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
  files['xl/_rels/workbook.xml.rels'] = xmlHeader + `<Relationships xmlns="${packageRelNS}">${workbook.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${relationshipNS}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${workbook.sheets.length + 1}" Type="${relationshipNS}/styles" Target="styles.xml"/></Relationships>`;
  files['xl/styles.xml'] = styles.xml;
  workbook.sheets.forEach((sheet, i) => {
    const rows = new Map();
    for (const [key, cell] of sheet.cells) {
      const [r, c] = key.split(',').map(Number); if (!rows.has(r)) rows.set(r, []);
      const ref = address(r, c), sid = styles.idFor(sheet.style(r, c)), val = workbook.value(sheet, r, c), formula = cell.raw.startsWith('=') ? `<f>${escapeXML(cell.raw.slice(1))}</f>` : '';
      let type = '', content = '';
      if (val instanceof FormulaError) { type = ' t="e"'; content = `<v>${escapeXML(val.code)}</v>`; }
      else if (typeof val === 'number') content = `<v>${val}</v>`;
      else if (typeof val === 'boolean') { type = ' t="b"'; content = `<v>${+val}</v>`; }
      else if (formula) { type = ' t="str"'; content = `<v>${escapeXML(val ?? '')}</v>`; }
      else if (val !== null && val !== undefined) { type = ' t="inlineStr"'; content = `<is><t xml:space="preserve">${escapeXML(val)}</t></is>`; }
      rows.get(r).push({ c, xml: `<c r="${ref}" s="${sid}"${type}>${formula}${content}</c>` });
    }
    for (const r of new Set([...sheet.rowHeights.keys(), ...sheet.rowStyles.keys()])) if (!rows.has(r)) rows.set(r, []);
    const rowXML = [...rows].sort(([a], [b]) => a - b).map(([r, cells]) => `<row r="${r + 1}"${sheet.rowStyles.has(r) ? ` s="${styles.idFor(sheet.rowStyles.get(r))}" customFormat="1"` : ''}${sheet.rowHeights.has(r) ? ` ht="${sheet.rowHeights.get(r) * 0.75}" customHeight="1"` : ''}${sheet.hiddenRows.has(r) ? ' hidden="1"' : ''}>${cells.sort((a, b) => a.c - b.c).map(c => c.xml).join('')}</row>`).join('');
    const pane = sheet.freezeRows || sheet.freezeCols ? `<pane xSplit="${sheet.freezeCols}" ySplit="${sheet.freezeRows}" topLeftCell="${address(sheet.freezeRows, sheet.freezeCols)}" activePane="${sheet.freezeRows && sheet.freezeCols ? 'bottomRight' : sheet.freezeRows ? 'bottomLeft' : 'topRight'}" state="frozen"/>` : '';
    files[`xl/worksheets/sheet${i + 1}.xml`] = xmlHeader + `<worksheet xmlns="${spreadsheetNS}"><dimension ref="A1:${address(sheet.usedRange().r2, sheet.usedRange().c2)}"/><sheetViews><sheetView workbookViewId="0" showGridLines="${sheet.gridlines ? 1 : 0}">${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="20.25"/>${sheet.colWidths.size || sheet.colStyles.size || sheet.hiddenCols.size ? '<cols>' + [...new Set([...sheet.colWidths.keys(), ...sheet.colStyles.keys(), ...sheet.hiddenCols])].sort((a,b) => a-b).map(c => `<col min="${c + 1}" max="${c + 1}"${sheet.colWidths.has(c) ? ` width="${Math.max(1, (sheet.colWidths.get(c) - 5) / 7)}" customWidth="1"` : ''}${sheet.hiddenCols.has(c) ? ' hidden="1"' : ''}${sheet.colStyles.has(c) ? ` style="${styles.idFor(sheet.colStyles.get(c))}"` : ''}/>`).join('') + '</cols>' : ''}<sheetData>${rowXML}</sheetData>${sheet.filters ? `<autoFilter ref="${address(sheet.filters.range.r1, sheet.filters.range.c1)}:${address(sheet.filters.range.r2, sheet.filters.range.c2)}"/>` : ''}${sheet.merges.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map(m => `<mergeCell ref="${address(m.r1, m.c1)}:${address(m.r2, m.c2)}"/>`).join('')}</mergeCells>` : ''}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
  });
  const compressed = {};
  // Native raw DEFLATE is ZIP method 8. Older browsers retain the valid stored ZIP path.
  let supported = false;
  try { new CompressionStream('deflate-raw'); supported = true; } catch {}
  if (supported) for (const [name, source] of Object.entries(files)) {
    const data = encoder.encode(source);
    const packed = new Uint8Array(await new Response(new Blob([data]).stream().pipeThrough(new CompressionStream('deflate-raw'))).arrayBuffer());
    if (packed.length < data.length) compressed[name] = packed;
  }
  return zipStore(files, compressed);
}
function readXML(bytes) {
  if (!bytes) throw new Error('The workbook is missing a required XML part.'); const source = decoder.decode(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw new Error('DTD declarations are not allowed in workbook XML.');
  const doc = new DOMParser().parseFromString(source, 'application/xml'); if (doc.getElementsByTagName('parsererror').length) throw new Error('Malformed workbook XML.'); return doc;
}
const elements = (node, name) => [...node.getElementsByTagNameNS('*', name)];
const firstElement = (node, name) => elements(node, name)[0];
const builtinFormats = { 0: 'general', 1: 'integer', 2: 'number', 3: 'integer', 4: 'number', 9: '0%', 10: '0.00%', 14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'hh:mm', 21: 'hh:mm:ss', 22: 'm/d/yy h:mm', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 49: '@', 44: 'currency' };
export async function importXLSX(buffer, title = 'Imported workbook') {
  const parts = await unzip(buffer), workbookDoc = readXML(parts.get('xl/workbook.xml')), relsDoc = readXML(parts.get('xl/_rels/workbook.xml.rels'));
  const dateSystem = firstElement(workbookDoc, 'workbookPr')?.getAttribute('date1904');
  if (dateSystem === '1' || dateSystem === 'true') throw new Error('Workbooks using the 1904 date system are not supported. Convert the workbook to the 1900 date system before importing.');
  const relationships = new Map(elements(relsDoc, 'Relationship').filter(e => e.getAttribute('TargetMode') !== 'External').map(e => [e.getAttribute('Id'), e.getAttribute('Target')]));
  const pathFor = target => target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
  const warnings = ['XLSX import supports cell values, formulas, basic styles, dimensions, merged cells and frozen panes. Charts, pivots, macros, advanced formatting and external links are not imported.'];
  let strings = [], styles = [{}];
  const sharedPath = [...relationships.values()].find(x => x.endsWith('sharedStrings.xml'));
  if (sharedPath && parts.has(pathFor(sharedPath))) strings = elements(readXML(parts.get(pathFor(sharedPath))), 'si').map(si => elements(si, 't').map(t => t.textContent).join(''));
  const stylePath = [...relationships.values()].find(x => x.endsWith('styles.xml'));
  if (stylePath && parts.has(pathFor(stylePath))) {
    const doc = readXML(parts.get(pathFor(stylePath))), formats = new Map(elements(doc, 'numFmt').map(e => [+e.getAttribute('numFmtId'), e.getAttribute('formatCode')]));
    const fonts = [...(firstElement(doc, 'fonts')?.children ?? [])], fills = [...(firstElement(doc, 'fills')?.children ?? [])], xfs = [...(firstElement(doc, 'cellXfs')?.children ?? [])];
    const cssColor = el => { const rgb = el?.getAttribute('rgb'); return rgb ? '#' + rgb.slice(-6) : undefined; };
    styles = xfs.map(xf => {
      const font = fonts[+xf.getAttribute('fontId')], fill = fills[+xf.getAttribute('fillId')], alignment = firstElement(xf, 'alignment'), id = +xf.getAttribute('numFmtId');
      const style = xf.getAttribute('applyNumberFormat') === '0' && id === 0 ? {} : numberFormatStyle(formats.get(id) || builtinFormats[id] || 'general');
      if (font) { const family = firstElement(font, 'name')?.getAttribute('val'); if (family) style.fontFamily = family; for (const [key,tag] of [['bold','b'],['italic','i'],['underline','u']]) { const flag = firstElement(font,tag); style[key] = !!flag && !['0','false','none'].includes(flag.getAttribute('val')); } const fs = firstElement(font, 'sz')?.getAttribute('val'); if (fs) style.fontSize = Math.min(72, +fs / 0.75); const color = cssColor(firstElement(font, 'color')); if (color) style.color = color; }
      const color = fill ? cssColor(firstElement(fill, 'fgColor')) : null; if (color && firstElement(fill, 'patternFill')?.getAttribute('patternType') === 'solid') style.fill = color;
      if (alignment) { const a = alignment.getAttribute('horizontal'); if (['left', 'center', 'right'].includes(a)) style.align = a; style.wrap = alignment.getAttribute('wrapText') === '1'; }
      if (+xf.getAttribute('borderId') > 0) style.border = true;
      return style;
    });
  }
  const wb = new Workbook(); wb.title = title; wb.sheets = []; let count = 0;
  for (const node of elements(workbookDoc, 'sheet')) {
    if (wb.sheets.length >= 256) throw new Error('Too many worksheets.');
    const id = node.getAttributeNS(relationshipNS, 'id') || node.getAttribute('r:id'), target = relationships.get(id); if (!target) continue;
    const doc = readXML(parts.get(pathFor(target))), sheet = new Sheet(node.getAttribute('name') || 'Sheet'), shared = new Map();
    for (const cell of elements(doc, 'c')) {
      if (++count > 200000) throw new Error('XLSX import is limited to 200,000 populated cells.');
      const ref = parseAddress(cell.getAttribute('r')); if (!ref) continue;
      const type = cell.getAttribute('t'), v = firstElement(cell, 'v')?.textContent ?? '', f = firstElement(cell, 'f'); let raw = v;
      if (f) {
        if (f.getAttribute('t') === 'shared') {
          const si = f.getAttribute('si'); if (f.textContent) { raw = '=' + f.textContent; shared.set(si, { raw, ...ref }); }
          else if (shared.has(si)) { const base = shared.get(si); raw = shiftFormula(base.raw, ref.r - base.r, ref.c - base.c); }
          else { raw = v; warnings.push(`Unresolved shared formula at ${sheet.name}!${cell.getAttribute('r')}; used cached value.`); }
        } else raw = '=' + f.textContent;
      } else if (type === 's') raw = "'" + (strings[+v] ?? '');
      else if (type === 'inlineStr') raw = "'" + elements(firstElement(cell, 'is') || cell, 't').map(t => t.textContent).join('');
      else if (type === 'str') raw = "'" + v;
      else if (type === 'b') raw = v === '1' ? 'TRUE' : 'FALSE';
      sheet.cells.set(keyOf(ref.r, ref.c), { raw, ...(cell.hasAttribute('s') ? { style: structuredClone(styles[+cell.getAttribute('s')] || {}) } : {}) });
    }
    for (const row of elements(doc, 'row')) { const r = +row.getAttribute('r') - 1; if (r < 0 || r >= 1048576) continue; if (row.hasAttribute('s')) sheet.rowStyles.set(r, structuredClone(styles[+row.getAttribute('s')] || {})); if (row.hasAttribute('ht')) sheet.rowHeights.set(r, Math.max(16, Math.min(400, +row.getAttribute('ht') / 0.75))); if (row.getAttribute('hidden') === '1') sheet.hiddenRows.add(r); }
    for (const col of elements(doc, 'col')) { const min = +col.getAttribute('min') - 1, max = Math.min(16383, +col.getAttribute('max') - 1), width = +col.getAttribute('width') * 7 + 5; for (let c = Math.max(0, min); c <= max; c++) { if (col.getAttribute('hidden') === '1' || col.getAttribute('hidden') === 'true') sheet.hiddenCols.add(c); if (col.hasAttribute('width')) sheet.colWidths.set(c, Math.max(26, Math.min(1200, width))); if (col.hasAttribute('style')) sheet.colStyles.set(c, structuredClone(styles[+col.getAttribute('style')] || {})); } }
    for (const cell of elements(doc, 'mergeCell')) { const q = parseRange(cell.getAttribute('ref')); if (q) sheet.merges.push(q); }
    const pane = firstElement(doc, 'pane'); if (pane?.getAttribute('state') === 'frozen') { sheet.freezeRows = Math.min(1000, +pane.getAttribute('ySplit') || 0); sheet.freezeCols = Math.min(100, +pane.getAttribute('xSplit') || 0); }
    const view = firstElement(doc, 'sheetView'); sheet.gridlines = view?.getAttribute('showGridLines') !== '0';
    const filter = firstElement(doc, 'autoFilter'), range = filter && parseRange(filter.getAttribute('ref')); if (range) sheet.filters = { range, criteria: {} };
    wb.sheets.push(sheet);
  }
  if (!wb.sheets.length) throw new Error('No readable worksheets were found.');
  for (const name of elements(workbookDoc, 'definedName')) { const key = name.getAttribute('name'); if (!key.startsWith('_xlnm.')) wb.names[key.toUpperCase()] = name.textContent; }
  const active = +firstElement(workbookDoc, 'workbookView')?.getAttribute('activeTab') || 0; wb.activeSheetId = (wb.sheets[active] || wb.sheets[0]).id;
  return { workbook: wb, warnings };
}
