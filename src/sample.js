import { Workbook, Sheet, keyOf, address } from './engine.js';
export function createSampleWorkbook() {
  const wb = new Workbook(); wb.title = 'Revenue operations · FY2026';
  const overview = new Sheet('Revenue overview'), transactions = new Sheet('Sales data'), assumptions = new Sheet('Assumptions'); wb.sheets = [overview, transactions, assumptions]; wb.activeSheetId = overview.id;
  const put = (s, r, c, raw, style = {}) => s.cells.set(keyOf(r, c), { raw: String(raw), style });
  const merge = (s, r1, c1, r2, c2) => s.merges.push({ r1, c1, r2, c2 });
  const green = '#176b4a', ink = '#223e2f', muted = '#82978a', fill = '#eff6f1';
  const widths = [24, 174, 98, 97, 97, 97, 113, 110, 102, 94, 118, 118, 36]; widths.forEach((w, i) => overview.colWidths.set(i, w));
  overview.rowHeights.set(0, 15); overview.rowHeights.set(1, 38); overview.rowHeights.set(2, 26); overview.rowHeights.set(3, 20); overview.rowHeights.set(4, 23); overview.rowHeights.set(5, 39); overview.rowHeights.set(6, 18); overview.rowHeights.set(7, 23); overview.rowHeights.set(8, 17); overview.rowHeights.set(9, 31); overview.rowHeights.set(10, 31); overview.rowHeights.set(19, 31); overview.rowHeights.set(20, 16); overview.rowHeights.set(21, 31); overview.rowHeights.set(22, 25);
  overview.gridlines = false; overview.dataRegion = { r1: 10, c1: 1, r2: 18, c2: 11 };
  merge(overview, 1, 1, 1, 8); put(overview, 1, 1, 'Revenue overview', { bold: true, fontSize: 27, color: ink });
  merge(overview, 2, 1, 2, 8); put(overview, 2, 1, 'A clear view of performance. A confident next move.', { fontSize: 12, color: muted });
  merge(overview, 1, 9, 1, 11); put(overview, 1, 9, 'Q1 2026  /  OPERATING REVIEW', { fontSize: 10, color: green, bold: true, align: 'right' });
  merge(overview, 2, 9, 2, 11); put(overview, 2, 9, 'ILLUSTRATIVE DEMO DATA', { fontSize: 9, color: muted, align: 'right' });
  const cards = [
    { start: 1, end: 3, label: 'TOTAL REVENUE', value: '=SUM(G12:G19)', format: 'currency', caption: '↑  18.6% vs. previous quarter' },
    { start: 4, end: 6, label: 'TARGET ATTAINMENT', value: '=SUM(G12:G19)/SUM(H12:H19)', format: 'percent', caption: 'On track for a strong year' },
    { start: 7, end: 9, label: 'REVENUE GROWTH', value: '=AVERAGE(J12:J19)', format: 'percent', caption: '↑  Across all sales channels' },
    { start: 10, end: 11, label: 'ACTIVE CHANNELS', value: '=COUNTA(B12:B19)', format: 'integer', caption: '6 regions · 8 owners' }
  ];
  for (const card of cards) {
    for (let r = 4; r <= 7; r++) { merge(overview, r, card.start, r, card.end); for (let c = card.start; c <= card.end; c++) put(overview, r, c, '', { fill }); }
    put(overview, 4, card.start, card.label, { fill, fontSize: 10, bold: true, color: '#65806e' });
    put(overview, 5, card.start, card.value, { fill, fontSize: 29, bold: true, color: green, align: 'left', format: card.format, decimals: card.format === 'percent' ? 1 : 0 });
    put(overview, 7, card.start, card.caption, { fill, fontSize: 10, color: '#628c72' });
  }
  merge(overview, 9, 1, 9, 8); put(overview, 9, 1, 'CHANNEL PERFORMANCE', { bold: true, fontSize: 10, color: green });
  merge(overview, 9, 9, 9, 11); put(overview, 9, 9, 'USD  ·  January — March', { fontSize: 10, align: 'right', color: '#7c8c82' });
  const headers = ['Sales channel', 'Region', 'January', 'February', 'March', 'Q1 actual', 'Q1 target', 'Variance', 'Growth', 'Performance', 'Owner'];
  headers.forEach((h, i) => put(overview, 10, i + 1, h, { bold: true, fontSize: 11, color: '#63816e', fill: '#eaf2ed', align: i >= 2 && i <= 8 ? 'right' : 'left', bottomBorder: '#d4e2d9' }));
  const channels = [
    ['Enterprise', 'North Am.', 168000, 182000, 205000, 525000, 0.242, 'Alex Morgan'],
    ['Direct sales', 'Europe', 124000, 136000, 148000, 390000, 0.187, 'Sophie Chen'],
    ['E-commerce', 'Global', 92000, 108000, 119000, 300000, 0.294, 'James Wilson'],
    ['Partnerships', 'Asia Pac.', 78000, 85000, 96000, 255000, 0.162, 'Olivia Park'],
    ['Mid-market', 'North Am.', 68000, 74000, 79000, 225000, 0.128, 'Liam Patel'],
    ['Resellers', 'Europe', 54000, 59000, 63000, 168000, 0.143, 'Emma Davis'],
    ['Marketplace', 'Global', 41000, 47000, 56000, 135000, 0.317, 'Noah Kim'],
    ['Self-service', 'LatAm', 32000, 35000, 38000, 108000, 0.093, 'Mia Santos']
  ];
  channels.forEach((row, i) => {
    const r = i + 11, n = r + 1, values = [row[0], row[1], row[2], row[3], row[4], `=SUM(D${n}:F${n})`, row[5], `=G${n}-H${n}`, row[6], `=IF(G${n}>=H${n},"Above target","Needs focus")`, row[7]];
    values.forEach((v, j) => put(overview, r, j + 1, v, { fill: i % 2 ? '#f5f8f6' : '#ffffff', fontSize: j === 9 || j === 10 ? 11 : 12, color: j === 0 ? ink : '#53695c', bold: j === 0 || j === 5, format: j >= 2 && j <= 7 ? 'currency' : j === 8 ? 'percent' : 'general' }));
  });
  for (let c = 1; c <= 11; c++) put(overview, 19, c, c === 1 ? 'Total' : c >= 3 && c <= 8 ? `=SUM(${address(11, c)}:${address(18, c)})` : c === 9 ? '=AVERAGE(J12:J19)' : '', { fill: '#eaf3ed', color: green, bold: true, fontSize: 12, format: c === 9 ? 'percent' : c >= 3 && c <= 8 ? 'currency' : 'general', bottomBorder: '#adc9b7' });
  overview.conditionalRules.push({ type: 'positive', range: { r1: 11, r2: 18, c1: 8, c2: 9 } });
  overview.cells.get(keyOf(12, 6)).note = 'Revenue is calculated from the three monthly inputs. Edit any month to see dependent formulas recalculate.';
  merge(overview, 21, 1, 21, 7); put(overview, 21, 1, 'THE QUARTER, AT A GLANCE', { bold: true, fontSize: 10, color: green });
  merge(overview, 22, 1, 22, 11); put(overview, 22, 1, 'Enterprise continues to lead. Marketplace is our fastest-growing channel.', { fontSize: 12, color: '#718779' });
  overview.charts.push({ id: 'chart-trend', title: 'Revenue momentum', subtitle: 'Monthly revenue · USD', type: 'column', range: { r1: 10, c1: 3, r2: 18, c2: 5 }, aggregate: true, row: 24, col: 1, width: 593, height: 230 });
  overview.charts.push({ id: 'chart-channel', title: 'A balanced portfolio', subtitle: 'Q1 revenue by sales channel', type: 'bar', range: { r1: 10, c1: 1, r2: 18, c2: 6 }, labelColumn: 1, valueColumn: 6, row: 24, col: 7, width: 520, height: 230 });
  const dataHeaders = ['Order ID', 'Order date', 'Channel', 'Region', 'Product', 'Units', 'Unit price', 'Revenue', 'Cost', 'Gross margin', 'Owner'];
  dataHeaders.forEach((h, i) => put(transactions, 0, i, h, { fill: green, color: '#ffffff', bold: true, fontSize: 12 }));
  transactions.freezeRows = 1; transactions.rowHeights.set(0, 34); transactions.colWidths.set(0, 120); transactions.colWidths.set(1, 125); transactions.colWidths.set(2, 150); transactions.colWidths.set(4, 150); transactions.colWidths.set(10, 140);
  for (let i = 0; i < 120; i++) {
    const r = i + 1, channel = channels[i % channels.length], units = 8 + (i * 17) % 83, price = [99, 149, 249, 499][i % 4], day = 46023 + (i * 3) % 89;
    const values = [`ORD-${(2001 + i)}`, day, channel[0], channel[1], ['Gridline Team', 'Gridline Pro', 'Gridline Business', 'Gridline Enterprise'][i % 4], units, price, `=F${r + 1}*G${r + 1}`, `=H${r + 1}*Assumptions!$B$3`, `=(H${r + 1}-I${r + 1})/H${r + 1}`, channel[7]];
    values.forEach((v, c) => put(transactions, r, c, v, { fill: i % 2 ? '#f4f8f5' : '#ffffff', format: c === 1 ? 'date' : c >= 6 && c <= 8 ? 'currency' : c === 9 ? 'percent' : 'general', fontSize: 12 }));
  }
  transactions.filters = { range: { r1: 0, c1: 0, r2: 120, c2: 10 }, criteria: {} };
  assumptions.colWidths.set(0, 240); assumptions.colWidths.set(1, 170); assumptions.colWidths.set(2, 520); assumptions.rowHeights.set(0, 36);
  ['MODEL ASSUMPTIONS', 'Value', 'Notes'].forEach((v, c) => put(assumptions, 0, c, v, { fill: green, color: '#ffffff', bold: true }));
  [['Revenue growth target', '20%', 'Illustrative planning input, not a forecast.'], ['Cost of goods sold', '58%', 'Referenced by every order in the Sales data sheet.'], ['Target gross margin', '=1-B3', 'Calculated as one minus the cost ratio.'], ['Reporting currency', 'USD', 'All amounts in this demo use US dollars.'], ['Reporting period', 'Q1 2026', 'January through March.'], ['Model status', 'Demo workbook', 'All data is fictional and provided for product exploration.']].forEach((row, i) => row.forEach((v, c) => put(assumptions, i + 1, c, v, { fill: i % 2 ? '#f4f8f5' : '#ffffff', color: c === 1 ? '#1666b5' : '#53695c', format: c === 1 && i < 3 ? 'percent' : 'general' })));
  wb.names.COST_RATIO = 'Assumptions!$B$3'; wb.engine.invalidate(); return wb;
}
