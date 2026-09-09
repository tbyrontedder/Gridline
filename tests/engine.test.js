import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, Sheet, FormulaError, FUNCTIONS, FormulaParser, address, parseAddress, parseRange, rangeAddress, shiftFormula, cellsIn, formatValue, dateSerial, serialDate, MAX_ROWS, MAX_COLS, keyOf } from '../src/engine.js';
import { createSampleWorkbook } from '../src/sample.js';
import { parseDelimited, serializeDelimited, workbookFromCSV, exportCSV, zipStore, unzip, exportXLSX } from '../src/io.js';
import { AxisLayout } from '../src/renderer.js';
const make = () => new Workbook();
function evaluate(formula, cells={}) {
 const wb=make(),sheet=wb.activeSheet;
 wb.transaction('fixture',()=>{for(const [ref,raw]of Object.entries(cells)){const p=parseAddress(ref);wb.setRaw(sheet,p.r,p.c,String(raw));}wb.setRaw(sheet,100,20,formula);});
 return wb.value(sheet,100,20);
}
function code(result){return result instanceof FormulaError?result.code:result;}
for(const [formula,expected]of [
 ['=1+2*3',7],['=(1+2)*3',9],['=2^3^2',64],['=-2^2',4],['=50%*240',120],['="hello"&" "&"world"','hello world'],['="a"="A"',true],['=5<>4',true],['=3<=3',true],
 ['=SUM(1,2,3)',6],['=AVERAGE(2,4,9)',5],['=MIN(9,2,-3)',-3],['=MAX(9,2,-3)',9],['=PRODUCT(2,3,4)',24],['=COUNT(1,2,"x")',2],['=COUNTA(1,"x","")',3],['=MEDIAN(2,9,3,6)',4.5],
 ['=ABS(-3)',3],['=SQRT(81)',9],['=POWER(2,10)',1024],['=MOD(-3,2)',1],['=INT(-1.2)',-2],['=ROUND(-1.25,1)',-1.3],['=ROUND(125,-1)',130],['=ROUNDUP(-1.21,1)',-1.3],['=ROUNDDOWN(-1.29,1)',-1.2],['=TRUNC(-1.8)',-1],
 ['=IF(FALSE,1/0,42)',42],['=IF(TRUE,42,UNKNOWN())',42],['=IF(FALSE,1)',false],['=IFERROR(1/0,7)',7],['=IFNA(NA(),3)',3],['=IFS(FALSE,1,2>1,9)',9],['=ISERROR(1/0)',true],['=ISNA(NA())',true],['=CHOOSE(2,1/0,42)',42],
 ['=AND(TRUE,TRUE)',true],['=OR(FALSE,TRUE)',true],['=NOT(FALSE)',true],['=XOR(TRUE,TRUE)',false],['=ISNUMBER(42)',true],['=ISTEXT("42")',true],['=ISBLANK(A99)',true],['=N(TRUE)',1],
 ['=LEN("Gridline")',8],['=LEFT("Gridline",4)','Grid'],['=RIGHT("Gridline",4)','line'],['=MID("Gridline",3,4)','idli'],['=TRIM("  clean   text ")','clean text'],['=UPPER("word")','WORD'],['=LOWER("WORD")','word'],['=PROPER("hello WORLD")','Hello World'],['=CONCAT("a","b","c")','abc'],['=TEXTJOIN("-",TRUE,"a","","b")','a-b'],['=SUBSTITUTE("a-a-a","a","x",2)','a-x-a'],['=FIND("line","Gridline")',5],['=SEARCH("LINE","Gridline")',5],['=REPT("ab",3)','ababab'],['=VALUE("$1,200")',1200],
 ['=DATE(2026,1,1)',46023],['=YEAR(DATE(2026,9,6))',2026],['=MONTH(DATE(2026,9,6))',9],['=DAY(DATE(2026,9,6))',6],['=ROW(C7)',7],['=COLUMN(C7)',3],['=ROWS(A1:B3)',3],['=COLUMNS(A1:B3)',2],
 ['=1/0','#DIV/0!'],['=UNKNOWN(3)','#NAME?'],['=SQRT(-1)','#NUM!'],['=SUM(1,)',1],['=LEFT("a",-1)','#VALUE!'],['=1+','#ERROR!'],['=IFNA(1/0,7)','#DIV/0!']
])test(`formula ${formula}`,()=>{const actual=code(evaluate(formula));assert.equal(actual,expected);});

test('A1 addressing round-trips sheet boundaries',()=>{assert.deepEqual(parseAddress('XFD1048576'),{r:MAX_ROWS-1,c:MAX_COLS-1});assert.equal(parseAddress('XFE1'),null);assert.equal(address(MAX_ROWS-1,MAX_COLS-1),'XFD1048576');assert.equal(rangeAddress(parseRange('D9:B2')),'B2:D9');});
test('formula fill respects absolute and mixed refs and quoted strings',()=>assert.equal(shiftFormula('=SUM(A1:$B$5)+\'My sheet\'!$C2+"A1"',3,2),'=SUM(C4:$B$5)+\'My sheet\'!$C5+"A1"'));
test('formula fill does not rewrite sheet-like cell names or function names',()=>assert.equal(shiftFormula('=LOG10(A1)+Sheet1!B2',1,1),'=LOG10(B2)+Sheet1!C3'));
test('formula fill preserves fully absolute refs',()=>assert.equal(shiftFormula('=$A$1+$B2+C$3',2,4),'=$A$1+$B4+G$3'));
test('negative copied reference becomes a reference error',()=>assert.equal(shiftFormula('=A1',-1,0),'=#REF!'));
test('numeric ranges ignore text and blanks',()=>assert.equal(evaluate('=SUM(A1:A5)',{A1:1,A2:'text',A3:3,A4:"'9"}),4));
test('SUMIF criteria',()=>assert.equal(evaluate('=SUMIF(A1:A3,">=2",B1:B3)',{A1:1,A2:2,A3:3,B1:10,B2:20,B3:30}),50));
test('COUNTIF wildcards',()=>assert.equal(evaluate('=COUNTIF(A1:A3,"a*")',{A1:'Alpha',A2:'apple',A3:'Beta'}),2));
test('SUMIFS multiple criteria',()=>assert.equal(evaluate('=SUMIFS(C1:C3,A1:A3,"x",B1:B3,">1")',{A1:'x',A2:'x',A3:'y',B1:1,B2:2,B3:3,C1:10,C2:20,C3:30}),20));
test('AVERAGEIF',()=>assert.equal(evaluate('=AVERAGEIF(A1:A3,">1")',{A1:1,A2:2,A3:4}),3));
test('SUMPRODUCT',()=>assert.equal(evaluate('=SUMPRODUCT(A1:A3,B1:B3)',{A1:1,A2:2,A3:3,B1:4,B2:5,B3:6}),32));
test('INDEX MATCH',()=>assert.equal(evaluate('=INDEX(B1:B3,MATCH("b",A1:A3,0))',{A1:'a',A2:'b',A3:'c',B1:10,B2:20,B3:30}),20));
test('exact VLOOKUP',()=>assert.equal(evaluate('=VLOOKUP("b",A1:B3,2,FALSE)',{A1:'a',A2:'b',A3:'c',B1:10,B2:20,B3:30}),20));
test('XLOOKUP default, approximate and not-found values',()=>{const cells={A1:10,A2:20,A3:30,B1:1,B2:2,B3:3};assert.equal(evaluate('=XLOOKUP(20,A1:A3,B1:B3)',cells),2);assert.equal(evaluate('=XLOOKUP(21,A1:A3,B1:B3,0,-1)',cells),2);assert.equal(evaluate('=XLOOKUP(99,A1:A3,B1:B3,"missing")',cells),'missing');});
test('dependency invalidation is transitive and preserves unrelated cache entries',()=>{const w=make(),s=w.activeSheet;w.transaction('setup',()=>{w.setRaw(s,0,0,'2');w.setRaw(s,0,1,'=A1*2');w.setRaw(s,0,2,'=B1+1');w.setRaw(s,0,3,'=99');});assert.equal(w.value(s,0,2),5);w.value(s,0,3);const before=w.engine.evaluations;w.setRaw(s,0,0,'4');assert.equal(w.value(s,0,2),9);assert.equal(w.value(s,0,3),99);assert.equal(w.engine.evaluations-before,2);});
test('dependencies are rewired after formula replacement',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,0,0,'3');w.setRaw(s,0,1,'4');w.setRaw(s,0,2,'=A1');assert.equal(w.value(s,0,2),3);w.setRaw(s,0,2,'=B1');assert.equal(w.value(s,0,2),4);const n=w.engine.evaluations;w.setRaw(s,0,0,'9');assert.equal(w.value(s,0,2),4);assert.equal(w.engine.evaluations,n);});
test('cycles report #CYCLE! and recover after editing',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,0,0,'=B1');w.setRaw(s,0,1,'=A1');assert.equal(code(w.value(s,0,0)),'#CYCLE!');w.setRaw(s,0,1,'8');assert.equal(w.value(s,0,0),8);});
test('cross-sheet dependencies and renamed sheet references',()=>{const w=make(),s=w.activeSheet,t=w.addSheet('Sales Data');w.setRaw(t,0,0,'7');w.setRaw(s,0,0,"='Sales Data'!A1*2");assert.equal(w.value(s,0,0),14);w.renameSheet(t,'New name');assert.equal(s.raw(0,0),"='New name'!A1*2");w.setRaw(t,0,0,'8');assert.equal(w.value(s,0,0),16);});
test('atomic transaction rollback',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,0,0,'original');assert.throws(()=>w.transaction('fail',()=>{w.setRaw(s,0,0,'bad');w.setRaw(s,0,1,'bad');throw new Error('stop');}));assert.equal(s.raw(0,0),'original');assert.equal(s.raw(0,1),'');});
test('batch undo redo is one history command',()=>{const w=make(),s=w.activeSheet;w.transaction('two edits',()=>{w.setRaw(s,0,0,'1');w.setRaw(s,0,1,'2');});assert.equal(w.undoStack.length,1);w.undo();assert.equal(s.raw(0,0),'');assert.equal(s.raw(0,1),'');w.redo();assert.equal(s.raw(0,0),'1');assert.equal(s.raw(0,1),'2');});
test('editing after undo invalidates redo branch',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,0,0,'1');w.setRaw(s,0,0,'2');w.undo();w.setRaw(s,0,0,'3');assert.equal(w.redo(),false);});
test('styles undo without changing formulas',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,0,0,'=1+2');w.applyStyle(s,parseRange('A1'),{bold:true});assert.equal(w.value(s,0,0),3);w.undo();assert.equal(s.get(0,0).style,undefined);});
test('row insertion shifts cells and absolute references',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,1,0,'8');w.setRaw(s,4,2,'=$A$2*2');w.structuralEdit(s,'row',1,1);assert.equal(s.raw(2,0),'8');assert.equal(s.raw(5,2),'=$A$3*2');assert.equal(w.value(s,5,2),16);});
test('deleted direct reference becomes #REF!',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,1,0,'8');w.setRaw(s,4,2,'=A2');w.structuralEdit(s,'row',1,-1);assert.equal(s.raw(3,2),'=#REF!');assert.equal(code(w.value(s,3,2)),'#REF!');});
test('deleting range endpoint shrinks the range',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,0,0,'1');w.setRaw(s,1,0,'2');w.setRaw(s,2,0,'3');w.setRaw(s,0,2,'=SUM(A1:A3)');w.structuralEdit(s,'row',2,-1);assert.equal(s.raw(0,2),'=SUM(A1:A2)');assert.equal(w.value(s,0,2),3);});
test('column insertion adjusts cross-sheet references',()=>{const w=make(),s=w.activeSheet,t=w.addSheet('Input');w.setRaw(t,0,1,'8');w.setRaw(s,0,0,'=Input!$B$1');w.structuralEdit(t,'column',0,1);assert.equal(s.raw(0,0),'=Input!$C$1');assert.equal(w.value(s,0,0),8);});
test('two-value seed fill creates a numeric series',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,0,0,'10');w.setRaw(s,1,0,'15');w.fill(s,parseRange('A1:A2'),parseRange('A1:A5'));assert.equal(s.raw(4,0),'30');w.undo();assert.equal(s.raw(4,0),'');});
test('formula fill translates relative references',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,0,1,'=A1*2');w.fill(s,parseRange('B1'),parseRange('B1:B4'));assert.equal(s.raw(3,1),'=A4*2');});
test('stable table sort preserves headers and adjusts row formulas',()=>{const w=make(),s=w.activeSheet;[['Name','Value','Double'],['C','3','=B2*2'],['A','1','=B3*2'],['B','2','=B4*2']].forEach((row,r)=>row.forEach((v,c)=>w.setRaw(s,r,c,v)));w.sort(s,parseRange('A1:C4'),0,false,true);assert.equal(s.raw(0,0),'Name');assert.equal(s.raw(1,0),'A');assert.equal(s.raw(1,2),'=B2*2');assert.equal(w.value(s,1,2),2);});
test('Gridline JSON round-trip preserves values and map-backed metadata',()=>{const w=createSampleWorkbook(),copy=Workbook.fromJSON(JSON.parse(JSON.stringify(w.toJSON())));assert.equal(copy.value(copy.activeSheet,5,1),2187000);assert.equal(copy.activeSheet.charts.length,2);assert.equal(copy.activeSheet.colWidths.get(1),174);assert.equal(copy.names.COST_RATIO,'Assumptions!$B$3');});
test('malformed workbook and invalid cell addresses are rejected',()=>{assert.throws(()=>Workbook.fromJSON({}));const d=make().toJSON();d.sheets[0].cells=[['-1,0',{raw:'bad'}]];assert.throws(()=>Workbook.fromJSON(d));});
test('sample has no calculation errors in all stored cells',()=>{const w=createSampleWorkbook();for(const s of w.sheets)for(const [key]of s.cells){const [r,c]=key.split(',').map(Number),v=w.value(s,r,c);assert.ok(!(v instanceof FormulaError),`${s.name}!${address(r,c)}: ${v}`);}});
test('sample cross-sheet assumptions recalculate all dependents',()=>{const w=createSampleWorkbook(),s=w.sheetByName('Sales data'),a=w.sheetByName('Assumptions');assert.ok(Math.abs(w.value(s,1,9)-.42)<1e-12);w.setRaw(a,2,1,'50%');assert.equal(w.value(s,1,9),.5);});
test('CSV parser handles escaped quotes, multiline values, delimiters and BOM',()=>{const rows=parseDelimited('\ufeffA,B,C\r\n1,"two, words","line1\nline2"\r\n2,"say ""hi""",3\r\n');assert.deepEqual(rows,[['A','B','C'],['1','two, words','line1\nline2'],['2','say "hi"','3']]);assert.deepEqual(parseDelimited('a;b\n1;2'),[['a','b'],['1','2']]);});
test('TSV serialization round-trips quoted content',()=>{const rows=[['a\tb','say "hi"','line1\nline2'],['1','2','3']];assert.deepEqual(parseDelimited(serializeDelimited(rows,'\t'),'\t'),rows);});
test('CSV input neutralizes formula injection',()=>{const w=workbookFromCSV('Header\n=1+1');assert.equal(w.activeSheet.raw(1,0),"'=1+1");assert.equal(w.value(w.activeSheet,1,0),'=1+1');assert.ok(exportCSV(w).includes("'=1+1"));});
test('ZIP writer/reader round-trip and CRC integrity',async()=>{const zip=zipStore({'a.txt':'hello','folder/b.txt':'world'});const parts=await unzip(zip);assert.equal(new TextDecoder().decode(parts.get('a.txt')),'hello');const corrupt=zip.slice();corrupt[35]^=1;await assert.rejects(()=>unzip(corrupt),/integrity/);});
test('XLSX output has all sheets, formulas, and styles',async()=>{const w=createSampleWorkbook(),parts=await unzip(exportXLSX(w));assert.ok(parts.has('[Content_Types].xml'));assert.ok(parts.has('xl/styles.xml'));assert.ok(parts.has('xl/worksheets/sheet3.xml'));const xml=new TextDecoder().decode(parts.get('xl/worksheets/sheet1.xml'));assert.ok(xml.includes('<f>SUM(G12:G19)</f>'));assert.ok(xml.includes('2187000'));});
test('axis prefix positions, binary search, and hidden rows',()=>{const axis=new AxisLayout(1000,27,new Map([[1,40],[9,20]]),new Set([3,4]));assert.equal(axis.offset(2),67);assert.equal(axis.offset(5),94);assert.equal(axis.find(94),5);assert.equal(axis.offset(1000),26952);});
test('range operations have explicit bounds',()=>assert.throws(()=>[...cellsIn(parseRange('A1:XFD1048576'))],/limit/));
test('date conversion round-trips contemporary dates',()=>{const date=new Date(Date.UTC(2026,8,6));assert.equal(serialDate(dateSerial(date)).getTime(),date.getTime());});
test('number formatting preserves documented formats',()=>{assert.equal(formatValue(12345.6,{format:'currency',decimals:2}),'$12,345.60');assert.equal(formatValue(.125,{format:'percent',decimals:1}),'12.5%');assert.equal(formatValue(new FormulaError('#REF!')),'#REF!');});

test('whole-column dependencies observe new rows beyond previous used range',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,0,0,'2');w.setRaw(s,0,1,'=SUM(A:A)');assert.equal(w.value(s,0,1),2);w.setRaw(s,999,0,'5');assert.equal(w.value(s,0,1),7);});
test('ISNUMBER/ISTEXT return false on error values',()=>{assert.equal(evaluate('=ISNUMBER(1/0)'),false);assert.equal(evaluate('=ISTEXT(NA())'),false);});
test('a formula referencing an empty cell yields numeric zero',()=>assert.equal(evaluate('=A99'),0));
test('non-finite literal numbers produce an explicit error',()=>{const w=make(),s=w.activeSheet;w.setRaw(s,0,0,'1e999');assert.equal(code(w.value(s,0,0)),'#NUM!');});


test('renaming a worksheet retargets defined names',()=>{const w=createSampleWorkbook(),s=w.sheetByName('Assumptions');w.renameSheet(s,'Model assumptions');assert.equal(w.names.COST_RATIO,"'Model assumptions'!$B$3");const t=w.sheetByName('Sales data');w.setRaw(t,0,20,'=COST_RATIO');assert.equal(w.value(t,0,20),.58);});
test('JSON metadata cannot replace sheet methods',()=>{const d=make().toJSON();d.sheets[0].get='unsafe';d.sheets[0].toJSON='unsafe';const w=Workbook.fromJSON(d);assert.equal(typeof w.activeSheet.get,'function');assert.equal(typeof w.activeSheet.toJSON,'function');});
test('function registry exposes the documented 83 names',()=>assert.equal(FUNCTIONS.size,83));

test('column formats stay sparse, preserve values and appearance, and survive JSON and history', () => {
  const wb = make(); let s = wb.activeSheet;
  wb.setRaw(s, 1, 1, '=2+3'); wb.applyStyle(s, parseRange('B2'), { bold:true, format:'percent' });
  wb.applyStyle(s, {r1:0,r2:MAX_ROWS-1,c1:1,c2:1}, {format:'currency',currency:'EUR',decimals:2});
  assert.equal(s.cells.size, 1); assert.equal(s.colStyles.size, 1); assert.equal(s.style(1,1).bold, true); assert.equal(s.raw(1,1), '=2+3');
  assert.equal(wb.display(s,1,1), '€5.00'); assert.deepEqual(s.usedRange(), {r1:0,c1:0,r2:1,c2:1});
  wb.undo(); s = wb.activeSheet; assert.equal(s.colStyles.size,0); assert.equal(wb.display(s,1,1),'500.0%');
  wb.redo(); s = wb.activeSheet;
  const restored = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON()))); s = restored.activeSheet;
  restored.setRaw(s,MAX_ROWS-1,1,'12.3'); assert.equal(restored.display(s,MAX_ROWS-1,1),'€12.30'); assert.equal(s.cells.size,2);
});
test('row defaults, cell overrides, and structural edits keep their coordinates', () => {
  const wb = make(); let s = wb.activeSheet;
  wb.applyStyle(s,{r1:0,r2:MAX_ROWS-1,c1:2,c2:2},{format:'currency'});
  wb.applyStyle(s,{r1:4,r2:4,c1:0,c2:MAX_COLS-1},{format:'percent',decimals:2});
  assert.equal(s.cells.size,0); wb.setRaw(s,4,2,'0.5'); assert.equal(wb.display(s,4,2),'50.00%');
  wb.applyStyle(s,parseRange('C5'),{format:'number',decimals:1}); assert.equal(wb.display(s,4,2),'0.5');
  wb.structuralEdit(s,'column',1,1); assert.equal(s.colStyles.has(3),true); assert.equal(s.colStyles.has(2),false);
  wb.structuralEdit(s,'row',3,1); assert.equal(s.rowStyles.has(5),true); assert.equal(wb.display(s,5,3),'0.5');
  wb.structuralEdit(s,'row',5,-1); assert.equal(s.rowStyles.size,0);
});
test('date and time presets display correctly and edit as readable inputs without losing formulas', () => {
  const wb = make(), s = wb.activeSheet;
  wb.applyStyle(s,{r1:0,r2:MAX_ROWS-1,c1:1,c2:1},{format:'date',pattern:'mm/dd/yy'});
  wb.setRaw(s,0,1,'2024-02-29'); assert.equal(wb.display(s,0,1),'02/29/24'); assert.equal(wb.editValue(s,0,1),'2024-02-29');
  wb.setRaw(s,1,1,'9/8/2026 13:45:30'); assert.equal(wb.editValue(s,1,1),'2026-09-08 13:45:30');
  wb.setRaw(s,2,1,'=B1+1'); assert.equal(wb.editValue(s,2,1),'=B1+1'); assert.equal(wb.display(s,2,1),'03/01/24');
  assert.throws(()=>wb.setRaw(s,0,1,'2025-02-29'), /valid date/); assert.equal(wb.editValue(s,0,1),'2024-02-29');
  assert.equal(formatValue(60,{format:'date',pattern:'yyyy-mm-dd'}),'1900-02-29');
  assert.equal(formatValue(45292.5625,{format:'time',pattern:'h:mm AM/PM'}),'1:30 PM');
  assert.equal(formatValue(1.5,{format:'time',pattern:'[h]:mm:ss'}),'36:00:00');
  wb.applyStyle(s,parseRange('C1'),{format:'time',pattern:'hh:mm:ss'}); wb.setRaw(s,0,2,'1:30:00 PM'); assert.equal(wb.display(s,0,2),'13:30:00');
  assert.equal(wb.value(s,0,2),.5625); assert.throws(()=>wb.setRaw(s,0,2,'25:00'), /valid time/);
});
test('format changes reset old options, text preserves leading zeros, and fill copies inherited formatting', () => {
  const wb = make(), s = wb.activeSheet;
  wb.applyStyle(s,{r1:0,r2:MAX_ROWS-1,c1:0,c2:0},{format:'currency',currency:'GBP',decimals:4});
  wb.setRaw(s,0,0,'42'); wb.fill(s,parseRange('A1'),parseRange('A1:B1')); assert.equal(wb.display(s,0,1),'£42.0000');
  wb.applyStyle(s,parseRange('A1'),{format:'general'}); assert.equal(wb.display(s,0,0),'42');
  const restored = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON()))); assert.equal(restored.display(restored.activeSheet,0,0),'42');
  wb.applyStyle(s,parseRange('C1'),{format:'text'}); wb.setRaw(s,0,2,'00123'); assert.equal(wb.value(s,0,2),'00123');
});
test('XLSX includes sparse row and column formats, even with no stored cells', async () => {
  const wb = make(), s = wb.activeSheet;
  wb.applyStyle(s,{r1:0,r2:MAX_ROWS-1,c1:1,c2:1},{format:'date',pattern:'yyyy-mm-dd'});
  wb.applyStyle(s,{r1:3,r2:3,c1:0,c2:MAX_COLS-1},{format:'currency',currency:'GBP',decimals:2});
  const parts = await unzip(exportXLSX(wb)), decode = name => new TextDecoder().decode(parts.get(name));
  assert.match(decode('xl/worksheets/sheet1.xml'), /<col min="2" max="2" style="\d+"\/>/);
  assert.match(decode('xl/worksheets/sheet1.xml'), /<row r="4" s="\d+" customFormat="1">/);
  assert.match(decode('xl/styles.xml'), /yyyy-mm-dd/); assert.match(decode('xl/styles.xml'), /£/);
  assert.equal(s.cells.size,0);
});

test('clearing a column removes its defaults and undo restores them', () => {
  const wb = make(); let s = wb.activeSheet; const column = {r1:0,r2:MAX_ROWS-1,c1:1,c2:1};
  wb.applyStyle(s,column,{format:'date',pattern:'yyyy-mm-dd'}); wb.setRaw(s,0,1,'2026-09-08');
  wb.clear(s,column,true); assert.equal(s.cells.size,0); assert.equal(s.colStyles.size,0);
  wb.undo(); s = wb.activeSheet; assert.equal(wb.display(s,0,1),'2026-09-08'); assert.equal(s.colStyles.size,1);
});


test('decimal controls on integer formats are retained in XLSX', async () => {
  const wb = make(), s = wb.activeSheet;
  wb.setRaw(s,0,0,'12.34'); wb.applyStyle(s,parseRange('A1'),{format:'integer',decimals:2});
  assert.equal(wb.display(s,0,0),'12.34');
  const parts = await unzip(exportXLSX(wb));
  assert.match(new TextDecoder().decode(parts.get('xl/styles.xml')), /formatCode="#,##0.00"/);
});

test('native import rejects invalid decimal settings in cell, row, and column styles', () => {
  for (const location of ['cells','rowStyles','colStyles']) for (const decimals of ['0"><img src=x onerror=alert(1)>', '2', -1, 11, 1.5, true, {}, [], NaN, Infinity]) {
    const data = make().toJSON(), style = {format:'number', decimals};
    data.sheets[0][location] = location === 'cells' ? [['0,0',{raw:'',style}]] : [[0,style]];
    assert.throws(() => Workbook.fromJSON(data), /Invalid decimal places/, `${location}: ${JSON.stringify(decimals)}`);
  }
});
test('native import retains supported decimal settings and older styles without precision', () => {
  for (const decimals of [undefined, null, ...Array.from({length:11}, (_,i) => i)]) {
    const data = make().toJSON(), style = {format:'number', decimals};
    data.sheets[0].cells = [['0,0',{raw:'42',style}]];
    data.sheets[0].rowStyles = [[1,style]]; data.sheets[0].colStyles = [[1,style]];
    const sheet = Workbook.fromJSON(JSON.parse(JSON.stringify(data))).activeSheet;
    assert.equal(sheet.get(0,0).style.decimals,decimals); assert.equal(sheet.rowStyles.get(1).decimals,decimals); assert.equal(sheet.colStyles.get(1).decimals,decimals);
  }
});
