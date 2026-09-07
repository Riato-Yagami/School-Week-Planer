const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('node:assert/strict');
const XLSX = require('xlsx-js-style');
const { unzipSync, strFromU8 } = require('fflate');
const root = process.cwd();
const context = { require, console, __basedir: root, config: require(path.join(root, 'config/config')), module: { exports: {} }, fun: { blendColors: require(path.join(root, 'functions/blendColors')) } };
vm.createContext(context);
vm.runInContext(fs.readFileSync('functions/generateSpreadsheet.js', 'utf8'), context);
vm.runInContext(`
Object.assign(weeks, { first: 1, last: 2, length: 5, startDate: new Date(2026, 8, 7), holidays: [], offDays: [], offDayLabels: {} });
initializeOffDays();
const sample = { name: '4e3', monday: { class: [{ time: 8, duration: 0.75, comment: 'VIE DE CLASSE' }, { time: 9, duration: 1.5 }] } };
const content = [];
addDay(content, [sample], getDay(1, 0, weeks.startDate));
globalThis.comment = content[0][5].v;
globalThis.stats = getSummaryStats(content, [sample, sample]);
globalThis.sheet = generateWeeksSheet([sample]);
`, context);
assert.equal(context.comment, '[vdc] / ');
assert.equal(context.stats.daysTotal, 1);
assert.equal(context.stats.total * 24, 4.5);
const cells = Object.values(context.sheet);
assert(cells.some(c => c && c.z === '[h]"h"mm"min"'));
assert.equal(XLSX.SSF.format('[h]"h"mm"min"', 25.5 / 24), '25h30min');
assert.equal(XLSX.SSF.format('[h]"h"mm"min"', 0.75 / 24), '0h45min');
const file = 'output/planning-college-jean-jaures-m-pesin-2026-2027.xlsx';
const zip = unzipSync(fs.readFileSync(file));
const xml = strFromU8(zip['xl/worksheets/sheet1.xml']);
const styles = strFromU8(zip['xl/styles.xml']);
const baseStyles = styles.replace(/<dxfs[\s\S]*?<\/dxfs>/, '');
assert(!baseStyles.includes('FFE699')); // No creation-date yellow in ordinary styles.
assert(xml.includes('INT($C1)=TODAY()'));
assert(xml.includes('WEEKDAY(TODAY(),2)&gt;5'));
assert(styles.includes('<left style="medium"><color rgb="FF'));
assert(styles.match(/<dxf><fill>[\s\S]*?<border>[\s\S]*?<\/dxf>/));
assert(!xml.includes('<pageMargins') || xml.indexOf('<conditionalFormatting') < xml.indexOf('<pageMargins'));
const workbook = XLSX.readFile(file);
const summaryCells = Object.values(workbook.Sheets[workbook.SheetNames[0]]);
assert(summaryCells.some(c => c && c.v === 'Jours'));
assert(summaryCells.some(c => c && c.z === undefined && c.f && c.f.startsWith('IF(')));
console.log('OK: separators, unique days, duration formats, dynamic highlights, workbook summary');


