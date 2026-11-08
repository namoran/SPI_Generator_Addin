/* Source-and-guard test for the "RFC worksheet in this workbook" import
 * path and the applied-RFC fingerprint.
 *
 * Self-contained (no handoff fixtures): a small synthetic RFC lines the
 * interesting cases — gains, decreases, multiple lines folding into one
 * cell, an empty row skipped, a second task.
 *
 *   1. The same RFC data is read two ways: as a 2D values matrix
 *      (what the pane receives from the open workbook — rfc-grid.js,
 *      exactly the code path generate uses), and as a SheetJS worksheet
 *      (what the pane receives from a picked RFC file — rfc-sheetjs.js,
 *      the file path). SpiCore.loadRfc() must produce identical items
 *      from both.
 *   2. rfc-fingerprint.js: same data reordered = same hash; one changed
 *      quantity = different hash; nothing getting fingerprints from
 *      folded or partial data.
 *
 * Run: npm run test:source
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..');

const SpiCore = require(path.join(root, 'src', 'taskpane', 'spi-core.js'));
const RfcGrid = require(path.join(root, 'src', 'taskpane', 'rfc-grid.js'));
const RfcFingerprint = require(path.join(root, 'src', 'taskpane', 'rfc-fingerprint.js'));
const SpiSheetJs = require(path.join(root, 'src', 'taskpane', 'rfc-sheetjs.js'));
const XLSX = require('xlsx');

let failures = 0;
const check = (name, ok, detail) => {
  if (ok) console.log(`${name}: PASS${detail ? ' — ' + detail : ''}`);
  else { failures++; console.log(`${name}: FAIL${detail ? ' — ' + detail : ''}`); }
};

/* A synthetic RFC: rows 22..111, columns A..I, per the FDEP layout.
 * Task in col A (1), pay item col B (2), UOM col F (6), units col H (8),
 * total col I (9). Rows with no task are skipped by loadRfc. */
const STEP = 0.01;
const makeCells = (rows) => rows.map((r) => ({ ...r }));

const rfcCells = makeCells([
  { task: 3, payItem: '521-1-1', uom: 'LF', units: 1000, total: 2000 },
  { task: 3, payItem: '530-4-2', uom: 'TON', units: 4.5, total: 900 },
  { task: 3, payItem: '521-1-1', uom: 'LF', units: -999.99, total: -1999.98 },
  { task: 22, payItem: '630-1-1', uom: 'SF', units: 12, total: 48 },
  { task: 22, payItem: '700-1-1', uom: 'EA', units: null, total: null }
]);

function buildMatrix(perm) {
  const matrix = [];
  for (let i = 0; i < SpiCore.RFC_LAST_ROW - SpiCore.RFC_FIRST_ROW + 1; i++) {
    matrix.push(new Array(9).fill(null));
  }
  const cells = perm ? perm(rfcCells) : rfcCells;
  cells.forEach((r, i) => {
    const mi = r.row === undefined ? i : r.row;
    matrix[mi][0] = r.task;
    matrix[mi][1] = r.payItem;
    matrix[mi][5] = r.uom;
    matrix[mi][7] = r.units;
    matrix[mi][8] = r.total;
  });
  return matrix;
}

/* --- 1: matrix adapter vs SheetJS parse of the same sheet --- */
const sheetAoA = [];
for (let i = 0; i < SpiCore.RFC_FIRST_ROW - 1; i++) sheetAoA.push([]);
buildMatrix().forEach((row, i) => sheetAoA.push([...row]));
for (let i = 0; i < 11; i++) sheetAoA.push([]); // through row 111 and past it

const ws = XLSX.utils.aoa_to_sheet(sheetAoA);
const Wb = { Sheets: { RFC: ws } };
const itemsMatrix = SpiCore.loadRfc(RfcGrid.fromMatrix(buildMatrix(), SpiCore.RFC_FIRST_ROW));
const itemsSheetJs = SpiCore.loadRfc(SpiSheetJs.fromWorksheet(Wb.Sheets.RFC, XLSX));

const sameItems = JSON.stringify(itemsMatrix) === JSON.stringify(itemsSheetJs);
check('loadRfc: matrix vs SheetJS identical', sameItems,
  !sameItems
    ? `matrix=${JSON.stringify(itemsMatrix, null, 1)} sheetjs=${JSON.stringify(itemsSheetJs, null, 1)}`
    : `${itemsMatrix.length} items`);

/* Empty sheet via both paths: loadRfc must come back empty. */
const emptyMatrix = buildMatrix().map((row) => row.map(() => null));
const empItems = SpiCore.loadRfc(RfcGrid.fromMatrix(emptyMatrix, SpiCore.RFC_FIRST_ROW));
check('loadRfc: empty RFC sheet -> no items', empItems.length === 0, `${empItems.length} items`);

/* The taskpane's sheet read reads exactly rows 22..111 columns A..I; make
 * sure the adapter maps that range correctly from a taller matrix too
 * (values at rows the loader never asked for must not reach the items). */
const tallMatrix = [[], [], [], [null]].concat(buildMatrix());
const tall = SpiCore.loadRfc(RfcGrid.fromMatrix(tallMatrix, SpiCore.RFC_FIRST_ROW - 3));
const tallOk = tall.length === itemsMatrix.length;
check('loadRfc: anchored adapter ignores rows above its block', tallOk, !tallOk ? `${tall.length} vs ${itemsMatrix.length} items` : '');

/* --- 2: fingerprints --- */
const fp = (items) => RfcFingerprint.fingerprint(items);

const fpA = fp(itemsMatrix);
const shuffled = buildMatrix(() => [...rfcCells].reverse());
const fpB = fp(SpiCore.loadRfc(RfcGrid.fromMatrix(shuffled, SpiCore.RFC_FIRST_ROW)));
check('fingerprint: row order does not matter', fpA === fpB, fpA);

/* one changed quantity -> different hash */
const qtyMoved = buildMatrix(() => rfcCells.map((r) => r.task === 3 && r.payItem === '521-1-1' && r.units === 1000 ? { ...r, units: 1001 } : r));
const fpC = fp(SpiCore.loadRfc(RfcGrid.fromMatrix(qtyMoved, SpiCore.RFC_FIRST_ROW)));
check('fingerprint: changed quantity differs', fpA !== fpC, fpC);

/* one changed pay item -> different hash */
const itemMoved = buildMatrix(() => rfcCells.map((r) => r.payItem === '700-1-1' ? { ...r, payItem: '700-1-2' } : r));
const fpD = fp(SpiCore.loadRfc(RfcGrid.fromMatrix(itemMoved, SpiCore.RFC_FIRST_ROW)));
check('fingerprint: changed pay item differs', fpA !== fpD, fpD);

/* trailing spaces and null-vs-blank cells are normalization, not data */
const padded = buildMatrix(() => rfcCells.map((r) => ({ ...r, payItem: r.payItem + '  ', uom: r.uom === null ? r.uom : r.uom + '\u00a0' })));
const fpE = fp(SpiCore.loadRfc(RfcGrid.fromMatrix(padded, SpiCore.RFC_FIRST_ROW)));
check('fingerprint: whitespace normalization matches', fpA === fpE, fpE);

/* between 32-bit collisions: hash is 16 hex characters */
check('fingerprint: 64 bits (16 hex chars)', /^[0-9a-f]{16}$/.test(fpA), fpA);

/* the SUM of folded lines is exactly what lands in the cell (round6) */
const add = (x, y) => SpiCore.round6(x + y);
check('rounding: 1000 + -999.99 === 0.01', add(1000, -999.99) === 0.01);

if (failures) {
  console.log(`${failures} test(s) failed`);
  process.exit(1);
}
console.log('all source tests passed');
