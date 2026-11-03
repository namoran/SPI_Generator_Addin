/* Plan-equivalence test: the add-in must write exactly what the web
 * SPI Generator writes.
 *
 * For each real FDEP package in ../../handoff/packages/ (RFC + old SPI +
 * the SPI FDEP issued with the change order):
 *
 *   1. The web generator's own code (../gen/docs/logic.js) computes the
 *      reference plan.
 *   2. The add-in core (src/taskpane/spi-core.js) computes the same plan
 *      twice: once over the same file reader the web tool uses, and once
 *      with the RFC read through SheetJS — the exact parse path the task
 *      pane uses (src/taskpane/rfc-sheetjs.js). All three plans must be
 *      identical, and the report lines must match the web tool's.
 *   3. The web generator's writer (patch.js) applies the plan to the old
 *      SPI. An independent reader (openpyxl, tests/dump_cells.py) then
 *      confirms: every planned cell in the generated file holds the
 *      planned new value; the generated file's quantities match the
 *      FDEP-issued new SPI across the whole pay item table; and no
 *      planned cell in the old SPI holds a formula (the case the two
 *      writers handle by refusing to write).
 *
 * What this does not test: the Office.js calls in taskpane.js. There is
 * no Excel in this environment, so the in-Excel run (button, pane,
 * writes landing in a live workbook) is verified by sideloading — see
 * README.md.
 *
 * Run: npm test   (needs python3 with openpyxl, and npm install once)
 */
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..');
const genDocs = path.resolve(root, '..', 'gen', 'docs');
const genModules = path.resolve(root, '..', 'gen', 'node_modules');
const packagesDir = path.resolve(root, '..', 'handoff', 'packages');
const tmp = path.join(here, 'tmp');
fs.mkdirSync(tmp, { recursive: true });

/* --- the web generator, loaded the same way its own test loads it --- */
const sandbox = { console, TextDecoder, TextEncoder };
vm.createContext(sandbox);
const runIn = (file) => vm.runInContext(fs.readFileSync(file, 'utf8'), sandbox, { filename: file });
sandbox.fflate = require(path.join(genModules, 'fflate'));
runIn(path.join(genDocs, 'ppr-xlsx.js'));
runIn(path.join(genDocs, 'logic.js'));
runIn(path.join(genDocs, 'zip-repack.js'));
runIn(path.join(genDocs, 'patch.js'));
const WebSPI = sandbox.SPI;
const WebPatch = sandbox.SPIPatch;

/* --- the add-in core and its SheetJS RFC adapter --- */
const SpiCore = require(path.join(root, 'src', 'taskpane', 'spi-core.js'));
const SpiSheetJs = require(path.join(root, 'src', 'taskpane', 'rfc-sheetjs.js'));
const XLSX = require('xlsx');

const u8 = (p) => {
  const b = fs.readFileSync(p);
  return new Uint8Array(b.buffer, b.byteOffset, b.byteLength);
};

function pprAdapter(sheet) {
  return {
    value: (r, c) => sheet.value(r, c),
    range: (r1, r2, c1, c2) => sheet.range(r1, r2, c1, c2),
    formulaAt: () => null, // fixtures hold values, not formulas; asserted below via openpyxl
  };
}

function normPlan(changes) {
  return changes.map((c) => ({
    address: c.address, task: c.task, payItem: c.payItem,
    row: c.row, col: c.col, oldValue: c.oldValue,
    units: c.units, newValue: c.newValue, lines: c.lines, parts: c.parts,
  }));
}

function dump(mode, workbook, addresses) {
  const addrFile = path.join(tmp, 'addresses.json');
  fs.writeFileSync(addrFile, JSON.stringify(addresses));
  const res = spawnSync('python3', [path.join(here, 'dump_cells.py'), mode, workbook, addrFile],
    { encoding: 'utf8' });
  if (res.status !== 0) throw new Error('dump_cells.py failed: ' + (res.stderr || res.stdout));
  return JSON.parse(res.stdout);
}

const near = (a, b) => Math.abs((a ?? 0) - (b ?? 0)) < 1e-9;

let failures = 0;
const pkgs = fs.readdirSync(packagesDir).sort();
for (const pkg of pkgs) {
  const dir = path.join(packagesDir, pkg);
  const rfcPath = path.join(dir, 'rfc.xlsx');
  const oldPath = path.join(dir, 'spi_old.xlsx');
  const newPath = path.join(dir, 'spi_new.xlsx');
  try {
    const rfcWb = WebSPI.open(u8(rfcPath), 'the RFC file');
    const spiWb = WebSPI.open(u8(oldPath), 'the old SPI');
    const rfcSheet = WebSPI.sheetOrThrow(rfcWb, WebSPI.RFC_SHEET, 'the RFC file');
    const sowSheet = WebSPI.sheetOrThrow(spiWb, WebSPI.SOW_SHEET, 'the old SPI');

    const webPlan = WebSPI.plan(rfcSheet, sowSheet);
    const corePlan = SpiCore.plan(pprAdapter(rfcSheet), pprAdapter(sowSheet));

    const sheetJsWb = XLSX.readFile(rfcPath);
    if (!sheetJsWb.Sheets[SpiCore.RFC_SHEET]) throw new Error('SheetJS: no RFC sheet');
    const corePlanSJ = SpiCore.plan(
      SpiSheetJs.fromWorksheet(sheetJsWb.Sheets[SpiCore.RFC_SHEET], XLSX),
      pprAdapter(sowSheet));

    const a = JSON.stringify(normPlan(webPlan));
    if (a !== JSON.stringify(normPlan(corePlan))) throw new Error('add-in plan (file reader) differs from web plan');
    if (a !== JSON.stringify(normPlan(corePlanSJ))) throw new Error('add-in plan (SheetJS RFC) differs from web plan');
    if (JSON.stringify(WebSPI.report(webPlan)) !== JSON.stringify(SpiCore.report(corePlan))) {
      throw new Error('report lines differ from web tool');
    }

    /* Apply the web writer, then check the result with openpyxl. */
    const bytes = WebPatch.build(u8(oldPath), spiWb, WebSPI.SOW_SHEET, webPlan);
    const genPath = path.join(tmp, pkg + '-generated.xlsx');
    fs.writeFileSync(genPath, bytes);

    const layout = SpiCore.spiLayout(pprAdapter(sowSheet));
    const gridAddrs = [];
    for (const label of layout.labels) {
      for (const t of Object.keys(layout.taskCols)) {
        gridAddrs.push(SpiCore.colName(layout.taskCols[t]) + layout.rows[label]);
      }
    }
    const plannedAddrs = webPlan.map((c) => c.address);
    const genVals = dump('values', genPath, gridAddrs);
    const newVals = dump('values', newPath, gridAddrs);
    const oldFormulas = dump('formulas', oldPath, plannedAddrs);

    for (const c of webPlan) {
      if (!near(genVals[c.address], c.newValue)) {
        throw new Error(`generated ${c.address} = ${genVals[c.address]}, planned ${c.newValue}`);
      }
      if (oldFormulas[c.address]) {
        throw new Error(`old SPI ${c.address} holds a formula (${oldFormulas[c.address]}); fixture assumption broken`);
      }
    }
    for (const addr of gridAddrs) {
      if (!near(genVals[addr], newVals[addr])) {
        throw new Error(`generated ${addr} = ${genVals[addr]} but FDEP-issued SPI has ${newVals[addr]}`);
      }
    }
    console.log(`${pkg}: PASS — ${webPlan.length} cells planned identically by web tool and add-in; ` +
      `generated file matches the FDEP-issued SPI on all ${gridAddrs.length} table cells`);
  } catch (e) {
    failures++;
    console.log(`${pkg}: FAIL — ${e.message}`);
  }
}

if (failures) {
  console.log(`${failures} of ${pkgs.length} packages failed`);
  process.exit(1);
}
console.log(`all ${pkgs.length} packages passed`);
