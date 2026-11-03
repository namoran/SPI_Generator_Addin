/* SPI Generator (Excel Office Add-in) — core planning logic.
 *
 * Port of the web SPI Generator's docs/logic.js, which is itself a port of
 * SPI-Generator.py. It reads the RFC, works out how many units to add to
 * each task/pay-item cell of the SPI, and reports the changes. The web
 * version then edits the .xlsm file on disk; this version runs inside
 * Excel, so taskpane.js applies the plan to the open workbook through the
 * Excel JavaScript API and Excel itself recalculates the totals.
 *
 * Sheet adapter interface (1-based rows and columns):
 *   value(r, c)        -> the cell's value, or null when the cell is empty
 *   range(r1, r2, c1, c2) -> 2D array of values for rows r1..r2, cols
 *                            c1..c2 (may be clamped at the sheet's edge)
 * The SPI adapter may also provide:
 *   formulaAt(r, c)    -> the cell's formula string, or null when the cell
 *                         holds no formula. The web generator refuses to
 *                         write into a formula cell (its patch.js throws),
 *                         and plan() does the same when this is provided.
 *
 * Pay items are located from the labels printed in the SPI workbook
 * itself -- there is no built-in pay item list.
 */
var SpiCore = (function () {
  'use strict';

  var RFC_SHEET = 'RFC';
  var SOW_SHEET = 'SOW Units';
  var RFC_FIRST_ROW = 22;   // Python: min_row=22 (skips the surcharge rows)
  var RFC_LAST_ROW = 111;   // Python: max_row=111

  var LABEL_RE = /^\d{1,2}([-.][0-9A-Za-z]+)*\.$/;

  /* Find the 'PAY ITEM' header cell in a scanned block of the SOW Units
   * sheet. The block is a 2D array starting at row 1, column 1.
   * Returns { row, col } (1-based). Throws the web generator's error
   * when there is no such cell. */
  function findHeader(scan) {
    for (var i = 0; i < scan.length; i++) {
      for (var j = 0; j < scan[i].length; j++) {
        var v = scan[i][j];
        if (typeof v === 'string' && v.trim().toLowerCase() === 'pay item') {
          return { row: i + 1, col: j + 1 };
        }
      }
    }
    throw new Error('no pay item table found in the SPI workbook: ' +
                    'no "PAY ITEM" header cell in the SOW Units sheet');
  }

  /* Task quantity columns, from the header rows around the PAY ITEM
   * header. valueAt(r, c) reads the sheet (1-based). Returns
   * { taskNumber: column }. */
  function findTaskCols(valueAt, headerRow) {
    var taskCols = {};
    for (var r = Math.max(1, headerRow - 1); r <= headerRow + 1; r++) {
      for (var c = 1; c <= 24; c++) {
        var hv = valueAt(r, c);
        if (typeof hv === 'string') {
          var m = /TASK\s*(\d+)/i.exec(hv);
          if (m) taskCols[parseInt(m[1], 10)] = c;
        }
      }
    }
    if (!Object.keys(taskCols).length) {
      throw new Error('no task quantity columns found in the SPI workbook');
    }
    return taskCols;
  }

  /* Pay item labels down the PAY ITEM column below the header, ending
   * at the totals/version block. colValues is the column's values
   * starting at headerRow + 1. Returns { labels, rows } with rows
   * mapping each label to its 1-based sheet row. */
  function parseLabels(colValues, headerRow) {
    var labels = [], rows = {}, started = false, misses = 0;
    for (var k = 0; k < colValues.length; k++) {
      var raw = colValues[k];
      var s = raw === null || raw === undefined ? '' : String(raw).trim();
      if (LABEL_RE.test(s)) {
        started = true; misses = 0;
        if (!(s in rows)) { rows[s] = headerRow + 1 + k; labels.push(s); }
      } else if (started) {
        if (++misses >= 10) break;
      }
    }
    if (labels.length < 10) {
      throw new Error('only ' + labels.length + ' pay items found in the SPI workbook');
    }
    return { labels: labels, rows: rows };
  }

  /* Locate the pay item table in an SOW Units sheet from the workbook
   * itself. The header row carries a 'PAY ITEM' cell and the 'TASK n
   * QUANT.' column headers; pay item labels run down the PAY ITEM
   * column below the header, ending at the totals/version block.
   * Throws when the table cannot be found -- positions are never
   * guessed, because a workbook from a different SPI version has
   * different rows (Version 10.0, for example, has no section 22, so
   * section 23 sits 31 rows above where a fixed list would look).
   * Mirrors spi_layout in the Python version. */
  function spiLayout(sheet) {
    var header = findHeader(sheet.range(1, 40, 1, 24));
    var taskCols = findTaskCols(function (r, c) { return sheet.value(r, c); }, header.row);
    var colVals = sheet.range(header.row + 1, header.row + 800, header.col, header.col);
    var flat = colVals.map(function (row) { return row[0]; });
    var parsed = parseLabels(flat, header.row);
    return { labels: parsed.labels, rows: parsed.rows, taskCols: taskCols };
  }

  /* The cell for a pay item and task, located from the workbook's own
   * pay item table. Throws naming the item when the workbook has no row
   * with that label. */
  function cellFor(pi, task, layout) {
    var row = layout.rows[pi];
    if (row === undefined) {
      throw new Error('Pay item "' + pi + '" is not in the SPI workbook: ' +
        'its pay item table has no row with that label.');
    }
    var tnum = parseInt(String(task).split(' ').pop(), 10);
    var col = layout.taskCols[tnum];
    if (col === undefined) {
      throw new Error('Task "' + task + '" has no quantity column in the SPI workbook.');
    }
    return { row: row, col: col };
  }

  /* Python load_rfc: rows 22..111, columns A..I, rows with no task skipped. */
  function loadRfc(sheet) {
    var m = sheet.range(RFC_FIRST_ROW, RFC_LAST_ROW, 1, 9);
    var out = [];
    for (var i = 0; i < m.length; i++) {
      if (m[i][0] === null || m[i][0] === undefined) continue;
      out.push({
        task: 'TASK ' + String(m[i][0]),
        payItem: String(m[i][1]),
        uom: m[i][5],
        units: m[i][7],
        total: m[i][8],
        sourceRow: i + RFC_FIRST_ROW
      });
    }
    return out;
  }

  /* Work out every cell to touch, with its current and new value.
   * An RFC can list the same pay item more than once in a task (e.g. a
   * negative line removing a placeholder and a positive line entering the
   * actual amount), so lines for the same cell are folded together in RFC
   * order -- the same accumulation the Python version performs when it
   * adds each line into the live cell. One planned change per cell. */
  function plan(rfcSheet, sowSheet) {
    var items = loadRfc(rfcSheet);
    var layout = spiLayout(sowSheet);
    var byCell = {};
    var changes = [];
    items.forEach(function (it) {
      var at = cellFor(it.payItem, it.task, layout);
      var units = it.units === null || it.units === undefined ? 0 : it.units;
      if (typeof units !== 'number') {
        throw new Error('Row ' + it.sourceRow + ' of the RFC has a non-numeric quantity for pay item ' +
          it.payItem + ' (' + it.task + ').');
      }
      var key = at.row + '|' + at.col;
      var ch = byCell[key];
      if (!ch) {
        /* The web generator's writer (patch.js) refuses to overwrite a
         * formula cell; check before planning any write to one. */
        if (typeof sowSheet.formulaAt === 'function' && sowSheet.formulaAt(at.row, at.col)) {
          throw new Error('Cell ' + colName(at.col) + at.row + ' (' + it.task + ', ' + it.payItem +
            ') holds a formula, so adding units there would destroy it. ' +
            'Unhide the pay item row in the SPI and try again.');
        }
        var current = sowSheet.value(at.row, at.col);
        var oldVal = current === null || current === undefined ? 0 : current;
        if (typeof oldVal !== 'number') {
          throw new Error('Row ' + it.sourceRow + ' of the RFC has a non-numeric quantity for pay item ' +
            it.payItem + ' (' + it.task + ').');
        }
        ch = {
          task: it.task,
          payItem: it.payItem,
          row: at.row,
          col: at.col,
          address: colName(at.col) + at.row,
          oldValue: oldVal,
          units: 0,
          newValue: oldVal,
          lines: 0,
          parts: []
        };
        byCell[key] = ch;
        changes.push(ch);
      }
      ch.units += units;
      ch.newValue += units;
      ch.lines += 1;
      ch.parts.push(units);
    });
    /* Quantities in this program are decimal values -- dollar amounts to
     * the cent, and FDEP-issued quantities carry at most 6 decimal places.
     * A raw binary-float sum can land just off the intended decimal
     * (1000 + -999.99 = 0.009999999999990905 instead of 0.01), so each
     * planned value is rounded to 6 decimals before it is written. */
    changes.forEach(function (ch) {
      ch.units = round6(ch.units);
      ch.newValue = round6(ch.newValue);
    });
    return changes;
  }

  function round6(v) {
    var r = Math.round(v * 1e6) / 1e6;
    return r === 0 ? 0 : r;
  }

  function colName(n) {
    var s = '';
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
    return s;
  }

  function report(changes) {
    var out = [];
    changes.forEach(function (c) {
      out.push(c.units + ' unit(s) added to ' + c.oldValue + ' in ' + c.task + ', ' + c.payItem +
               '. The new value is ' + c.newValue + '.');
      if (c.lines > 1) {
        out.push('Note: this combines ' + c.lines + ' RFC lines (' + c.parts.join(', ') +
                 ') for ' + c.task + ', ' + c.payItem + '.');
      }
    });
    return out;
  }

  /* The web generator's missing-sheet error, worded for this add-in's
   * two inputs: the picked RFC file and the open workbook. */
  function missingSheetError(label, name, names) {
    return new Error(label + ' has no "' + name + '" sheet. Sheets found: ' + names.join(', '));
  }

  return {
    loadRfc: loadRfc,
    plan: plan,
    report: report,
    spiLayout: spiLayout,
    findHeader: findHeader,
    findTaskCols: findTaskCols,
    parseLabels: parseLabels,
    cellFor: cellFor,
    round6: round6,
    colName: colName,
    missingSheetError: missingSheetError,
    RFC_SHEET: RFC_SHEET,
    SOW_SHEET: SOW_SHEET
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = SpiCore;
if (typeof window !== 'undefined') window.SpiCore = SpiCore;
