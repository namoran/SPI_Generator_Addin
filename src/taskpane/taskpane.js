/* SPI Generator (Excel Office Add-in) — task pane wiring.
 *
 * The user opens their SPI workbook in Excel and picks the RFC file here.
 * The RFC is parsed in the pane with SheetJS (rfc-sheetjs.js). The plan —
 * which cells change, and to what — comes from spi-core.js, the same
 * logic the web SPI Generator uses. The plan is then applied to the open
 * workbook with the Excel JavaScript API: each planned cell gets its new
 * quantity, and each row that received a value is unhidden. Nothing else
 * in the workbook is touched, and Excel recalculates the totals.
 *
 * The plan is computed in full before anything is written, so an error
 * (a pay item missing from the SPI, a formula in a quantity cell, a
 * non-numeric quantity) leaves the workbook unchanged.
 */
(function () {
  'use strict';

  var rfcAdapter = null;

  var fileInput = document.getElementById('rfc-file');
  var rfcInfo = document.getElementById('rfc-info');
  var runBtn = document.getElementById('generate');
  var statusEl = document.getElementById('status');
  var resultEl = document.getElementById('result');
  var summaryEl = document.getElementById('summary');
  var reportEl = document.getElementById('report');

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) {
      showStatus('This add-in runs in Microsoft Excel.', true);
    }
  });

  function showStatus(text, isError) {
    statusEl.textContent = text;
    statusEl.className = isError ? 'error' : '';
    statusEl.classList.remove('hidden');
  }

  function clearStatus() {
    statusEl.classList.add('hidden');
    statusEl.className = '';
  }

  function readBuffer(file) {
    if (file.arrayBuffer) return file.arrayBuffer();
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload = function () { resolve(r.result); };
      r.onerror = function () { reject(new Error('could not read ' + file.name)); };
      r.readAsArrayBuffer(file);
    });
  }

  fileInput.addEventListener('change', function () {
    var file = fileInput.files && fileInput.files[0];
    rfcAdapter = null;
    runBtn.disabled = true;
    rfcInfo.classList.add('hidden');
    resultEl.classList.add('hidden');
    if (!file) return;
    clearStatus();
    readBuffer(file).then(function (buf) {
      var wb;
      try {
        wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
      } catch (e) {
        throw new Error('Could not read the RFC file: ' + e.message);
      }
      if (wb.SheetNames.indexOf(SpiCore.RFC_SHEET) === -1) {
        throw SpiCore.missingSheetError('The RFC file', SpiCore.RFC_SHEET, wb.SheetNames);
      }
      var adapter = SpiSheetJs.fromWorksheet(wb.Sheets[SpiCore.RFC_SHEET], XLSX);
      var items = SpiCore.loadRfc(adapter);
      if (!items.length) {
        throw new Error('No RFC line items found in ' + file.name +
          ' (sheet "RFC", rows 22 to 111).');
      }
      rfcAdapter = adapter;
      rfcInfo.textContent = file.name + ' — ' + items.length + ' RFC line' +
        (items.length === 1 ? '' : 's') + ' found.';
      rfcInfo.classList.remove('hidden');
      runBtn.disabled = false;
    }).catch(function (err) {
      showStatus(err && err.message ? err.message : String(err), true);
    });
  });

  /* A cache of SOW Units cell values and formulas, filled from Excel in
   * blocks, behind the synchronous adapter spi-core.js plans against.
   * Office.js returns "" for an empty cell; the core logic follows the
   * file-based tools, where an empty cell is null, so "" becomes null. */
  function GridCache() {
    this.values = {};
    this.formulas = {};
  }
  GridCache.prototype.adapter = function () {
    var self = this;
    function value(r, c) {
      var v = self.values[r + '|' + c];
      return v === undefined ? null : v;
    }
    return {
      value: value,
      range: function (r1, r2, c1, c2) {
        var out = [];
        for (var r = r1; r <= r2; r++) {
          var row = [];
          for (var c = c1; c <= c2; c++) row.push(value(r, c));
          out.push(row);
        }
        return out;
      },
      formulaAt: function (r, c) {
        var f = self.formulas[r + '|' + c];
        return (typeof f === 'string' && f.charAt(0) === '=') ? f : null;
      }
    };
  };

  function loadBlock(context, sheet, cache, r1, r2, c1, c2) {
    var CHUNK = 200;
    var jobs = [];
    for (var start = r1; start <= r2; start += CHUNK) {
      jobs.push({ r: start, rows: Math.min(CHUNK, r2 - start + 1) });
    }
    return jobs.reduce(function (p, job) {
      return p.then(function () {
        var rng = sheet.getRangeByIndexes(job.r - 1, c1 - 1, job.rows, c2 - c1 + 1);
        rng.load('values,formulas');
        return context.sync().then(function () {
          for (var i = 0; i < job.rows; i++) {
            for (var j = 0; j < c2 - c1 + 1; j++) {
              var key = (job.r + i) + '|' + (c1 + j);
              var v = rng.values[i][j];
              cache.values[key] = v === '' ? null : v;
              cache.formulas[key] = rng.formulas[i][j];
            }
          }
        });
      });
    }, Promise.resolve());
  }

  function generate() {
    return Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.load('items/name');
      return context.sync().then(function () {
        var names = sheets.items.map(function (s) { return s.name; });
        if (names.indexOf(SpiCore.SOW_SHEET) === -1) {
          throw SpiCore.missingSheetError('The open workbook', SpiCore.SOW_SHEET, names);
        }
        return applyPlan(context, sheets.getItem(SpiCore.SOW_SHEET));
      });
    });
  }

  function applyPlan(context, sheet) {
    var used = sheet.getUsedRangeOrNullObject();
    used.load('rowCount,columnCount');
    return context.sync().then(function () {
      if (used.isNullObject || !used.rowCount) {
        throw new Error('The "SOW Units" sheet in the open workbook is empty.');
      }
      var rowCount = used.rowCount;
      var colCount = Math.max(used.columnCount, 24);
      var cache = new GridCache();
      var adapter = cache.adapter();

      /* Header block first, so the label column's extent is known; then
       * the label column, so the rows the plan can touch are known; then
       * those rows in full. spi-core.js re-derives the same layout from
       * the cache when it plans, so the pane and the plan share one
       * layout code path. */
      return loadBlock(context, sheet, cache, 1, Math.min(41, rowCount), 1, 24)
        .then(function () {
          var header = SpiCore.findHeader(adapter.range(1, 40, 1, 24));
          var labelEnd = Math.min(rowCount, header.row + 800);
          return loadBlock(context, sheet, cache, header.row + 1, labelEnd, header.col, header.col)
            .then(function () { return header; });
        })
        .then(function (header) {
          var colVals = adapter.range(header.row + 1, header.row + 800, header.col, header.col)
            .map(function (row) { return row[0]; });
          var parsed = SpiCore.parseLabels(colVals, header.row);
          var maxRow = parsed.labels.reduce(function (m, label) {
            return Math.max(m, parsed.rows[label]);
          }, header.row);
          return loadBlock(context, sheet, cache, header.row + 1, maxRow, 1, 24);
        })
        .then(function () {
          var changes = SpiCore.plan(rfcAdapter, adapter);

          var rowsToUnhide = {};
          changes.forEach(function (ch) {
            sheet.getRangeByIndexes(ch.row - 1, ch.col - 1, 1, 1).values = [[ch.newValue]];
            rowsToUnhide[ch.row] = true;
          });
          Object.keys(rowsToUnhide).forEach(function (r) {
            sheet.getRangeByIndexes(Number(r) - 1, 0, 1, colCount).rowHidden = false;
          });
          return context.sync().then(function () { return changes; });
        });
    });
  }

  runBtn.addEventListener('click', function () {
    if (!rfcAdapter) return;
    runBtn.disabled = true;
    resultEl.classList.add('hidden');
    showStatus('Processing…', false);
    generate().then(function (changes) {
      clearStatus();
      summaryEl.textContent = changes.length + ' cell' + (changes.length === 1 ? '' : 's') +
        ' updated in the "SOW Units" sheet.';
      reportEl.innerHTML = '';
      SpiCore.report(changes).forEach(function (line) {
        var li = document.createElement('li');
        li.textContent = line;
        reportEl.appendChild(li);
      });
      resultEl.classList.remove('hidden');
    }).catch(function (err) {
      showStatus(err && err.message ? err.message : String(err), true);
    }).then(function () {
      runBtn.disabled = false;
    });
  });
})();
