/* RFC Importer (Excel Office Add-in) — task pane wiring.
 *
 * The user opens their SPI workbook in Excel and picks the source of the
 * RFC data in this pane: an RFC file, or the RFC worksheet inside the
 * same workbook. The RFC lines are read through an adapter (SheetJS for
 * a file; a values block read with the Excel API for the worksheet).
 * The plan — which cells change, and to what — comes from spi-core.js,
 * the same logic the web SPI Generator uses. The plan is then applied to
 * the open workbook with the Excel JavaScript API: each planned cell
 * gets its new quantity, and each row that received a value is
 * unhidden. Nothing else in the workbook is touched, and Excel
 * recalculates the totals.
 *
 * The plan is computed in full before anything is written, so an error
 * (a pay item missing from the SPI, a formula in a quantity cell, a
 * non-numeric quantity) leaves the workbook unchanged.
 *
 * Imports are idempotent: before any write, a fingerprint of the RFC's
 * tasks, pay items, and quantities (rfc-fingerprint.js) is compared
 * against the record of imports stored in the workbook's document
 * settings. Matching data is refused — the import adds quantities, so
 * applying the same data twice would double those cells — with an
 * "Apply anyway" override for an identical change order re-imported on
 * purpose.
 */
(function () {
  'use strict';

  var SETTINGS_KEY = 'RfcImporter.appliedRfcs';

  var rfcAdapter = null;   // the picked RFC file, adapted
  var pendingApply = null; // an RFC blocked as already imported, waiting on "Apply anyway"

  var fileInput = document.getElementById('rfc-file');
  var rfcInfo = document.getElementById('rfc-info');
  var runBtn = document.getElementById('generate');
  var overrideBtn = document.getElementById('override');
  var statusEl = document.getElementById('status');
  var resultEl = document.getElementById('result');
  var summaryEl = document.getElementById('summary');
  var reportEl = document.getElementById('report');
  var fileField = document.getElementById('file-field');
  var sourceRadios = document.getElementsByName('rfc-source');

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

  function selectedSource() {
    for (var i = 0; i < sourceRadios.length; i++) {
      if (sourceRadios[i].checked) return sourceRadios[i].value;
    }
    return 'file';
  }

  for (var sr = 0; sr < sourceRadios.length; sr++) {
    sourceRadios[sr].addEventListener('change', resetForSource);
  }

  function resetForSource() {
    var sheet = selectedSource() === 'sheet';
    fileField.classList.toggle('hidden', sheet);
    rfcAdapter = null;
    pendingApply = null;
    overrideBtn.classList.add('hidden');
    runBtn.disabled = sheet ? false : true; // the worksheet needs no pick to be ready
    rfcInfo.classList.add('hidden');
    rfcInfo.textContent = '';
    resultEl.classList.add('hidden');
    clearStatus();
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

  /* --- the RFC worksheet inside the open workbook --- */

  /* A read-only sheet adapter over rows 22..111, columns A..I of the
   * "RFC" worksheet in the open workbook, read with the Excel API in
   * one sync. Values are read live at import time, so lines typed or
   * pasted into the sheet — saved or not — are what gets imported. */
  function buildSheetRfc(context) {
    var sheets = context.workbook.worksheets;
    sheets.load('items/name');
    return context.sync().then(function () {
      var names = sheets.items.map(function (s) { return s.name; });
      if (names.indexOf(SpiCore.RFC_SHEET) === -1) {
        throw SpiCore.missingSheetError('This workbook', SpiCore.RFC_SHEET, names);
      }
      var rows = SpiCore.RFC_LAST_ROW - SpiCore.RFC_FIRST_ROW + 1;
      var rng = sheets.getItem(SpiCore.RFC_SHEET)
        .getRangeByIndexes(SpiCore.RFC_FIRST_ROW - 1, 0, rows, 9);
      rng.load('values');
      return context.sync().then(function () {
        return RfcGrid.fromMatrix(rng.values, SpiCore.RFC_FIRST_ROW);
      });
    });
  }

  /* The RFC ready to plan: { rfc, items, source, info }. */
  function prepareRfc() {
    if (selectedSource() === 'sheet') {
      return Excel.run(function (context) {
        return buildSheetRfc(context).then(function (adapter) {
          var items = SpiCore.loadRfc(adapter);
          if (!items.length) {
            throw new Error('No RFC line items on the "RFC" worksheet in this ' +
              'workbook (rows 22 to 111). Fill in the line items in that sheet, ' +
              'or import an RFC file instead.');
          }
          return { rfc: adapter, items: items, source: 'RFC worksheet in this workbook',
            info: 'RFC worksheet in this workbook — ' + items.length + ' RFC line' +
              (items.length === 1 ? '' : 's') + ' found.' };
        });
      });
    }
    if (!rfcAdapter) {
      return Promise.reject(new Error('Select an RFC file first.'));
    }
    var items = SpiCore.loadRfc(rfcAdapter);
    return Promise.resolve({ rfc: rfcAdapter, items: items, source: 'RFC file',
      info: rfcInfo.textContent });
  }

  /* --- the record of imports (document settings) --- */

  function emptyRecord() {
    return { version: 1, applied: [] };
  }

  /* Set when this host cannot store the import record, with the reason.
   * Imports still work, but duplicate-import detection is off, and the
   * pane says so. */
  var recordStoreNote = null;

  function settingsRead() {
    return new Promise(function (resolve) {
      var settings = null;
      try {
        settings = Office.context.document.settings;
      } catch (e) {
        settings = null;
      }
      if (!settings || typeof settings.getAsync !== 'function') {
        recordStoreNote = 'this host has no document settings to store the import record';
        resolve(emptyRecord());
        return;
      }
      settings.getAsync(function (ar) {
        var record = null;
        var failed = !ar || ar.status !== Office.AsyncResultStatus.Succeeded;
        if (!failed && ar.value) {
          try { record = JSON.parse(ar.value); } catch (e) { record = null; }
        }
        if (failed && ar && ar.error && ar.error.message) {
          recordStoreNote = 'the import record could not be read: ' + ar.error.message;
        }
        if (!record || typeof record !== 'object' || !Array.isArray(record.applied)) {
          record = emptyRecord();
        }
        resolve(record);
      });
    });
  }

  function settingsWrite(record) {
    return new Promise(function (resolve, reject) {
      var settings = null;
      try {
        settings = Office.context.document.settings;
      } catch (e) {
        settings = null;
      }
      if (!settings || typeof settings.setAsync !== 'function') {
        reject(new Error(recordStoreNote || 'this host has no document settings to store the import record'));
        return;
      }
      settings.setAsync(JSON.stringify(record), function (ar) {
        if (!ar || ar.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error('the import record could not be saved to this workbook: ' +
            (ar && ar.error && ar.error.message ? ar.error.message : 'unknown error')));
          return;
        }
        settings.saveAsync(function (ar2) {
          if (!ar2 || ar2.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error('the import record could not be saved to this workbook: ' +
              (ar2 && ar2.error && ar2.error.message ? ar2.error.message : 'unknown error')));
            return;
          }
          resolve();
        });
      });
    });
  }

  function priorDates(record, fp) {
    return ((record && record.applied) || []).filter(function (e) { return e.hash === fp; })
      .map(function (e) { return e.date; });
  }

  function isoDateOnly(iso) {
    return String(iso).slice(0, 10);
  }

  /* --- blocking (already imported) and applying --- */

  function showBlocked(dates) {
    showStatus('This RFC appears to have identical tasks, pay items, and ' +
      'quantities to a previously imported RFC' +
      ' (imported ' + dates.map(isoDateOnly).join(', ') + '). ' +
      'Nothing was changed — the matching task quantity cells in this workbook ' +
      'already reflect these quantities.', true);
    overrideBtn.classList.remove('hidden');
  }

  function applyRfc(prep, fp, record) {
    showStatus('Importing…', false);
    return generate(prep.rfc).then(function (changes) {
      resultEl.classList.remove('hidden');
      summaryEl.textContent = changes.length + ' cell' + (changes.length === 1 ? '' : 's') +
        ' updated in the "SOW Units" sheet.';
      reportEl.innerHTML = '';
      SpiCore.report(changes).forEach(function (line) {
        var li = document.createElement('li');
        li.textContent = line;
        reportEl.appendChild(li);
      });

      clearStatus();
      rfcInfo.textContent = prep.info;
      rfcInfo.classList.remove('hidden');

      record.applied.push({
        hash: fp,
        date: new Date().toISOString(),
        source: prep.source,
        lines: prep.items.length
      });
      settingsWrite(record).catch(function (err) {
        showStatus('Imported ' + changes.length + ' cell(s), but the import record ' +
          'could not be saved to this workbook (' + err.message + '). ' +
          'Re-importing the same RFC would not be detected.', true);
      });
    });
  }

  /* --- the Excel wiring (unchanged from the file-only version) --- */

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

  function generate(rfcSheet) {
    return Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.load('items/name');
      return context.sync().then(function () {
        var names = sheets.items.map(function (s) { return s.name; });
        if (names.indexOf(SpiCore.SOW_SHEET) === -1) {
          throw SpiCore.missingSheetError('The open workbook', SpiCore.SOW_SHEET, names);
        }
        return applyPlan(context, sheets.getItem(SpiCore.SOW_SHEET), rfcSheet);
      });
    });
  }

  function applyPlan(context, sheet, rfcSheet) {
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
          var changes = SpiCore.plan(rfcSheet, adapter);

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
    runBtn.disabled = true;
    resultEl.classList.add('hidden');
    overrideBtn.classList.add('hidden');
    pendingApply = null;
    showStatus('Checking…', false);

    Promise.all([settingsRead(), prepareRfc().catch(function (err) {
      return Promise.reject(err);
    })]).then(function (results) {
      var record = results[0];
      var prep = results[1];
      var fp = RfcFingerprint.fingerprint(prep.items);
      if (recordStoreNote) {
        showStatus('Note: duplicate-import detection is off — ' + recordStoreNote +
          '. The import still runs; check the quantities before relying on them.', false);
      }
      var dates = priorDates(record, fp);
      if (dates.length) {
        pendingApply = prep;
        showBlocked(dates);
        runBtn.disabled = false;
        return;
      }
      return applyRfc(prep, fp, record).then(function () {
        runBtn.disabled = false;
      });
    }).catch(function (err) {
      showStatus(err && err.message ? err.message : String(err), true);
      runBtn.disabled = false;
    });
  });

  overrideBtn.addEventListener('click', function () {
    if (!pendingApply) return;
    var prep = pendingApply;
    pendingApply = null;
    overrideBtn.classList.add('hidden');
    runBtn.disabled = true;
    showStatus('Checking…', false);
    Promise.all([settingsRead(), Promise.resolve(prep)]).then(function (results) {
      var record = results[0];
      var fp = RfcFingerprint.fingerprint(prep.items);
      return applyRfc(prep, fp, record);
    }).then(function () {
      runBtn.disabled = false;
    }).catch(function (err) {
      showStatus(err && err.message ? err.message : String(err), true);
      runBtn.disabled = false;
    });
  });
})();
