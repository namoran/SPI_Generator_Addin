/* SPI Generator (Excel Office Add-in) — RFC sheet adapter.
 *
 * Adapts a SheetJS worksheet to the sheet interface spi-core.js uses
 * (1-based value / range reads). SheetJS stores the last calculated
 * value of a formula cell in the cell's `v` when it reads a file, so the
 * values read here are the workbook's cached values — the same values
 * the web SPI Generator and the Python version (openpyxl, data_only)
 * read from an RFC.
 *
 * Used by taskpane.js for the RFC file the user picks, and by the Node
 * test in tests/ for the same code path.
 */
(function () {
  'use strict';

  function fromWorksheet(ws, XLSX) {
    function value(r, c) {
      var cell = ws[XLSX.utils.encode_cell({ r: r - 1, c: c - 1 })];
      if (!cell || cell.v === undefined || cell.v === null) return null;
      return cell.v;
    }
    function range(r1, r2, c1, c2) {
      var out = [];
      for (var r = r1; r <= r2; r++) {
        var row = [];
        for (var c = c1; c <= c2; c++) row.push(value(r, c));
        out.push(row);
      }
      return out;
    }
    return { value: value, range: range };
  }

  var api = { fromWorksheet: fromWorksheet };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof window !== 'undefined') window.SpiSheetJs = api;
})();
