/* RFC Importer (Excel Office Add-in) — matrix sheet adapter.
 *
 * Adapts a block of values read from a live workbook to the sheet
 * interface spi-core.js plans against (1-based value / range reads).
 * Office.js returns "" for an empty cell; the core logic follows the
 * file-based tools, where an empty cell is null, so "" becomes null.
 *
 * Used by taskpane.js for the RFC worksheet in the open workbook, and
 * by the Node test in tests/ for the same code path.
 */
(function () {
  /* matrix: 2D array of cell values; firstRow: sheet row of matrix[0] (1-based;
   * the columns start at sheet column 1). Requests outside the matrix come
   * back as null, like an empty cell. */
  function fromMatrix(matrix, firstRow) {
    function value(r, c) {
      var i = r - firstRow;
      var j = c - 1;
      if (i < 0 || j < 0 || !matrix || i >= matrix.length) return null;
      var row = matrix[i];
      if (!row || j >= row.length) return null;
      var v = row[j];
      if (v === '' || v === undefined || v === null) return null;
      return v;
    }
    function range(r1, r2, c1, c2) {
      var out = [];
      if (r1 > r2) return out;
      for (var r = r1; r <= r2; r++) {
        var row = [];
        for (var c = c1; c <= c2; c++) row.push(value(r, c));
        out.push(row);
      }
      return out;
    }
    return { value: value, range: range };
  }

  var api = { fromMatrix: fromMatrix };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof window !== 'undefined') window.RfcGrid = api;
})();
