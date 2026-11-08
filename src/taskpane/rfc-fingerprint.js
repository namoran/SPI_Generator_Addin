/* RFC Importer (Excel Office Add-in) — RFC fingerprint.
 *
 * Computes an identity for an RFC's line items, used to make imports
 * idempotent: an RFC whose tasks, pay items, and quantities were already
 * imported into this workbook is refused (with an "apply anyway"
 * override) instead of being added a second time, because the import
 * adds quantities to the task cells — re-importing the same data would
 * double every affected cell.
 *
 * The identity is order-independent (lines sorted before hashing), so
 * the same data re-pasted into an RFC sheet in a different row order
 * still matches. Strings are trimmed, so {' LF '} matches 'LF'; numbers
 * are rounded the same way spi-core.js rounds a planned value, so
 * 1000 + -999.99 === 0.01 does not change the identity.
 *
 * The hash is pair of 32-bit FNV-1a hashes with different seeds. Kept
 * off crypto.subtle (async) and off BigInt so the code runs in any
 * webview the add-in can load in, and is plain node-testable.
 */
(function () {
  'use strict';

  /* Math.imul is absent in old webviews; this is the standard polyfill. */
  function imul(a, b) {
    var ah = (a >>> 16) & 0xffff;
    var al = a & 0xffff;
    var bh = (b >>> 16) & 0xffff;
    var bl = b & 0xffff;
    return ((al * bl) + ((((ah * bl) + (al * bh)) << 16) >>> 0)) | 0;
  }

  function fnv1a32(str, seed) {
    var h = seed | 0;
    for (var i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = imul(h, 0x01000193);
    }
    return h >>> 0;
  }

  function round6(v) {
    var r = Math.round(v * 1e6) / 1e6;
    return r === 0 ? 0 : r;
  }

  function textOf(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/\u00a0/g, ' ').trim();
  }

  function numOrText(v) {
    if (typeof v === 'number' && isFinite(v)) return String(round6(v));
    return textOf(v);
  }

  /* The canonical, order-independent text for a list of loadRfc items.
   * One line per RFC line: length-prefixed, \u0001-joined fields. */
  function canonicalLines(items) {
    return (items || []).map(function (it) {
      var s = textOf(it.task) + '\u0001' + textOf(it.payItem) + '\u0001' +
        numOrText(it.uom) + '\u0001' + numOrText(it.units) + '\u0001' +
        numOrText(it.total);
      return s.length + ':' + s;
    }).sort();
  }

  function hash32(joined, seed) {
    return fnv1a32(joined, seed).toString(16);
  }

  function fingerprint(items) {
    var joined = canonicalLines(items).join('\n');
    var a = hash32(joined, 0x811c9dc5); /* FNV offset basis */
    var b = hash32(joined, 0x1b873593);
    while (a.length < 8) a = '0' + a;
    while (b.length < 8) b = '0' + b;
    return a + b;
  }

  var api = { fingerprint: fingerprint, canonicalLines: canonicalLines };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof window !== 'undefined') window.RfcFingerprint = api;
})();
