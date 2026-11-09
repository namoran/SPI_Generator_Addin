# RFC Importer — Excel Add-in

The RFC Importer as a Microsoft Excel add-in (Office Add-ins platform:
a task pane web page plus a manifest, driven by the Excel JavaScript
API). It does the same job as the web SPI Generator at
https://namoran.github.io/SPI_Generator/ : it transfers the line items
from a PRP change order (RFC) into an SPI workbook. (Until version
1.1.0.0 the add-in was listed as "SPI Generator"; the same job, a name
that describes the action.)

The difference is where the work happens. The web tool edits a copy of
the SPI file and hands back a download. The add-in works on the SPI
workbook you already have open in Excel:

1. Open the SPI workbook in Excel.
2. Open the RFC Importer pane (Home tab) and choose the RFC data
   source: an RFC file, or the RFC worksheet inside the open workbook
   itself.
3. Select Import. The RFC quantities are added into the open workbook's
   task quantity cells, the rows that received a value are unhidden,
   and the pane lists every change. Excel recalculates the totals.
   Save the workbook to keep the changes.

Only the task quantity cells change — the same cells the web generator
changes. Both sources are the RFC's rows 22–111, columns A–I: the same
data layout, so a picked file and a filled-in worksheet import
identically. The pay item table is located from the labels printed in
the workbook itself, so SPI versions with different layouts (Version
10.0, for example) work the same way they do in the web tool. If a
quantity cell holds a formula, or a pay item from the RFC is not in the
workbook, the add-in stops with an error and writes nothing.

The import is idempotent. Before writing anything, the add-in compares
a fingerprint of the RFC's tasks, pay items, and quantities
(`src/taskpane/rfc-fingerprint.js`) against the record of imports
saved inside the workbook itself. Identical data is refused — with
the message and dates of every previous import — because importing the
same RFC twice would double the quantities it carries. "Apply anyway"
overrides the refusal deliberately (a genuinely identical change order
re-imported on purpose). The record lives in the workbook, so saving
persists the quantities and the record together; a workbook closed
without saving keeps neither. The fingerprint is order- and
whitespace-insensitive, so re-entering the same lines in a different
order still matches.

An optional **Import note** (the "Import note (optional)" field in the
pane) is saved in each import's record entry, alongside the fingerprint,
date, source, line count, and a complete change log of every cell the
import touched. If the same RFC is imported again, the refusal message
quotes the notes saved on those previous imports, so the audit trail
carries context (e.g. "RFC #2 approved by FDEP").

The record is stored in the workbook's **custom document properties**,
through the Excel JavaScript API (`ExcelApi 1.7`) — the same API the
add-in uses to write the SOW Units cells. The properties are invisible,
travel with the file, and are part of the OOXML spec; the record is
chunked across them because a property's value is capped at 255
characters outside Excel on the web. A record in the older
`Office.context.document.settings` store (as earlier versions wrote) is
still read and merged in, so nothing saved by an older build is lost.
If a host supports neither store, the pane keeps the same guard in its
own memory for the pane's lifetime — re-importing the same RFC is still
refused until the pane is closed — and says so in a note after the
import. Host diagnostics appear as `[RFC Importer]` lines in the
console.

Nothing is uploaded. The RFC is read in the pane (or in the open
workbook itself) on your computer, and the add-in collects no data.
See `src/legal/privacy.html`.

## Layout

| path | what it is |
| --- | --- |
| `manifest.xml` | production manifest; URLs point at the GitHub Pages address below (v1.1.0.0) |
| `manifest.dev.xml` | the same add-in pointed at https://localhost:3000 for testing (different add-in Id) |
| `src/taskpane/taskpane.html` / `.css` / `.js` | the pane and its Excel wiring |
| `src/taskpane/spi-core.js` | the planning logic, ported from the web generator's `logic.js` |
| `src/taskpane/rfc-grid.js` | adapter over a block of live workbook values (the RFC worksheet source) |
| `src/taskpane/rfc-sheetjs.js` | adapter over a SheetJS worksheet (the RFC file source) |
| `src/taskpane/rfc-fingerprint.js` | the identity (`tasks, pay items, quantities`) used for idempotent imports |
| `src/taskpane/xlsx.full.min.js` | SheetJS, vendored (Apache-2.0; see `SheetJS-LICENSE.txt`) |
| `src/commands/` | the manifest's FunctionFile (no commands; the pane is the add-in) |
| `src/legal/privacy.html`, `src/legal/terms.html` | the pages the store listing links to |
| `assets/` | icons, made from the Moran Rocks logo |
| `index.html` | landing page for the hosting root |
| `tests/` | plan equivalence against the web generator, and source/guard tests (below) |
| `tools/` | manifest validation and the local dev server |
| `SUBMISSION.md` | store listing text and the submission checklist |

## Testing by sideloading

Sideloading loads the add-in from a manifest file, without the store.

With the local server (development manifest):

```sh
npm install          # once
npx office-addin-dev-certs install   # once; Excel requires HTTPS
npm start            # serves this folder at https://localhost:3000
```

Then in Excel (desktop): **Insert → Add-ins → Upload My Add-in**, and
choose `manifest.dev.xml`. The RFC Importer button appears on the Home
tab. Open an SPI workbook, open the pane, pick an RFC source, and
Import.

To test the production manifest before the v1.1 store listing goes
live, host this folder at the address in `manifest.xml` (next section)
and sideload `manifest.xml` the same way.

## Hosting plan

The add-in's web files are static. They live in the GitHub repository
`namoran/SPI_Generator_Addin`, served by GitHub Pages from the project
root — the address `manifest.xml` points at:

```
https://namoran.github.io/SPI_Generator_Addin/
```

Code updates deploy by pushing to that repository; every installed copy
picks them up with no store resubmission. Only a manifest change (new
permissions or a different structure) goes back through the store. The
v1.1 rename is a manifest change: see `SUBMISSION.md` for the steps.

## Tests

```sh
npm test           # plan equivalence against the web generator
npm run test:source  # RFC sources (matrix vs SheetJS) and fingerprints
npm run validate   # manifest well-formedness and submission fields
```

`npm test` runs the six real FDEP packages in `../handoff/packages/`.
For each one it checks that the add-in core plans exactly the cells the
web generator plans (with the RFC read both ways the tools read it,
including through SheetJS), that the web generator's output for that
plan matches the SPI FDEP issued, cell for cell across the pay item
table, and that the report lines match.

`npm run test:source` is self-contained (no handoff fixtures). It checks
that the RFC sheet read from the live workbook (the values matrix the
pane gets, `rfc-grid.js`) loads RFC items identical to the SheetJS file
path for the same cell values, and that the import fingerprint is
order- and whitespace-insensitive while still changing with any change
to a task, pay item, or quantity.

There is no Excel in the build environment, so the Office.js calls in
`taskpane.js` are not covered by the automated test. The in-Excel run —
button, pane, writes landing in a live workbook, the fingerprint refusal
and its "Apply anyway" override — is verified by sideloading, per the
steps above.

## Relationship to the other tools

The web SPI Generator stays as it is; this project does not modify it.
The logic port is one-way: `spi-core.js` was ported from the web tool's
`logic.js` on 2026-11-03. If the web tool's logic changes, the port and
this test need the same change.

## Contact

Nicholas Moran · Moran Rocks LLC · https://moran.rocks · 786-263-9185
