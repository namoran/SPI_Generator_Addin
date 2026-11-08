# Office Store (AppSource) submission — RFC Importer

State as of 2026-11-07: the v1.0 listing has not been submitted; the
add-in is built and its logic is tested against the web SPI Generator
(see README.md, Tests). Version 1.1.0.0 renamed the add-in from "SPI
Generator" to "RFC Importer", added the RFC-worksheet-in-this-workbook
source and the idempotent-import fingerprint guard, and is a manifest
change. The steps below are the remaining ones, and they are Nicholas's
to do — they need his Microsoft account and his GitHub account.

## Facts confirmed 2026-11-03

- The store does not host the add-in. You submit the manifest and the
  listing; the manifest points at the add-in's web files, and those stay
  on GitHub Pages. The one hosting requirement is HTTPS, which GitHub
  Pages provides.
- Registering a Partner Center developer account currently carries no
  fee, for individual or company accounts, per Microsoft Learn
  (Microsoft's current sign-up flow). Submitting a free add-in costs
  nothing.
- Marketplace publishing uses a company account. Moran Rocks LLC is the
  publisher name in the manifest.
- After certification, code updates pushed to GitHub Pages reach users
  with no resubmission. A manifest change (permissions, structure) needs
  a new submission.
- The publisher is responsible for the hosting staying up; if it goes
  down, Microsoft can remove the listing.

## Checklist — Nicholas only

1. **Create the GitHub repository.** Create `SPI_Generator_Addin` under
   the `namoran` account (the Muse GitHub App cannot create
   repositories). Push this project to it as-is, project root at the
   repository root. In the repository settings, enable GitHub Pages
   serving the root of the default branch. Confirm
   https://namoran.github.io/SPI_Generator_Addin/ loads the landing
   page and https://namoran.github.io/SPI_Generator_Addin/src/taskpane/taskpane.html
   loads the pane.
2. **Sideload the production manifest once.** In Excel: Insert →
   Add-ins → Upload My Add-in → `manifest.xml`. Run it against a real
   RFC/SPI pair and confirm the quantities land correctly. Take the
   screenshots listed below during this run.
3. **Create the Partner Center account.** Sign up at Partner Center as a
   company account for Moran Rocks LLC and complete the verification
   Microsoft asks for. No fee at registration under the current flow.
4. **Create the listing.** In Partner Center, create an Office Add-in
   offer, upload `manifest.xml`, and fill in the listing from the text
   below. Use these URLs:
   - Privacy policy: https://namoran.github.io/SPI_Generator_Addin/src/legal/privacy.html
   - Terms of use: https://namoran.github.io/SPI_Generator_Addin/src/legal/terms.html
   - Support: https://moran.rocks
   Upload the icons from `assets/` (icon-128.png is the store-size
   icon; icon-64.png and icon-32.png cover the smaller slots) and the
   screenshots from step 2. In the certification notes, state: "Open
   the sample SPI workbook, open the RFC Importer pane from the Home
   tab, select the sample RFC file, and select Import. The pane lists
   the quantities imported into the SOW Units sheet. Importing the same
   RFC a second time is refused with a message." (Sample RFC/SPI
   files for the reviewers: the pairs in `../handoff/packages/`, e.g.
   C03C59's `rfc.xlsx` and `spi_old.xlsx`.)
5. **Submit for certification.** Microsoft runs automated validation,
   then tests the add-in in Excel. Status moves Pre-processing →
   Certification → Published. If it comes back "Attention needed," the
   report in Partner Center says what to fix; fix and resubmit. After
   certification passes, the listing typically appears in the
   marketplace and in Excel's add-in store within about an hour.

## Screenshots to take (step 2, in Excel, PNG, 1366 × 768 or larger)

1. The RFC Importer button on the Home tab, with an SPI workbook open.
2. The pane open, an RFC selected, showing the line count it found.
3. The pane after Import, showing the list of quantities imported.
4. The SOW Units sheet after Import, showing updated quantities and
   recalculated totals.
5. (Optional, for the certification notes) A second Import attempt,
   showing the "identical tasks, pay items, and quantities" refusal and
   the "Apply anyway" override button.

## Draft store listing text

**Name:** RFC Importer

**Short description:**
Imports pay item quantities from an FDEP RFC into the SPI workbook open
in Excel.

**Long description:**
RFC Importer transfers the line items from a PRP change order (RFC)
into a Schedule of Payment Items (SPI) workbook.

Open the SPI workbook in Excel, open the RFC Importer pane, choose
where the RFC data comes from — the RFC workbook file, or the RFC
worksheet inside the SPI workbook — and select Import. The add-in adds
each RFC quantity into the matching task quantity cell in the SOW Units
sheet, unhides the rows that received a value, and lists every change
it made. Nothing else in the workbook is changed. Excel recalculates
the totals.

Importing the same RFC data a second time changes nothing: the add-in
refuses it with a message (and an "Apply anyway" override), so repeated
imports cannot double the quantities.

The pay item table is read from the workbook itself, so the add-in works
with the SPI layouts in current use, including versions whose pay items
sit on different rows. If a pay item in the RFC is not in the workbook,
or a quantity cell holds a formula, the add-in stops and reports the
problem before writing anything.

The RFC data is read on your computer. The add-in collects no data and
uploads nothing.

RFC Importer (formerly SPI Generator) is published by Moran Rocks LLC.
A web version of the tool is at
https://namoran.github.io/SPI_Generator/.

**Category:** Productivity (Partner Center will offer its current list;
pick the closest fit.)

## Pages already in this project for the listing

- Privacy policy: `src/legal/privacy.html` — states what the code does:
  the RFC is read locally in the pane, quantities are written to the
  open workbook, no data is collected, stored, transmitted, or shared,
  and the only network activity is loading the add-in's own files and
  Microsoft's Office JavaScript library.
- Terms of use: `src/legal/terms.html` — what the add-in does, the
  user's responsibility to review the reported changes before relying on
  the workbook, provided as is, free.
