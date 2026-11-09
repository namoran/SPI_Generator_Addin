# Office Store (AppSource) submission — RFC Importer

State as of 2026-11-09: the v1.0 listing has not been submitted; the
add-in is built and tested live in Excel on the web (pane build 109,
GitHub Pages). GitHub repository `namoran/SPI_Generator_Addin` exists
with Pages serving the root of `main`, and the pane URLs below are
live. Version 1.1.0.0 renamed the add-in from "SPI Generator" to "RFC
Importer", added the RFC-worksheet-in-this-workbook source and the
idempotent-import fingerprint guard, and is a manifest change. Steps 1–2
in the checklist are done except for taking the listing screenshots.
What remains is Nicholas's to do — it needs his Microsoft account.
Registration is outlined step by step below.

## Facts confirmed 2026-11-03

- The store does not host the add-in. You submit the manifest and the
  listing; the manifest points at the add-in's web files, and those stay
  on GitHub Pages. The one hosting requirement is HTTPS, which GitHub
  Pages provides.
- Registering a Partner Center developer account currently carries no
  fee, for individual or company accounts, per Microsoft Learn
  (Microsoft's current sign-up flow). Submitting a free add-in costs
  nothing. Tenant creation asks for a payment method, but that results
  in no charge at sign-up; charges apply only when a purchase is made.
- Marketplace enrollment requires a Microsoft Entra work account.
  Personal Microsoft accounts (Outlook.com / Hotmail / Live.com) are
  not supported for Marketplace publisher enrollment (confirmed by
  Microsoft Q&A, January 2026). Moran Rocks LLC must therefore enroll
  under a work account — an existing tenant, or a new tenant created
  in the enrollment flow (see the walkthrough below).
- Marketplace publishing uses a company account. Moran Rocks LLC is the
  publisher name in the manifest.
- After certification, code updates pushed to GitHub Pages reach users
  with no resubmission. A manifest change (permissions, structure) needs
  a new submission.
- The publisher is responsible for the hosting staying up; if it goes
  down, Microsoft can remove the listing.

## Checklist — Nicholas only

1. ~~**Create the GitHub repository.**~~ DONE. `SPI_Generator_Addin`
   exists under the `namoran` account; Pages serves the root of `main`.
   Confirmed live: https://namoran.github.io/SPI_Generator_Addin/ and
   https://namoran.github.io/SPI_Generator_Addin/src/taskpane/taskpane.html
   (pane build 109).
2. ~~**Sideload the production manifest once.**~~ DONE in Excel on the
   web (same browser pane the listing uses). One item remains: **take
   the screenshots** listed below (PNG, 1366 × 768 or larger).
3. **Create the Partner Center account.** See the registration walkthrough
   below. No fee for a free add-in; a payout/tax profile is only needed
   if the add-in charges.
4. **Create the listing.** In Partner Center, create an Office Add-in
   offer under the **Microsoft 365 and Copilot** tab, upload
   `manifest.xml`, and fill in the listing from the text below. Use
   these URLs:
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

## Registration walkthrough (current Partner Center flow, 2026-11)

To publish an Office Add-in, Moran Rocks LLC enrolls once in Partner
Center; the Microsoft AI Cloud Partner Program (MAICPP) enrollment is
created as part of that same flow. The **Microsoft 365 and Copilot**
program is the one for Office Add-ins (not Commercial Marketplace,
which is for Marketplace offers). Steps below follow the current
[enrollment article](https://learn.microsoft.com/en-us/partner-center/account-settings/create-account);
screen wording may vary slightly.

### Have ready before you start

- A Microsoft Entra **work account** in the tenant your organization
  uses for Partner Center (a personal Microsoft account alone does
  not meet the prerequisite — personal accounts are unsupported for
  Marketplace enrollment). The user who starts MAICPP enrollment
  needs the Global admin tenant role; you sign on as Nicholas Moran
  (786-263-9185, moran.rocks). If Moran Rocks LLC has no tenant yet,
  create one in the enrollment flow: it is free, a payment method is
  collected with no charge at sign-up (charges only if you buy
  something later), and this is the supported path for a company
  without an existing tenant.
- Authority to accept legal agreements on behalf of the LLC.
- Registered legal business name, address, country/region, and primary
  contact information.
- A business registration number (Moran Rocks LLC's Florida record is
  L23000516677, Active — confirm it at Sunbiz before entering).
- Business registration and domain-ownership documents for
  verification (have the Sunbiz certificate of status and proof that
  moran.rocks is owned at hand).
- Publisher name **Moran Rocks LLC** (must match the `ProviderName`
  in `manifest.xml`) and a unique Publisher ID — try `moranrocksllc`
  first; if that ID is taken, the form's validation rules show what is
  allowed.

### Click path

1. Open [Partner Center](https://go.microsoft.com/fwlink/?linkid=2165507),
   select **Build**, then **Next**.
2. Select **Login with work account** and sign in with the work
   account. If there is no work account yet, create one first
   (linked from the enrollment flow), then return.
3. Confirm company, country or region, signed-in account, and tenant.
4. On the program selection page, select **Microsoft 365 and
   copilot**, then **Next**.
5. Select the legal entity or billing account for the business.
6. Review the legal business and primary contact information; enter
   the business registration number when asked.
7. Enter the Publisher name (**Moran Rocks LLC**).
8. Review the **Office Publisher Agreement** and **Microsoft AI Cloud
   Partner Program Agreement**, accept, and select **Enroll**.
9. Complete verification from the **Verification Summary** page
   (email/identity/business checks). Verification usually takes
   **three to five business days** and can take longer with extra
   review.
10. Confirm enrollment under **Account settings > Programs**, and the
    publisher under **Account settings > Organization profile >
    Identifiers > Publisher**. Then use the account picker to select
    that Seller ID (the seller context) before opening
    **Marketplace offers > Microsoft 365 and Copilot** and starting
    the offer.

### What it costs and what it does not

- Registration and submitting a free add-in cost nothing under the
  current flow. No payout/tax setup is needed unless the add-in charges.
- Creating a work account or submitting the enrollment request does
  **not** complete verification — MAICPP enrollment becomes active
  only after verification inputs are submitted and authorized.

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
