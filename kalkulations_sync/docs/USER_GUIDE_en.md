# Kalkulations-Sync Excel – User Guide

This guide covers the full workflow: export quotation lines to an Excel file,
calculate externally, and import the changed values back into the quotation.

The button and field labels referenced here (for example **Export Calculation**,
**Confirm**, **Excel template**) are shown for an English interface. The module
ships with a complete German translation, so on a German installation the same
controls read **Kalkulation exportieren**, **Bestätigen** and **Excel-Template**.

## Contents

1. [Prerequisites](#1-prerequisites)
2. [Setup](#2-setup)
3. [Building a template](#3-building-a-template)
4. [Exporting a calculation](#4-exporting-a-calculation)
5. [External editing](#5-external-editing)
6. [Importing a calculation](#6-importing-a-calculation)
7. [Creating new positions](#7-creating-new-positions)
8. [Supported and skipped field types](#8-supported-and-skipped-field-types)
9. [Troubleshooting](#9-troubleshooting)

---

## 1. Prerequisites

- The **Kalkulations-Sync Excel** module is installed.
- You work in the **Sales** app on a quotation in state **Quotation** or **Quotation
  Sent**. The buttons are hidden in any other state (for example *Sales Order*).
- Your user belongs to the **Sales / User** group. Without it the buttons are not
  visible.
- **Microsoft Excel** is available for the external editing step (see
  [section 5](#5-external-editing) for the background on formula caching).

---

## 2. Setup

These steps are usually performed once per company.

### 2.1 Upload an Excel template

**Goal:** Define what the exported calculation file looks like and which fields are
importable.

**Prerequisite:** Settings access for the Sales configuration.

**Steps:**

1. Open **Settings**.
2. Select the **Sales** tab.
3. Locate the **Kalkulations-Sync** block.
4. Next to **Excel template**, click the upload control and select your `.xlsx`
   file.
5. Click **Save** at the top.

![Settings – Kalkulations-Sync block](img/settings.png)

**Result:** The template is stored and used for every export.

**Notes:**

- On a fresh installation a bundled template (`Vorlage_Kalkulation.xlsx`) is seeded
  into every company that has no template yet. You can start immediately or replace
  it with your own.
- The template must contain the placeholder `{{line.id}}` in exactly one cell;
  otherwise the export fails with an error (see
  [section 3](#3-building-a-template)).

### 2.2 Default product for new positions (optional)

**Goal:** Define a fallback product used when a new position is created on import
**without** its own product reference.

**Steps:**

1. In the same **Kalkulations-Sync** block, open **Standard-Produkt für neue
   Positionen** (default product for new positions).
2. Select a product and click **Save**.

**Result:** New positions (added via `N` in the ID column or a copied row) that do
not carry their own product are created with this product.

**Notes:**

- If no default product is configured and a new row carries no product, the **entire
  import is aborted** with an error. For plain quantity/price updates of existing
  positions this field is not needed.

---

## 3. Building a template

The template is an ordinary Excel file with placeholders and markers. It alone
decides **what** is exported and imported — no code change is required.

### 3.1 Mandatory: `{{line.id}}`

Exactly **one cell** must contain `{{line.id}}`. That cell marks the **master row** —
the row Odoo duplicates for every quotation line.

- Without `{{line.id}}` the export fails with an error.
- The ID cell must **never** be edited manually in the exported file (exception: `N`
  for new positions, see [section 7](#7-creating-new-positions)).

### 3.2 Variant A – `{{line.field}}` placeholders (simplest method)

Place placeholders directly in the master row. On export the value is filled in; on
import the changed value is read back — provided the field is writable.

```
Column A:   {{line.id}}
Column B:   {{line.product_id.name}}
Column C:   {{line.product_uom_qty}}
Column D:   {{line.price_unit}}
Column F:   {{line.name}}
```

Example placeholders:

| Placeholder | Meaning |
|---|---|
| `{{line.id}}` | Position ID (mandatory) |
| `{{line.product_uom_qty}}` | Quantity |
| `{{line.price_unit}}` | Sales price |
| `{{line.purchase_price}}` | Cost price (only with `sale_margin` installed) |
| `{{line.name}}` | Position description |
| `{{line.<any_field>}}` | Any other `sale.order.line` field |

### 3.3 Variant B – `[field]` markers in a header row (for formula columns)

If a value is **not** written via a placeholder but computed by a formula in the
master row, the column can still be importable. Place a marker in square brackets in
a row **above** the master row, in the **same column** as the formula.

```
Row N-1 (markers):  [product_uom_qty]   [price_unit]   GP %
Row N   (labels):   Quantity            Sales price    (no action)
Row N+1 (master):   {{line.id}}   =qty formula   =price formula   =GP formula
```

- The `line.` prefix is optional: `[price_unit]` and `[line.price_unit]` are
  equivalent.
- German label aliases are recognized (case-insensitive):

  | Marker | Field |
  |---|---|
  | `[Menge]` | `product_uom_qty` |
  | `[Preis je Einheit]` / `[Preis je ME]` | `price_unit` |
  | `[Kosten je Einheit]` / `[Kosten je ME]` | `purchase_price` |
  | `[Bezeichnung]` | `name` |
  | `[x_gaeb_menge]` | the technical field name directly |

- Columns **without** a marker and without a `{{line.field}}` placeholder are ignored
  on import (for example a pure GP-% formula column).
- If markers appear in several rows above the master row, the one closest to the
  master row wins.

### 3.4 Header placeholders `{{object.field}}`

Outside the master row you can embed quotation-level values. They are filled in on
export and **ignored** on import:

```
{{object.name}}              → Quotation number
{{object.partner_id.name}}   → Customer name
{{object.date_order}}        → Quotation date
{{object.user_id.name}}      → Responsible salesperson
```

### 3.5 Formulas and sums

- Formulas in the master row are copied per position on export; relative row
  references (for example `=C5*D5`) are shifted automatically (`=C6*D6`, `=C7*D7`, …).
- Absolute references with `$` (for example `$D$1`) are left unchanged.
- Sum formulas below the data block (for example `=SUM(C5:C5)`) expand automatically
  to cover the whole range (`=SUM(C5:C[last row])`).

### 3.6 Checklist before uploading

- [ ] `{{line.id}}` is present in exactly one cell.
- [ ] Every importable column has either `{{line.field}}` in the master row or
      `[field]` in a row above it.
- [ ] Markers for formula columns sit in the same column as the formula.
- [ ] No markers on computed (non-stored) fields — they are skipped on import anyway.
- [ ] The file opens in Excel and all formulas calculate.

> **Bundled example:** A ready-to-use template demonstrating both variants ships at
> `static/templates/kalkulation_template.xlsx`. You can upload it directly under
> *Settings → Sales → Kalkulations-Sync → Excel template*.

---

## 4. Exporting a calculation

**Goal:** Download the quotation's positions as a filled Excel file.

**Prerequisite:** Quotation in state **Quotation** or **Quotation Sent**, a template
is configured, at least one order line exists.

**Steps:**

1. Open the quotation in the **Sales** app.
2. In the form header, click **Export Calculation**.
   *(Alternatively via the gear/Actions menu → "⬇ Export Calculation".)*

**Result:**

- The file downloads immediately. The file name follows the pattern
  `YYMMDD_Kalk_<Customer>_<Quotation number>.xlsx`.
- The file is also stored as an attachment in the quotation's **chatter** and a log
  entry is created.

**Notes:**

- Section and note lines are not exported as positions.
- A hidden sheet `kalksync_meta` stores the timestamp, quotation ID and column
  mapping. **Do not delete this sheet** — it is required for import.

---

## 5. External editing

**Goal:** Adjust the calculation outside Odoo (quantities, prices, other mapped
fields).

**Steps:**

1. Open the exported `.xlsx` file in **Microsoft Excel**.
2. Adjust values in the mapped columns.
3. **Save** the file (in `.xlsx` format).

**Important:**

- **Do not change the ID column.** Each row is matched by its position ID. If the ID
  is overwritten or deleted, the row cannot be imported (exception: intentional `N`
  rows for new positions).
- **Save after calculating.** Excel computes formulas and stores the results in an
  internal cache. That cache is only written when you **save**. If the file is not
  opened and saved in Excel, pure formula cells return **no** value on import and the
  affected rows are reported as errors.
- Do not delete the hidden `kalksync_meta` sheet.

**Result:** A saved file with calculated values, ready for import.

---

## 6. Importing a calculation

**Goal:** Write the changed values from the Excel file back into the quotation.

**Prerequisite:** The same file (or one produced from the same export), saved in
Excel.

**Steps:**

1. Open the quotation and click **Import Calculation** in
   the form header.
2. In the **Import Calculation** dialog, upload the file under **Excel file
   (.xlsx)**.
3. Review the **difference view** (see below).
4. Click **Confirm**.

![Import dialog with the info banner](img/import_wizard.png)

**Reading the difference view:**

After upload, a table shows the **Odoo value** and **Excel value** per changed field,
with a **difference**. The count badges at the top summarize how many rows have each
status. Color coding:

| Status | Meaning |
|---|---|
| **Changed** | The Excel value differs from Odoo and will be applied. |
| **Unchanged** | No difference (hidden by default). |
| **New** | A row with `N` → a new position will be created. |
| **Missing** | The position exists in Odoo but is absent from the file (warning, no change). |
| **Error** | The row cannot be imported (for example a changed ID or a missing formula value). |
| **Ignored** | Section rows and similar are not changed. |

- The **Show only changes** toggle hides or shows unchanged
  rows.
- As long as **Error** rows exist, the **Confirm** button is unavailable. Fix the
  Excel file first and upload it again.

**Result:**

- Changed values are written to the quotation lines; new positions are created.
- The imported file is stored as a chatter attachment and a log entry (number of
  updated/new positions) is recorded.

**Note on concurrent editing:**

If a position was changed directly in Odoo **after** the export, a yellow warning in
the dialog lists the affected position. The import still overwrites the position with
the Excel value when you confirm — in that case verify that the Excel value is
actually the one you want.

---

## 7. Creating new positions

**Goal:** Add extra quotation lines through the Excel file.

**Steps:**

1. Insert a new row in the Excel file (the easiest way is to copy an existing
   position row).
2. Put the letter **`N`** (or `n`) in the **ID cell** of that row.
3. Enter the desired values in the mapped columns.
4. Save the file and import it as in [section 6](#6-importing-a-calculation).

**ID-column behavior on import:**

| ID cell | Behavior |
|---|---|
| empty | Row is silently ignored (for example sum/blank rows). |
| `N` or `n` | A new position is created. |
| number (existing ID) | The existing position is updated. |
| number that appears twice | Treated as a copied = new position. |
| any other text | Error → this row is blocked from import. |

**Result:** On confirm, new positions are created with all importable fields from the
row.

**Notes:**

- If the new row carries no product (no `{{line.product_id}}` mapping), the **default
  product for new positions** is used (see
  [section 2.2](#22-default-product-for-new-positions-optional)). If none is
  configured, the import aborts.

---

## 8. Supported and skipped field types

**Importable:**

| Type | Compared as | Example fields |
|---|---|---|
| Float / Integer / Monetary | Numeric, tolerance `1e-6` | `product_uom_qty`, `price_unit` |
| Char / Text / Html | String (exact) | `name`, `x_gaeb_oz` |
| Boolean | Yes/No | `x_gaeb_manuell` |

**Boolean values in Excel:** `Ja`, `ja`, `1`, `True`, `yes`, `wahr` yield **true**;
anything else yields **false**.

**Automatically skipped (even if present in the template):**

- Computed, non-stored fields (for example a GP formula column).
- Relational fields: Many2one, One2many, Many2many.
- Binary and Serialized fields.

> **Note on Many2one fields:** These are intentionally not importable. On export the
> display name is written; a string cannot be reliably resolved back to the correct
> record on import. Such columns are display-only (for example a product name).

---

## 9. Troubleshooting

| Message / symptom | Cause | Resolution |
|---|---|---|
| "No calculation template configured." | No template stored. | Upload a template under *Settings → Sales → Kalkulations-Sync → Excel template*. |
| "The template does not contain a `{{line.id}}` placeholder." | The mandatory placeholder is missing. | Put `{{line.id}}` into exactly one master-row cell and re-upload the template. |
| "Export is only possible in state 'Quotation' or 'Quotation Sent'." | The quotation is already an order/closed. | Work on a quotation in the correct state. |
| "No order lines available." | The quotation has no lines. | Add at least one product position. |
| "This file was not exported with Kalkulations-Sync (missing 'kalksync_meta' sheet)." | Wrong file, or the hidden sheet was deleted. | Re-export and use the original export file; do not remove the `kalksync_meta` sheet. |
| "This file was exported for quotation ID … not for the current quotation." | File of another quotation was uploaded. | Use the file that matches the open quotation. |
| **Error** status: "Formula cell without a calculated value …" | File was not saved in Excel after editing; the formula cache is missing. | Open the file in Microsoft Excel, **save** it, and re-import. |
| **Error** status: "The line ID was modified." | The ID cell was overwritten. | Restore the ID column from the original export file (use `N` for genuinely new rows). |
| **Error** status: "Line ID … is not present in the quotation." | The position was deleted in Odoo or the ID was tampered with. | Check the affected row; re-export if needed. |
| **Error** status: "Invalid value '…' for field '…'." | Text in a numeric column. | Enter a valid number (both comma and dot are accepted as the decimal separator). |
| "New line '…' cannot be created: no product assigned and no default product configured." | An `N` row has no product and no fallback is set. | Configure a default product in settings, or map a product into the row. |
| Yellow warning: "Line '…' was modified in Odoo after the export." | The position was edited directly in Odoo after export. | Verify that the Excel value should really be applied before confirming. |
| Yellow warning: "Line '…' is not present in the Excel file." | A row is missing from the Excel file. | Not an error — the position stays unchanged. Re-export if needed. |
| Excel reports a "corrupted file" on open. | Fixed in current module versions. | Make sure the module is up to date and regenerate the export file. |
