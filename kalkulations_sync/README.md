# Kalkulations-Sync Excel

Template-driven Excel round-trip for quotation lines (`sale.order.line`) in Odoo 19.

Export your quotation lines into a customizable Excel calculation sheet, refine
pricing and quantities in a familiar spreadsheet (with formulas, sums and your
own layout), then import the changed values straight back into the quotation —
with a per-field difference preview before anything is written.

Which fields are exchanged is decided **entirely by the Excel template**, not by
code. Standard fields and custom fields added by other modules work the same way;
no Python changes are needed to support a new field.

## Features

- **Universal field support** — any writable `sale.order.line` field (standard or
  custom) can be exported and re-imported just by placing it in the template.
- **Two template mechanics** — `{{line.field}}` placeholders in the master row, or
  `[field]` markers in a header row above the master row (for formula columns).
- **Formula-aware export** — the master row is duplicated per line and relative row
  references in formulas are shifted automatically; `SUM()` ranges expand to cover
  all positions.
- **Difference preview** — the import wizard shows old (Odoo) vs. new (Excel) value
  per field, color-coded by status (changed / new / missing / error), before you
  confirm.
- **New positions** — a row with `N` in the ID column creates a new order line on
  import.
- **Concurrency guard** — if a line was modified in Odoo after the export, the
  wizard warns you instead of silently overwriting.
- **Safe by design** — computed (non-stored), relational, binary and serialized
  fields are skipped automatically.
- **Community and Enterprise** — no Enterprise-only dependency.
- **English UI, German translation included** — source language is English; a
  complete German translation ships in `i18n/de.po`. German `[markers]` such as
  `[Menge]` or `[Preis je Einheit]` are still accepted in templates.

## Screenshots

| Settings block | Import wizard |
|---|---|
| ![Settings](docs/img/settings.png) | ![Import wizard](docs/img/import_wizard.png) |

## Requirements

- Odoo **19.0** (Community **or** Enterprise)
- Python package **`openpyxl`** (`pip install openpyxl`)
- Odoo module dependencies: `sale_management`, `mail`, `base_setup`
- Microsoft Excel (or a spreadsheet application that writes back formula result
  caches) for the external editing step — see the user guide for why this matters.

## Installation

1. Copy the `kalkulations_sync` folder into one of your Odoo `addons` paths.
2. Install `openpyxl` in the Python environment that runs Odoo:
   ```bash
   pip install openpyxl
   ```
3. Restart Odoo and update the app list, then install **Kalkulations-Sync Excel**
   from the Apps menu.

On a fresh install a ready-to-use starter template is seeded into every company
that does not already have one (see *Quick start*).

## Quick start

1. **Upload a template** (once per company): *Settings → Sales → Kalkulations-Sync
   → Excel-Template*. A starter template ships with the module and is pre-filled
   automatically on first install.
2. **Export:** open a quotation in state *Quotation* or *Quotation Sent* and click
   **Export Calculation** in the form header. The filled `.xlsx` is downloaded
   and attached to the chatter.
3. **Calculate externally:** adjust quantities, prices and any other mapped fields
   in Excel. Do **not** change the ID column. Open and **save** the file in Excel so
   that formula results are cached.
4. **Import:** click **Import Calculation**, upload the file, review the
   per-field differences, and press **Confirm**. Changed values are written to
   the quotation lines; rows marked `N` create new lines.

The buttons are also available from the quotation's *Actions* (gear) menu.

## Supported field types & template mechanics

| Field type | Compared as |
|---|---|
| Float / Integer / Monetary | Numeric, tolerance `1e-6` |
| Char / Text / Html | String (exact) |
| Boolean | Yes/No (`Ja`, `1`, `True`, `yes`, `wahr` → true) |

Automatically **skipped** (never written back): computed non-stored fields,
relational fields (Many2one / One2many / Many2many), and Binary / Serialized
fields.

A template must contain the `{{line.id}}` placeholder in exactly one cell (the
*master row*). Columns become importable either by carrying a `{{line.field}}`
placeholder in the master row, or a `[field]` marker in a row above it. The full
template reference — including header placeholders, German label aliases and the
hidden `kalksync_meta` sheet — is in the user guides.

## Documentation

- User guide (German): [`docs/USER_GUIDE_de.md`](docs/USER_GUIDE_de.md)
- User guide (English): [`docs/USER_GUIDE_en.md`](docs/USER_GUIDE_en.md)
- Changelog: [`CHANGELOG.md`](CHANGELOG.md)

## License

This module is licensed under the **GNU Lesser General Public License v3.0
(LGPL-3)**. The full license text is in [`LICENSE`](LICENSE).

## Contributing

Issues and pull requests are welcome. Please keep the module's single external
dependency (`openpyxl`) intact, run the test suite under `tests/` before
submitting, and update `CHANGELOG.md` together with the manifest `version` for any
functional change.
