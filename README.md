# kalkulations_sync

Template-driven Excel round-trip for quotation lines (`sale.order.line`) in
**Odoo 19** (Community or Enterprise).

Export quotation lines into a customizable Excel calculation sheet, refine pricing
and quantities in a spreadsheet, then import the changed values back — with a
per-field difference preview before anything is written. Which fields are exchanged
is decided entirely by the Excel template, not by code.

## The Odoo module

The installable module lives in [`kalkulations_sync/`](kalkulations_sync/). Add the
**repository root** to your Odoo `addons_path`, then install **Kalkulations-Sync
Excel** from the Apps menu.

- Module overview, features & quick start: [`kalkulations_sync/README.md`](kalkulations_sync/README.md)
- User guide (English): [`kalkulations_sync/docs/USER_GUIDE_en.md`](kalkulations_sync/docs/USER_GUIDE_en.md)
- User guide (German): [`kalkulations_sync/docs/USER_GUIDE_de.md`](kalkulations_sync/docs/USER_GUIDE_de.md)
- Changelog: [`kalkulations_sync/CHANGELOG.md`](kalkulations_sync/CHANGELOG.md)

## Requirements

- Odoo 19.0 (Community or Enterprise)
- Python package `openpyxl`
- Odoo dependency: `sale_management`

## License

Licensed under the **GNU Lesser General Public License v3.0 (LGPL-3)** — see
[`LICENSE`](LICENSE).
