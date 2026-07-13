{
    'name': 'Kalkulations-Sync Excel',
    'version': '19.0.1.4.0',
    'category': 'Sales',
    'summary': 'Universal Excel export/import for quotation lines — any field, including custom',
    'author': 'Alex Feld',
    'website': 'https://github.com/amfeld/kalkulations_sync',
    'license': 'LGPL-3',
    'depends': ['sale_management', 'mail', 'base_setup'],
    'external_dependencies': {'python': ['openpyxl']},
    'data': [
        'security/ir.model.access.csv',
        'views/res_config_settings_views.xml',
        'views/import_wizard_views.xml',
        'views/sale_order_views.xml',
    ],
    'description': """
Kalkulations-Sync Excel
=======================

Export quotation lines as an Excel calculation sheet and import the changed values
back into Odoo. Driven entirely by the template — no code needed for new fields.


Workflow
--------
1. **Upload a template** (once per company):
   Settings → Sales → Kalkulations-Sync → upload an Excel template

2. **Export the calculation:**
   Quotation (draft or sent) → **"Export Calculation"** button in the form header

3. **Calculate externally:**
   Adjust the values, open and save the file in Microsoft Excel,
   do not change the ID column.

4. **Import the calculation:**
   Quotation → **"Import Calculation"** button
   → upload the file → review the differences → Confirm


Highlights
----------
- Any ``sale.order.line`` field is importable (standard and custom)
- Import columns are declared via ``[field]`` markers in a header row;
  ``{{line.field}}`` placeholders export values (display only)
- New positions: a row with ``N`` in the ID column creates a new order line
- Concurrency guard: warns if a line was changed in Odoo after the export
- Computed/relational fields are skipped automatically
- Compatible with Community and Enterprise


Supported field types
----------------------
- Float / Integer / Monetary — numeric comparison with tolerance 1e-6
- Char / Text / Html — string comparison
- Boolean — Yes/No (Excel: ``Yes``, ``1``, ``True``; German ``Ja``/``Wahr`` also accepted)
""",

    'post_init_hook': 'post_init_hook',
    'installable': True,
    'auto_install': False,
    'application': False,
}
