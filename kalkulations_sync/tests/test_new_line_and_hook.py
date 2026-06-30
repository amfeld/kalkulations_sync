"""Tests for new-line creation defaults and the install seeding hook.

Covers two previously-untested pieces of real logic:
- action_confirm's new-line path applies the company default product AND propagates
  its unit of measure, and reports the new-line count in the chatter message.
- _seed_default_template only fills companies that have no template yet (never
  overwrites a customer-uploaded template).
"""

import base64
import io

import openpyxl
from openpyxl.utils import column_index_from_string

from .common import KalkSyncBaseCase
from ..hooks import _seed_default_template


def _corrupt_id_to_n(b64_file, line_id):
    wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(b64_file)), data_only=True)
    ws = wb.active
    meta = wb['kalksync_meta']
    data_row_start = int(meta['B6'].value)
    id_col_idx = column_index_from_string(meta['B4'].value)
    for row in ws.iter_rows(min_row=data_row_start, values_only=False):
        if row[id_col_idx - 1].value == line_id:
            row[id_col_idx - 1].value = 'N'
            break
    buf = io.BytesIO()
    wb.save(buf)
    return base64.b64encode(buf.getvalue()).decode()


class TestNewLineDefaults(KalkSyncBaseCase):

    def test_new_line_uses_default_product_and_uom(self):
        """An N-row without a product gets the company default product and its UoM."""
        b64 = self.order._generate_kalkulation_excel()
        b64_n = _corrupt_id_to_n(b64, self.line1.id)
        existing_ids = self.order.order_line.ids

        wizard = self.env['kalksync.import.wizard'].create({
            'sale_order_id': self.order.id,
            'file_data': b64_n,
            'file_name': 'new.xlsx',
        })
        wizard._onchange_file_data()
        wizard.action_confirm()

        new_line = self.order.order_line.filtered(lambda l: l.id not in existing_ids)
        self.assertEqual(len(new_line), 1)
        self.assertEqual(new_line.product_id, self.product)
        self.assertEqual(new_line.product_uom_id, self.product.uom_id)

    def test_new_line_sequence_after_existing(self):
        """New line gets a sequence greater than every existing line (appended, not on top)."""
        b64 = self.order._generate_kalkulation_excel()
        b64_n = _corrupt_id_to_n(b64, self.line1.id)
        max_seq_before = max(self.order.order_line.mapped('sequence'))
        existing_ids = self.order.order_line.ids

        wizard = self.env['kalksync.import.wizard'].create({
            'sale_order_id': self.order.id,
            'file_data': b64_n,
            'file_name': 'new.xlsx',
        })
        wizard._onchange_file_data()
        wizard.action_confirm()

        new_line = self.order.order_line.filtered(lambda l: l.id not in existing_ids)
        self.assertGreater(new_line.sequence, max_seq_before)

    def test_new_line_count_in_chatter(self):
        """The confirm message must mention that a new line was created."""
        b64 = self.order._generate_kalkulation_excel()
        b64_n = _corrupt_id_to_n(b64, self.line1.id)

        wizard = self.env['kalksync.import.wizard'].create({
            'sale_order_id': self.order.id,
            'file_data': b64_n,
            'file_name': 'new.xlsx',
        })
        wizard._onchange_file_data()
        wizard.action_confirm()

        last_msg = self.order.message_ids[0]
        self.assertIn('newly created', last_msg.body)


class TestSeedTemplateHook(KalkSyncBaseCase):

    def test_seed_fills_empty_company(self):
        """A company without a template gets the bundled default seeded in."""
        company = self.env['res.company'].create({'name': 'Seed-Test GmbH'})
        company.amf_kalksync_template = False
        self.assertFalse(company.amf_kalksync_template)

        _seed_default_template(self.env)

        self.assertTrue(company.amf_kalksync_template)
        self.assertEqual(company.amf_kalksync_template_name, 'Vorlage_Kalkulation.xlsx')

    def test_seed_does_not_overwrite_existing_template(self):
        """A company that already uploaded a template must keep it untouched."""
        custom = base64.b64encode(b'custom-template-bytes')
        company = self.env['res.company'].create({'name': 'Keep-Mine GmbH'})
        company.amf_kalksync_template = custom
        company.amf_kalksync_template_name = 'mein_template.xlsx'

        _seed_default_template(self.env)

        self.assertEqual(company.amf_kalksync_template, custom)
        self.assertEqual(company.amf_kalksync_template_name, 'mein_template.xlsx')
