"""Behavioral tests for action_export_kalkulation and the recordset cell helpers.

test_export.py covers the Excel-generation internals; this file covers the public
action's side effects (attachment created, chatter posted, download URL returned)
and the two-level traversal / recordset branches of the cell-value helpers that the
placeholder mechanism relies on.
"""

import base64
import io

import openpyxl

from ..models.sale_order import _resolve_path, _to_cell_value

from .common import KalkSyncBaseCase


class TestExportActionSideEffects(KalkSyncBaseCase):

    def test_export_creates_attachment_and_posts_chatter(self):
        msgs_before = len(self.order.message_ids)
        att_before = self.env['ir.attachment'].search_count([
            ('res_model', '=', 'sale.order'),
            ('res_id', '=', self.order.id),
        ])

        result = self.order.action_export_kalkulation()

        att_after = self.env['ir.attachment'].search([
            ('res_model', '=', 'sale.order'),
            ('res_id', '=', self.order.id),
        ])
        self.assertEqual(len(att_after), att_before + 1)
        self.assertGreater(len(self.order.message_ids), msgs_before)

        # Returned action is a download URL pointing at the created attachment.
        self.assertEqual(result['type'], 'ir.actions.act_url')
        self.assertIn('/web/content/', result['url'])

    def test_exported_attachment_is_valid_xlsx(self):
        """The attachment created by the action must be a readable kalksync workbook."""
        self.order.action_export_kalkulation()
        att = self.env['ir.attachment'].search([
            ('res_model', '=', 'sale.order'),
            ('res_id', '=', self.order.id),
        ], order='id desc', limit=1)
        wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(att.datas)), data_only=True)
        self.assertIn('kalksync_meta', wb.sheetnames)

    def test_export_filename_contains_order_name(self):
        self.order.action_export_kalkulation()
        att = self.env['ir.attachment'].search([
            ('res_model', '=', 'sale.order'),
            ('res_id', '=', self.order.id),
        ], order='id desc', limit=1)
        self.assertIn(self.order.name, att.name)
        self.assertTrue(att.name.endswith('.xlsx'))


class TestResolvePathDepth(KalkSyncBaseCase):

    def test_three_level_path_capped_at_two(self):
        """_resolve_path walks at most two segments; the third is ignored.

        partner_id.country_id.name would be 3 levels — only partner_id.country_id is
        traversed, so the result is the country recordset (or its falsy form), never
        the country name string.
        """
        result = _resolve_path(self.order, 'partner_id.country_id.name')
        # Two-level walk stops at country_id → recordset or empty, never a name string.
        self.assertNotIsInstance(result, str)


class TestToCellValueRecordset(KalkSyncBaseCase):

    def test_recordset_id_path_returns_id(self):
        """field_path ending in 'id' on a recordset returns the numeric id."""
        self.assertEqual(_to_cell_value(self.line1, 'id'), self.line1.id)
        self.assertEqual(
            _to_cell_value(self.line1.product_id, 'product_id.id'),
            self.product.id,
        )

    def test_recordset_non_id_path_returns_display_name(self):
        """A non-id recordset path returns the display_name, not the id."""
        self.assertEqual(
            _to_cell_value(self.line1.product_id, 'product_id'),
            self.product.display_name,
        )

    def test_empty_recordset_returns_none(self):
        empty = self.env['product.product']
        self.assertIsNone(_to_cell_value(empty, 'product_id'))
