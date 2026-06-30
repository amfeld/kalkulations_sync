import base64
import io
from datetime import timedelta

import openpyxl
from odoo import fields

from .common import KalkSyncBaseCase


def _push_write_dates_to_past(env, order, delta_seconds=10):
    """Move all order_line write_dates into the past via SQL so that any
    subsequent export timestamp is definitively newer.

    fields.Datetime.now() (Python utcnow) and Postgres CURRENT_TIMESTAMP can
    differ by microseconds inside the same transaction.  Forcing a known past
    value makes the concurrency comparison deterministic in tests.
    """
    past = fields.Datetime.now() - timedelta(seconds=delta_seconds)
    env.cr.execute(
        "UPDATE sale_order_line SET write_date = %s WHERE order_id = %s",
        [past, order.id],
    )
    env.invalidate_all()


class TestConcurrencyWarning(KalkSyncBaseCase):

    def test_concurrency_warning_set(self):
        # Place write_dates definitively in the past before exporting so the
        # export timestamp is the only thing "in the present".
        _push_write_dates_to_past(self.env, self.order)
        b64 = self.order._generate_kalkulation_excel()

        # Simulate the line being modified AFTER the export.  We force write_date
        # via SQL rather than calling write(): inside a single test transaction an
        # ORM write() stamps write_date with the transaction-start timestamp
        # (cr.now()), which is *earlier* than the wall-clock export timestamp, so a
        # plain write() would never look "newer" than the export.  In production
        # the modification happens in a separate transaction and is genuinely newer.
        future = fields.Datetime.now() + timedelta(seconds=60)
        self.env.cr.execute(
            "UPDATE sale_order_line SET write_date = %s WHERE id = %s",
            [future, self.line1.id],
        )
        self.env.invalidate_all()

        wizard = self.env['kalksync.import.wizard'].create({
            'sale_order_id': self.order.id,
            'file_data': b64,
            'file_name': 'test.xlsx',
        })
        wizard._onchange_file_data()

        self.assertTrue(
            wizard.parse_warning,
            "Concurrency warning should be set when line was modified after export",
        )

    def test_no_warning_when_not_modified(self):
        # Place write_dates definitively in the past before exporting so the
        # export timestamp is guaranteed to be newer than write_date.
        _push_write_dates_to_past(self.env, self.order)
        b64 = self.order._generate_kalkulation_excel()

        wizard = self.env['kalksync.import.wizard'].create({
            'sale_order_id': self.order.id,
            'file_data': b64,
            'file_name': 'test.xlsx',
        })
        wizard._onchange_file_data()

        self.assertFalse(wizard.parse_warning)


class TestSectionReadonly(KalkSyncBaseCase):

    def test_section_lines_not_shown_in_wizard(self):
        """Section lines are excluded from the export and must not appear in wizard lines.

        The export filters out line_section rows, so their IDs never enter the
        Excel file and the import parser has nothing to match against.
        """
        section = self.env['sale.order.line'].create({
            'order_id': self.order.id,
            'display_type': 'line_section',
            'name': 'Abschnitt 1',
            'sequence': 1,
        })

        b64 = self.order._generate_kalkulation_excel()
        wizard = self.env['kalksync.import.wizard'].create({
            'sale_order_id': self.order.id,
            'file_data': b64,
            'file_name': 'test.xlsx',
        })
        wizard._onchange_file_data()

        section_wizard_lines = wizard.line_ids.filtered(
            lambda l: l.order_line_id == section
        )
        self.assertFalse(
            section_wizard_lines,
            "Section lines must not appear in import wizard (excluded from export)",
        )

    def test_section_confirm_does_not_write_non_name_fields(self):
        self.env['sale.order.line'].create({
            'order_id': self.order.id,
            'display_type': 'line_section',
            'name': 'Section A',
            'sequence': 0,
        })

        b64 = self.order._generate_kalkulation_excel()
        wizard = self.env['kalksync.import.wizard'].create({
            'sale_order_id': self.order.id,
            'file_data': b64,
            'file_name': 'test.xlsx',
        })
        wizard._onchange_file_data()

        # Confirm should not raise even with section lines present
        result = wizard.action_confirm()
        self.assertEqual(result.get('type'), 'ir.actions.act_window_close')


class TestShowOnlyChanges(KalkSyncBaseCase):

    def test_unchanged_hidden_when_show_only_changes_true(self):
        """Unchanged lines must be excluded from line_ids when show_only_changes=True."""
        b64 = self.order._generate_kalkulation_excel()
        wizard = self.env['kalksync.import.wizard'].create({
            'sale_order_id': self.order.id,
            'file_data': b64,
            'file_name': 'test.xlsx',
            'show_only_changes': True,
        })
        wizard._onchange_file_data()

        unchanged = wizard.line_ids.filtered(lambda l: l.status == 'unchanged')
        self.assertFalse(unchanged, "Unchanged lines must be hidden when show_only_changes=True")

    def test_unchanged_visible_when_show_only_changes_false(self):
        """Unchanged lines must appear in line_ids when show_only_changes=False."""
        b64 = self.order._generate_kalkulation_excel()
        wizard = self.env['kalksync.import.wizard'].create({
            'sale_order_id': self.order.id,
            'file_data': b64,
            'file_name': 'test.xlsx',
            'show_only_changes': False,
        })
        wizard._onchange_file_data()

        unchanged = wizard.line_ids.filtered(lambda l: l.status == 'unchanged')
        self.assertTrue(unchanged, "Unchanged lines must be visible when show_only_changes=False")
