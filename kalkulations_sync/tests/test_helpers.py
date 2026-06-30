"""Unit tests for pure-Python helpers in models/sale_order.py and wizard/import_wizard.py."""

from odoo.tests.common import TransactionCase

from ..models.sale_order import (
    _replace_line_placeholders,
    _resolve_header_marker,
    _resolve_path,
    _shift_formula_below_data,
    _to_cell_value,
)
from ..wizard.import_wizard import KalkSyncImportWizard, _coerce_field_value

from .common import KalkSyncBaseCase


# ---------------------------------------------------------------------------
# _shift_formula_below_data
# ---------------------------------------------------------------------------

class TestShiftFormulaBelowData(TransactionCase):
    """Verifies that formula refs in rows below the data block are shifted correctly."""

    def test_sum_range_end_shifted(self):
        """Range end >= master_row is expanded by offset; range start stays anchored."""
        self.assertEqual(
            _shift_formula_below_data('=SUM(A3:A5)', 3, 2), '=SUM(A3:A7)'
        )

    def test_single_ref_at_master_row_shifted(self):
        self.assertEqual(_shift_formula_below_data('=A3', 3, 2), '=A5')

    def test_single_ref_above_master_row_unchanged(self):
        """Refs strictly below master_row (header rows) must not be touched."""
        self.assertEqual(_shift_formula_below_data('=A2', 3, 2), '=A2')

    def test_single_ref_below_master_row_shifted(self):
        self.assertEqual(_shift_formula_below_data('=A10', 3, 2), '=A12')

    def test_absolute_row_ref_unchanged(self):
        self.assertEqual(_shift_formula_below_data('=$A$3', 3, 2), '=$A$3')

    def test_absolute_col_relative_row_shifted(self):
        """$Col + relative row → only the row gets shifted."""
        self.assertEqual(_shift_formula_below_data('=$A3', 3, 2), '=$A5')

    def test_mixed_refs_in_formula(self):
        """Formula with refs both above and at master_row shifts only qualifying ones."""
        result = _shift_formula_below_data('=A3+B2', 3, 2)
        self.assertIn('A5', result)
        self.assertIn('B2', result)

    def test_non_formula_unchanged(self):
        self.assertEqual(_shift_formula_below_data('hello', 3, 2), 'hello')
        self.assertEqual(_shift_formula_below_data(42, 3, 2), 42)


# ---------------------------------------------------------------------------
# _resolve_path
# ---------------------------------------------------------------------------

class TestResolvePath(KalkSyncBaseCase):
    """Verifies that _resolve_path walks declared fields and blocks private attrs."""

    def test_direct_field(self):
        result = _resolve_path(self.order, 'name')
        self.assertEqual(result, self.order.name)

    def test_dotted_path(self):
        result = _resolve_path(self.order, 'partner_id.name')
        self.assertEqual(result, self.partner.name)

    def test_nonexistent_field_returns_none(self):
        self.assertIsNone(_resolve_path(self.order, 'nonexistent_field_xyz'))

    def test_private_attr_blocked(self):
        """_cr, env, sudo are Python attrs, not declared Odoo fields — must be blocked."""
        self.assertIsNone(_resolve_path(self.order, '_cr'))
        self.assertIsNone(_resolve_path(self.order, 'env'))

    def test_falsy_record_returns_none(self):
        self.assertIsNone(_resolve_path(False, 'name'))


# ---------------------------------------------------------------------------
# _to_cell_value
# ---------------------------------------------------------------------------

class TestToCellValue(TransactionCase):

    def test_none_returns_none(self):
        self.assertIsNone(_to_cell_value(None, 'name'))

    def test_false_returns_none(self):
        self.assertIsNone(_to_cell_value(False, 'name'))

    def test_none_for_id_field_returns_zero(self):
        self.assertEqual(_to_cell_value(None, 'id'), 0)

    def test_false_for_id_field_returns_zero(self):
        self.assertEqual(_to_cell_value(False, 'id'), 0)

    def test_bool_true_returns_yes(self):
        self.assertEqual(_to_cell_value(True, 'active'), 'Yes')

    def test_plain_int_passthrough(self):
        self.assertEqual(_to_cell_value(42, 'foo'), 42)

    def test_plain_string_passthrough(self):
        self.assertEqual(_to_cell_value('hello', 'foo'), 'hello')

    def test_plain_float_passthrough(self):
        self.assertAlmostEqual(_to_cell_value(3.14, 'foo'), 3.14)


# ---------------------------------------------------------------------------
# _replace_line_placeholders
# ---------------------------------------------------------------------------

class TestReplaceLinePlaceholders(KalkSyncBaseCase):

    def test_single_placeholder_returns_native_type(self):
        """Entire-cell placeholder {{line.price_unit}} → Python float, not string."""
        result = _replace_line_placeholders('{{line.price_unit}}', self.line1)
        self.assertAlmostEqual(result, self.line1.price_unit)

    def test_mixed_placeholder_returns_string(self):
        """Placeholder embedded in text → string substitution."""
        result = _replace_line_placeholders('Preis: {{line.price_unit}} EUR', self.line1)
        self.assertIsInstance(result, str)
        self.assertIn(str(self.line1.price_unit), result)

    def test_non_string_unchanged(self):
        self.assertEqual(_replace_line_placeholders(42, self.line1), 42)
        self.assertIsNone(_replace_line_placeholders(None, self.line1))

    def test_no_placeholder_unchanged(self):
        self.assertEqual(_replace_line_placeholders('=A1+B1', self.line1), '=A1+B1')

    def test_missing_field_in_mixed_string_replaced_with_empty(self):
        """Unknown field path in mixed string → replaced with empty string."""
        result = _replace_line_placeholders('X{{line.no_such_field}}Y', self.line1)
        self.assertEqual(result, 'XY')


# ---------------------------------------------------------------------------
# _coerce_field_value (wizard/import_wizard.py)
# ---------------------------------------------------------------------------

class TestCoerceFieldValue(TransactionCase):

    def test_none_returns_none(self):
        self.assertIsNone(_coerce_field_value('float', None))

    def test_german_decimal_comma_to_float(self):
        """German comma-decimal strings must be parsed correctly."""
        self.assertAlmostEqual(_coerce_field_value('float', '1,5'), 1.5)

    def test_float_passthrough(self):
        self.assertAlmostEqual(_coerce_field_value('float', 3.14), 3.14)

    def test_integer_truncates_float_string(self):
        self.assertEqual(_coerce_field_value('integer', '42.7'), 42)

    def test_integer_from_german_comma(self):
        self.assertEqual(_coerce_field_value('integer', '10,0'), 10)

    def test_boolean_ja(self):
        self.assertTrue(_coerce_field_value('boolean', 'Ja'))
        self.assertTrue(_coerce_field_value('boolean', 'ja'))

    def test_boolean_wahr(self):
        self.assertTrue(_coerce_field_value('boolean', 'wahr'))

    def test_boolean_true_int(self):
        self.assertTrue(_coerce_field_value('boolean', True))
        self.assertTrue(_coerce_field_value('boolean', 1))

    def test_boolean_false_values(self):
        self.assertFalse(_coerce_field_value('boolean', False))
        self.assertFalse(_coerce_field_value('boolean', 0))
        self.assertFalse(_coerce_field_value('boolean', 'nein'))

    def test_char_int_becomes_string(self):
        self.assertEqual(_coerce_field_value('char', 42), '42')

    def test_char_passthrough(self):
        self.assertEqual(_coerce_field_value('char', 'hello'), 'hello')


# ---------------------------------------------------------------------------
# _fmt_num (KalkSyncImportWizard static method)
# ---------------------------------------------------------------------------

class TestFmtNum(TransactionCase):

    def test_zero(self):
        self.assertEqual(KalkSyncImportWizard._fmt_num(0.0), '0')

    def test_integer_value_no_decimal(self):
        self.assertEqual(KalkSyncImportWizard._fmt_num(50.0), '50')

    def test_trailing_zeros_stripped(self):
        self.assertEqual(KalkSyncImportWizard._fmt_num(1.1), '1.1')

    def test_negative_value(self):
        self.assertEqual(KalkSyncImportWizard._fmt_num(-2.5), '-2.5')

    def test_full_precision(self):
        self.assertEqual(KalkSyncImportWizard._fmt_num(3.14159), '3.14159')


# ---------------------------------------------------------------------------
# _resolve_header_marker
# ---------------------------------------------------------------------------

class TestResolveHeaderMarker(KalkSyncBaseCase):
    """Verifies German Klarname + technical name lookups in header marker resolution."""

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.line_fields = cls.env['sale.order.line']._fields

    def test_technical_field_name(self):
        self.assertEqual(
            _resolve_header_marker('price_unit', self.line_fields), 'price_unit'
        )

    def test_german_preis_je_einheit(self):
        self.assertEqual(
            _resolve_header_marker('Preis je Einheit', self.line_fields), 'price_unit'
        )

    def test_german_preis_je_me(self):
        self.assertEqual(
            _resolve_header_marker('Preis je ME', self.line_fields), 'price_unit'
        )

    def test_german_menge(self):
        self.assertEqual(
            _resolve_header_marker('Menge', self.line_fields), 'product_uom_qty'
        )

    def test_german_bezeichnung(self):
        self.assertEqual(
            _resolve_header_marker('Bezeichnung', self.line_fields), 'name'
        )

    def test_line_prefix_stripped(self):
        """[line.price_unit] style → prefix stripped, then resolved normally."""
        self.assertEqual(
            _resolve_header_marker('line.price_unit', self.line_fields), 'price_unit'
        )

    def test_case_insensitive(self):
        self.assertEqual(
            _resolve_header_marker('BEZEICHNUNG', self.line_fields), 'name'
        )

    def test_unknown_marker_returns_none(self):
        self.assertIsNone(
            _resolve_header_marker('nonexistent_field_xyz_999', self.line_fields)
        )

    def test_non_importable_field_returns_none(self):
        """Many2one fields are not importable; their marker must resolve to None."""
        self.assertIsNone(
            _resolve_header_marker('order_id', self.line_fields)
        )
