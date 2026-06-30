"""Tests for the wizard's computed status fields (_compute_counts, _compute_has_errors).

These computes drive the wizard summary line and the confirm-button gating in the
view. The existing suite only checks `has_errors` indirectly; here we pin down the
exact bucket counts and prove they react to the actual line statuses (not tautology:
we build a mix of statuses and assert each counter independently).
"""

from .common import KalkSyncBaseCase


class TestWizardCounts(KalkSyncBaseCase):

    def _wizard_with_lines(self, status_specs):
        """Create a wizard and attach in-memory wizard lines with the given statuses.

        status_specs: list of status strings. Each becomes one wizard line.
        Returns the wizard with its line_ids populated and computes evaluated.
        """
        wizard = self.env['kalksync.import.wizard'].create({
            'sale_order_id': self.order.id,
        })
        wizard.line_ids = [
            (0, 0, {
                'field_name': 'price_unit',
                'field_label': 'X',
                'status': status,
            })
            for status in status_specs
        ]
        return wizard

    def test_counts_split_by_status(self):
        wizard = self._wizard_with_lines(
            ['changed', 'changed', 'error', 'new', 'missing', 'unchanged', 'ignored']
        )
        self.assertEqual(wizard.count_changed, 2)
        self.assertEqual(wizard.count_errors, 1)
        self.assertEqual(wizard.count_new, 1)
        self.assertEqual(wizard.count_missing, 1)

    def test_has_errors_true_with_error_line(self):
        wizard = self._wizard_with_lines(['changed', 'error'])
        self.assertTrue(wizard.has_errors)

    def test_has_errors_false_without_error_line(self):
        wizard = self._wizard_with_lines(['changed', 'new', 'missing', 'unchanged'])
        self.assertFalse(wizard.has_errors)

    def test_counts_zero_when_empty(self):
        wizard = self._wizard_with_lines([])
        self.assertEqual(wizard.count_changed, 0)
        self.assertEqual(wizard.count_errors, 0)
        self.assertEqual(wizard.count_new, 0)
        self.assertEqual(wizard.count_missing, 0)
        self.assertFalse(wizard.has_errors)

    def test_counts_recompute_after_status_change(self):
        """Computes must react when a line's status changes — not a frozen snapshot."""
        wizard = self._wizard_with_lines(['unchanged'])
        self.assertEqual(wizard.count_changed, 0)
        wizard.line_ids[0].status = 'changed'
        self.assertEqual(wizard.count_changed, 1)
