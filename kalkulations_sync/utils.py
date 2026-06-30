"""Shared helpers used by both models/sale_order.py and wizard/import_wizard.py."""

# Field types that cannot be meaningfully imported/exported as scalar cell values.
_NON_IMPORTABLE_TYPES = frozenset({
    'one2many', 'many2many', 'many2one', 'reference', 'binary', 'serialized',
})


def _is_importable_field(line_fields: dict, field_name: str) -> bool:
    """Return True if field_name can be written back from Excel.

    Skips relational/binary fields and computed non-stored fields.
    """
    f = line_fields.get(field_name)
    if f is None:
        return False
    if f.compute and not f.store:
        return False
    return f.type not in _NON_IMPORTABLE_TYPES
