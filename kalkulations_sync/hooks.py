"""Install/upgrade hooks for kalkulations_sync."""

import base64
import logging

from odoo import tools

_logger = logging.getLogger(__name__)

# Bundled default template shipped with the module.
_DEFAULT_TEMPLATE_PATH = 'kalkulations_sync/static/templates/kalkulation_template.xlsx'
_DEFAULT_TEMPLATE_NAME = 'Vorlage_Kalkulation.xlsx'


def _seed_default_template(env):
    """Pre-fill the company Excel template from the bundled default.

    Only companies without a template get seeded — never overwrite a template
    a customer has already uploaded.
    """
    path = tools.misc.file_path(_DEFAULT_TEMPLATE_PATH)
    with open(path, 'rb') as fh:
        data = base64.b64encode(fh.read())

    companies = env['res.company'].search([('amf_kalksync_template', '=', False)])
    if not companies:
        return
    companies.write({
        'amf_kalksync_template': data,
        'amf_kalksync_template_name': _DEFAULT_TEMPLATE_NAME,
    })
    _logger.info(
        "kalkulations_sync: seeded default template into %d company(ies)",
        len(companies),
    )


def post_init_hook(env):
    """Seed the default template on fresh install."""
    _seed_default_template(env)
