from odoo import fields, models


class ResConfigSettings(models.TransientModel):
    _inherit = 'res.config.settings'

    amf_kalksync_template = fields.Binary(
        related='company_id.amf_kalksync_template',
        readonly=False,
        string='Calculation template (.xlsx)',
    )
    amf_kalksync_template_name = fields.Char(
        related='company_id.amf_kalksync_template_name',
        readonly=False,
        string='Template file name',
    )
    amf_kalksync_default_product_id = fields.Many2one(
        related='company_id.amf_kalksync_default_product_id',
        readonly=False,
        string='Default product for new lines',
    )
