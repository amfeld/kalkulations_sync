from odoo import fields, models


class ResCompany(models.Model):
    _inherit = 'res.company'

    amf_kalksync_template = fields.Binary(
        string='Calculation template (.xlsx)',
        attachment=True,
    )
    amf_kalksync_template_name = fields.Char(string='Template file name')
    amf_kalksync_default_product_id = fields.Many2one(
        'product.product',
        string='Default product for new lines',
        help='Fallback product used when a new line without a product reference '
             'is created via "N" or a row copy.',
    )
