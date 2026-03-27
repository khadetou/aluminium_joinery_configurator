from odoo import fields, models


class SaleOrder(models.Model):
    _inherit = "sale.order"

    joinery_configuration_id = fields.Many2one(
        "aluminium.joinery.configuration", string="Configuration menuiserie", copy=False
    )


class SaleOrderLine(models.Model):
    _inherit = "sale.order.line"

    joinery_configuration_id = fields.Many2one(
        "aluminium.joinery.configuration", string="Configuration menuiserie", copy=False
    )
    joinery_summary_id = fields.Many2one(
        "aluminium.joinery.material.summary", string="Synthese menuiserie", copy=False
    )
