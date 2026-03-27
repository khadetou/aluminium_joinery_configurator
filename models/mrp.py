from odoo import fields, models


class MrpBom(models.Model):
    _inherit = "mrp.bom"

    joinery_configuration_id = fields.Many2one(
        "aluminium.joinery.configuration", string="Configuration menuiserie", copy=False
    )
    joinery_configuration_line_id = fields.Many2one(
        "aluminium.joinery.configuration.line", string="Ligne de configuration", copy=False
    )


class MrpProduction(models.Model):
    _inherit = "mrp.production"

    joinery_configuration_id = fields.Many2one(
        "aluminium.joinery.configuration", string="Configuration menuiserie", copy=False
    )
    joinery_configuration_line_id = fields.Many2one(
        "aluminium.joinery.configuration.line", string="Ligne de configuration", copy=False
    )
