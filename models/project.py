from odoo import fields, models


class ProjectProject(models.Model):
    _inherit = "project.project"

    joinery_configuration_id = fields.Many2one(
        "aluminium.joinery.configuration", string="Configuration menuiserie", copy=False
    )
