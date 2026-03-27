from __future__ import annotations

from odoo import fields
from odoo.exceptions import UserError


class ManufacturingBuilder:
    def __init__(self, env):
        self.env = env

    def build_boms_for_configuration(self, configuration):
        if not configuration.line_ids:
            raise UserError("Ajoutez des lignes de configuration avant de creer les nomenclatures.")
        created_boms = self.env["mrp.bom"]
        for line in configuration.line_ids:
            if line.bom_id:
                created_boms |= line.bom_id
                continue
            product_tmpl = line.modele_id.manufactured_product_tmpl_id or line.modele_id.sale_product_tmpl_id
            if not product_tmpl:
                continue
            bom = self.env["mrp.bom"].create(
                {
                    "product_tmpl_id": product_tmpl.id,
                    "product_qty": 1.0,
                    "code": f"{configuration.name}-{line.sequence}",
                    "joinery_configuration_id": configuration.id,
                    "joinery_configuration_line_id": line.id,
                    "project_id": configuration.project_project_id.id or False,
                    "type": "normal",
                    "ready_to_produce": "asap",
                }
            )
            line.bom_id = bom
            created_boms |= bom
            self._create_bom_components(bom, line)
        return created_boms

    def _create_bom_components(self, bom, line):
        unresolved_fillings = line.result_line_ids.filtered(lambda result: result.category == "filling" and not result.product_id)
        if unresolved_fillings:
            labels = ", ".join(
                f"{result.designation or result.ref_text} ({result.width_mm:g} x {result.height_mm:g} mm)"
                for result in unresolved_fillings
            )
            raise UserError(
                "Impossible de creer une nomenclature complete tant que les remplissages suivants "
                f"n'ont pas d'article lie: {labels}."
            )
        component_vals = []
        for result in line.result_line_ids:
            if not result.product_id:
                continue
            qty = result.get_manufacturing_quantity()
            if not qty:
                continue
            component_vals.append(
                {
                    "bom_id": bom.id,
                    "product_id": result.product_id.id,
                    "product_qty": qty,
                    "product_uom_id": result.product_id.uom_id.id,
                    "date_start": fields.Date.today(),
                }
            )
        if component_vals:
            self.env["mrp.bom.line"].create(component_vals)
