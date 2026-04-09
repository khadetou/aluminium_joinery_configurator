from __future__ import annotations

from collections import defaultdict

from odoo import fields
from odoo.exceptions import UserError


CATEGORY_LABELS = {
    "profile": "Profiles",
    "accessoire": "Accessoires",
    "joint": "Joints",
    "filling": "Remplissage",
}


class QuotationBuilder:
    def __init__(self, env):
        self.env = env

    def _build_sale_line_description(self, summary):
        description = summary.designation or summary.product_id.display_name
        if summary.category == "profile" and summary.bar_length_mm:
            required_m = (summary.total_length_mm or 0.0) / 1000.0
            standard_m = (summary.bar_length_mm or 0.0) / 1000.0
            billed_m = (summary.billed_length_mm or 0.0) / 1000.0
            description = (
                f"{description}\n"
                f"Besoin: {required_m:g} m | Palette: {standard_m:g} m | "
                f"Barres: {summary.bars_required:g} | Facturable: {billed_m:g} m"
            )
        return description

    def build_for_configuration(self, configuration):
        if not configuration.summary_ids:
            raise UserError("Calculez la configuration avant de generer un devis.")
        order = configuration.sale_order_id or self.env["sale.order"].create(
            {
                "partner_id": configuration.partner_id.id,
                "date_order": fields.Datetime.now(),
                "joinery_configuration_id": configuration.id,
                "origin": configuration.name,
            }
        )
        order.order_line.filtered(
            lambda line: line.joinery_configuration_id == configuration
        ).unlink()

        sequence = 10
        grouped = defaultdict(list)
        for summary in configuration.summary_ids.sorted(lambda r: (r.category, r.id)):
            grouped[summary.category].append(summary)

        line_values = []
        for category in ("profile", "accessoire", "joint", "filling"):
            summaries = grouped.get(category)
            if not summaries:
                continue
            category_lines = []
            for summary in summaries:
                if summary.product_id:
                    qty = summary.get_sale_quantity()
                    if not qty:
                        continue
                    category_lines.append(
                        {
                            "order_id": order.id,
                            "sequence": sequence + len(category_lines) + 1,
                            "product_id": summary.product_id.id,
                            "product_uom_id": summary.product_id.uom_id.id,
                            "product_uom_qty": qty,
                            "name": self._build_sale_line_description(summary),
                            "joinery_configuration_id": configuration.id,
                            "joinery_summary_id": summary.id,
                        }
                    )
                    continue
                if category == "filling":
                    category_lines.append(
                        {
                            "order_id": order.id,
                            "sequence": sequence + len(category_lines) + 1,
                            "display_type": "line_note",
                            "name": (
                                f"{summary.designation or 'Remplissage'} - "
                                f"{summary.total_qty:g} x {summary.width_mm:g} x {summary.height_mm:g} mm "
                                "(article de remplissage non lie)"
                            ),
                            "joinery_configuration_id": configuration.id,
                            "joinery_summary_id": summary.id,
                        }
                    )
            if not category_lines:
                continue
            line_values.append(
                {
                    "order_id": order.id,
                    "sequence": sequence,
                    "display_type": "line_section",
                    "name": CATEGORY_LABELS[category],
                    "joinery_configuration_id": configuration.id,
                }
            )
            sequence += 1
            line_values.extend(category_lines)
            sequence += len(category_lines)

        if configuration.project_service_product_id:
            line_values.append(
                {
                    "order_id": order.id,
                    "sequence": sequence,
                    "display_type": "line_section",
                    "name": "Services projet",
                    "joinery_configuration_id": configuration.id,
                }
            )
            sequence += 1
            line_values.append(
                {
                    "order_id": order.id,
                    "sequence": sequence,
                    "product_id": configuration.project_service_product_id.id,
                    "product_uom_id": configuration.project_service_product_id.uom_id.id,
                    "product_uom_qty": configuration.project_service_qty or 1.0,
                    "name": configuration.project_service_product_id.display_name,
                    "joinery_configuration_id": configuration.id,
                }
            )

        if not line_values:
            raise UserError("Aucune ligne de devis n'a pu etre creee a partir de la synthese calculee.")

        self.env["sale.order.line"].create(line_values)
        configuration.write({"sale_order_id": order.id, "state": "quoted"})
        return order
