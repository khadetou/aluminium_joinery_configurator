from __future__ import annotations

import math
from collections import defaultdict

from odoo import fields
from odoo.exceptions import UserError

from .formula import FormulaError, safe_eval_formula


class JoineryEngine:
    def __init__(self, env):
        self.env = env

    def calculate_configuration(self, configuration):
        configuration.line_ids.mapped("result_line_ids").unlink()
        configuration.summary_ids.unlink()
        errors = []
        for line in configuration.line_ids.sorted("sequence"):
            try:
                self.calculate_line(line)
            except Exception as exc:  # pragma: no cover - defensive surface for UI actions
                line.write(
                    {
                        "state": "error",
                        "calculation_message": str(exc),
                        "last_calculated_at": fields.Datetime.now(),
                    }
                )
                errors.append(f"{line.display_name}: {exc}")
        self.aggregate_materials(configuration)
        if errors:
            configuration.write(
                {
                    "state": "draft",
                    "calculation_note": "\n".join(errors),
                }
            )
            raise UserError("\n".join(errors))
        configuration.write(
            {
                "state": "calculated",
                "calculation_note": False,
                "last_calculated_at": fields.Datetime.now(),
            }
        )
        return True

    def calculate_line(self, line):
        rules = line.modele_id.rule_ids.filtered("active").sorted(lambda rec: (rec.sequence, rec.id))
        filling_rules = line.modele_id.filling_rule_ids.filtered("active").sorted(lambda rec: (rec.sequence, rec.id))
        line.result_line_ids.unlink()
        if not rules and not filling_rules:
            raise UserError(
                f"Aucune regle de calcul active n'a ete trouvee pour le modele '{line.modele_id.display_name}'."
            )
        vals_list = []
        for rule in rules:
            vals = self._build_result_vals(line, rule)
            if vals:
                vals_list.append(vals)
        for filling_rule in filling_rules:
            vals = self._build_filling_result_vals(line, filling_rule)
            if vals:
                vals_list.append(vals)
        if vals_list:
            self.env["aluminium.joinery.result.line"].create(vals_list)
        line.write(
            {
                "state": "calculated",
                "calculation_hash": line._compute_calculation_hash_value(),
                "last_calculated_at": fields.Datetime.now(),
                "calculation_message": False,
            }
        )
        return vals_list

    def aggregate_materials(self, configuration):
        grouped = defaultdict(
            lambda: {
                "configuration_line_id": False,
                "total_qty": 0.0,
                "total_length_mm": 0.0,
                "width_mm": 0.0,
                "height_mm": 0.0,
                "designation": False,
                "ref_text": False,
                "product_id": False,
                "gamme_id": False,
                "category": False,
                "bar_length_mm": 0.0,
                "unit_price": 0.0,
            }
        )
        for result in configuration.result_line_ids:
            key = (
                result.configuration_line_id.id or 0,
                result.category,
                result.gamme_id.id or 0,
                result.product_id.id or 0,
                result.ref_text or "",
            )
            data = grouped[key]
            data["configuration_line_id"] = result.configuration_line_id.id
            data["designation"] = result.designation
            data["ref_text"] = result.ref_text
            data["product_id"] = result.product_id.id
            data["gamme_id"] = result.gamme_id.id
            data["category"] = result.category
            data["total_qty"] += result.qty or 0.0
            data["total_length_mm"] += result.length_mm or 0.0
            data["width_mm"] = max(data["width_mm"], result.width_mm or 0.0)
            data["height_mm"] = max(data["height_mm"], result.height_mm or 0.0)
            data["bar_length_mm"] = result.bar_length_mm or data["bar_length_mm"]
            data["unit_price"] = result.unit_price or data["unit_price"]

        summary_vals = []
        for data in grouped.values():
            bar_length = data["bar_length_mm"] or 0.0
            bars_required = math.ceil(data["total_length_mm"] / bar_length) if bar_length and data["total_length_mm"] else 0
            pricing_qty = bars_required or data["total_qty"] or 0.0
            summary_vals.append(
                {
                    "configuration_id": configuration.id,
                    "configuration_line_id": data["configuration_line_id"] or False,
                    "gamme_id": data["gamme_id"] or False,
                    "product_id": data["product_id"] or False,
                    "category": data["category"],
                    "ref_text": data["ref_text"],
                    "designation": data["designation"],
                    "total_qty": data["total_qty"],
                    "total_length_mm": data["total_length_mm"],
                    "bar_length_mm": bar_length,
                    "bars_required": bars_required,
                    "width_mm": data["width_mm"],
                    "height_mm": data["height_mm"],
                    "unit_price": data["unit_price"],
                    "total_price": pricing_qty * data["unit_price"],
                }
            )
        if summary_vals:
            self.env["aluminium.joinery.material.summary"].create(summary_vals)
        return True

    def _build_result_vals(self, line, rule):
        variables = self._get_variables(line)
        designation = rule.designation_override or rule.product_id.display_name or rule.name
        ref_text = rule.product_id.default_code or rule.ref_text or rule.name
        common = {
            "configuration_id": line.configuration_id.id,
            "configuration_line_id": line.id,
            "gamme_id": line.gamme_id.id,
            "serie_id": line.serie_id.id,
            "modele_id": line.modele_id.id,
            "category": rule.category,
            "sequence": rule.sequence,
            "product_id": rule.product_id.id,
            "ref_text": ref_text,
            "designation": designation,
            "cut_type": rule.cut_type,
            "bar_length_mm": self._get_bar_length_mm(rule.product_id, line, rule.category),
            "unit_price": rule.product_id.lst_price if rule.product_id else 0.0,
            "computed_json": {
                "rule_id": rule.id,
                "formula_family": rule.formula_family,
                "expression_value": rule.expression_value,
                "product_resolution_mode": rule.product_resolution_mode,
            },
        }
        if rule.category == "filling":
            qty = self._compute_formula(rule.qty_expression or "Q", variables)
            width = self._compute_formula(rule.width_expression or "L", variables)
            height = self._compute_formula(rule.height_expression or "H", variables)
            return {
                **common,
                "qty": self._apply_rounding(qty, rule.rounding_mode),
                "width_mm": self._apply_rounding(width, rule.rounding_mode),
                "height_mm": self._apply_rounding(height, rule.rounding_mode),
            }
        value = self._apply_rounding(self._compute_rule_value(rule, variables), rule.rounding_mode)
        if rule.category in ("profile", "joint"):
            return {
                **common,
                "qty": 1.0,
                "length_mm": value,
            }
        return {
            **common,
            "qty": value,
        }

    def _build_filling_result_vals(self, line, rule):
        variables = self._get_variables(line)
        product = rule.product_id
        width = self._apply_rounding(
            self._compute_named_family(
                rule.family_width,
                variables,
                coef_l=rule.width_coef_l,
                coef_h=rule.width_coef_h,
                constant=rule.width_constant,
                apply_quantity=False,
            ),
            "none",
        )
        height = self._apply_rounding(
            self._compute_named_family(
                rule.family_height,
                variables,
                coef_l=rule.height_coef_l,
                coef_h=rule.height_coef_h,
                constant=rule.height_constant,
                apply_quantity=False,
            ),
            "none",
        )
        qty = self._apply_rounding(
            self._compute_named_family(
                rule.family_qty,
                variables,
                multiplier=rule.qty_multiplier,
                constant=rule.qty_constant,
                apply_quantity=True,
            ),
            "none",
        )
        return {
            "configuration_id": line.configuration_id.id,
            "configuration_line_id": line.id,
            "gamme_id": line.gamme_id.id,
            "serie_id": line.serie_id.id,
            "modele_id": line.modele_id.id,
            "category": "filling",
            "sequence": rule.sequence,
            "product_id": product.id or False,
            "ref_text": product.default_code or rule.product_default_code or rule.rule_code or rule.name,
            "designation": product.display_name or rule.name,
            "qty": qty,
            "width_mm": width,
            "height_mm": height,
            "unit_price": product.lst_price if product else 0.0,
            "computed_json": {
                "filling_rule_id": rule.id,
                "family_width": rule.family_width,
                "family_height": rule.family_height,
                "family_qty": rule.family_qty,
                "product_default_code": rule.product_default_code,
            },
        }

    def _compute_rule_value(self, rule, variables):
        if rule.formula_family:
            return self._compute_named_family(
                rule.formula_family,
                variables,
                multiplier=rule.multiplier,
                coef_l=rule.coef_l,
                offset_l=rule.offset_l,
                divisor_l=rule.divisor_l,
                coef_h=rule.coef_h,
                offset_h=rule.offset_h,
                divisor_h=rule.divisor_h,
                constant=rule.constant,
                apply_quantity=True,
            )
        if rule.expression_value:
            return self._compute_formula(rule.expression_value, variables)
        base_value = self._get_base_value(rule, variables)
        if rule.operator == "add":
            base_value += rule.fixed_offset_mm
        elif rule.operator == "subtract":
            base_value -= rule.fixed_offset_mm
        value = base_value * (1.0 if rule.multiplier is None else rule.multiplier)
        if rule.apply_quantity:
            value *= variables["Q"]
        if rule.divisor:
            value /= rule.divisor
        return value

    def _compute_named_family(
        self,
        family,
        variables,
        multiplier=1.0,
        coef_l=0.0,
        offset_l=0.0,
        divisor_l=1.0,
        coef_h=0.0,
        offset_h=0.0,
        divisor_h=1.0,
        constant=0.0,
        apply_quantity=True,
    ):
        l_term = self._safe_division((coef_l * variables["L"]) + offset_l, divisor_l)
        h_term = self._safe_division((coef_h * variables["H"]) + offset_h, divisor_h)
        q_value = variables["Q"] if apply_quantity else 1.0
        factor = 1.0 if multiplier is None else multiplier

        if family == "qty_only":
            return (q_value * factor) + constant
        if family == "linear_l":
            return (q_value * factor * l_term) + constant
        if family == "linear_h":
            return (q_value * factor * h_term) + constant
        if family in ("sum_h_l", "perimeter"):
            return (q_value * factor * (l_term + h_term)) + constant
        if family in ("joint_combo", "generic_affine"):
            return (q_value * factor * ((coef_l * variables["L"]) + (coef_h * variables["H"]) + constant))
        if family == "fill_dim":
            return (coef_l * variables["L"]) + (coef_h * variables["H"]) + constant
        raise UserError(f"Famille de formule non supportee: {family}")

    def _get_base_value(self, rule, variables):
        mapping = {
            "none": 1.0,
            "width": variables["L"],
            "height": variables["H"],
            "sum_both": variables["L"] + variables["H"],
            "perimeter": 2.0 * (variables["L"] + variables["H"]),
        }
        return mapping.get(rule.base_dimension or "none", 1.0)

    def _compute_formula(self, expression, variables):
        try:
            return safe_eval_formula(expression, variables)
        except FormulaError as exc:
            raise UserError(str(exc)) from exc

    def _safe_division(self, value, divisor):
        if not divisor:
            raise UserError("Un diviseur nul a ete rencontre dans une regle de calcul.")
        return value / divisor

    def _apply_rounding(self, value, rounding_mode):
        if rounding_mode == "up":
            return math.ceil(value)
        if rounding_mode == "half_up":
            return math.floor(value + 0.5)
        return value

    def _get_bar_length_mm(self, product, line, category=None):
        if category == "joint":
            template = product.product_tmpl_id if product else False
            return (
                (template.joinery_bar_length_mm or template.schuller_bar_length_mm or template.ajc_bar_length_mm)
                if template
                else 0.0
            )
        if not product:
            return line.gamme_id.default_bar_length_mm
        template = product.product_tmpl_id
        return (
            template.joinery_bar_length_mm
            or template.schuller_bar_length_mm
            or template.ajc_bar_length_mm
            or line.gamme_id.default_bar_length_mm
        )

    def _get_variables(self, line):
        return {
            "Q": float(line.qty or 0.0),
            "L": float(line.width_mm or 0.0),
            "H": float(line.height_mm or 0.0),
        }
