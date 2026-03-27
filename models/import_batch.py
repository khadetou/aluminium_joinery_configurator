from __future__ import annotations

import base64
from collections import defaultdict

from odoo import _, fields, models
from odoo.exceptions import UserError

from ..services.xlsx_tools import read_xlsx_sheets


MODULE_NAME = "aluminium_joinery_configurator"


class AluminiumJoineryCatalogImport(models.Model):
    _name = "aluminium.joinery.catalog.import"
    _description = "Import de catalogue menuiserie"
    _order = "id desc"

    name = fields.Char(required=True)
    file = fields.Binary(required=True, attachment=True)
    filename = fields.Char()
    state = fields.Selection(
        [
            ("draft", "Brouillon"),
            ("processed", "Traite"),
            ("failed", "Echec"),
        ],
        default="draft",
    )
    sync_missing_children = fields.Boolean(
        string="Synchroniser les enfants absents",
        default=True,
        help="Supprime les regles, remplissages, modeles et lignes de BOM importes precedemment mais absents du nouveau classeur.",
    )
    import_log = fields.Text(string="Journal d'import", readonly=True)
    imported_gamme_ids = fields.Many2many(
        "aluminium.joinery.gamme",
        relation="ajc_catalog_import_gamme_rel",
        column1="import_id",
        column2="gamme_id",
    )
    imported_serie_ids = fields.Many2many(
        "aluminium.joinery.serie",
        relation="ajc_catalog_import_serie_rel",
        column1="import_id",
        column2="serie_id",
    )
    imported_modele_ids = fields.Many2many(
        "aluminium.joinery.modele",
        relation="ajc_catalog_import_modele_rel",
        column1="import_id",
        column2="modele_id",
    )
    imported_rule_ids = fields.Many2many(
        "aluminium.joinery.rule",
        relation="ajc_catalog_import_rule_rel",
        column1="import_id",
        column2="rule_id",
    )
    imported_filling_rule_ids = fields.Many2many(
        "aluminium.joinery.filling.rule",
        relation="ajc_catalog_import_fill_rule_rel",
        column1="import_id",
        column2="filling_rule_id",
    )
    imported_product_tmpl_ids = fields.Many2many(
        "product.template",
        relation="ajc_catalog_import_product_rel",
        column1="import_id",
        column2="product_tmpl_id",
    )
    imported_bom_ids = fields.Many2many(
        "mrp.bom",
        relation="ajc_catalog_import_bom_rel",
        column1="import_id",
        column2="bom_id",
    )

    def action_process_import(self):
        for rec in self:
            if not rec.file:
                raise UserError(_("Ajoutez un fichier xlsx a importer."))
            try:
                sheets = read_xlsx_sheets(base64.b64decode(rec.file))
                context = rec._process_workbook(sheets)
            except Exception as exc:  # pragma: no cover - surfaced in UI
                rec.write(
                    {
                        "state": "failed",
                        "import_log": f"{rec.import_log or ''}\nERREUR: {exc}".strip(),
                    }
                )
                raise

            rec.write(
                {
                    "state": "processed",
                    "import_log": "\n".join(context["log"]),
                    "imported_gamme_ids": [(6, 0, context["gammes"].ids)],
                    "imported_serie_ids": [(6, 0, context["series"].ids)],
                    "imported_modele_ids": [(6, 0, context["modeles"].ids)],
                    "imported_rule_ids": [(6, 0, context["rules"].ids)],
                    "imported_filling_rule_ids": [(6, 0, context["filling_rules"].ids)],
                    "imported_product_tmpl_ids": [(6, 0, context["products"].ids)],
                    "imported_bom_ids": [(6, 0, context["boms"].ids)],
                }
            )
        return True

    def _process_workbook(self, sheets: dict[str, list[dict[str, str]]]) -> dict[str, models.Model]:
        required = ["Gammes", "Articles", "Series Complete"]
        missing = [name for name in required if name not in sheets]
        if missing:
            raise UserError(_("Feuilles manquantes dans le classeur: %s") % ", ".join(missing))

        context = {
            "log": [],
            "gammes": self.env["aluminium.joinery.gamme"],
            "series": self.env["aluminium.joinery.serie"],
            "modeles": self.env["aluminium.joinery.modele"],
            "rules": self.env["aluminium.joinery.rule"],
            "filling_rules": self.env["aluminium.joinery.filling.rule"],
            "products": self.env["product.template"],
            "boms": self.env["mrp.bom"],
        }
        self._import_gammes(sheets.get("Gammes", []), context)
        self._import_articles(sheets.get("Articles", []), context)
        self._import_series_complete(sheets.get("Series Complete", []), context)
        if sheets.get("BOMs Complete"):
            self._import_boms_complete(sheets["BOMs Complete"], context)
        context["log"].append(
            "Import termine: "
            f"{len(context['gammes'])} gammes, "
            f"{len(context['series'])} series, "
            f"{len(context['modeles'])} modeles, "
            f"{len(context['rules'])} regles, "
            f"{len(context['filling_rules'])} remplissages, "
            f"{len(context['products'])} articles, "
            f"{len(context['boms'])} BOMs."
        )
        return context

    def _import_gammes(self, rows, context):
        for row in rows:
            if not any(row.values()):
                continue
            vals = {
                "code": row.get("code"),
                "name": row.get("name") or row.get("code"),
                "default_bar_length_mm": self._to_float(row.get("default_bar_length_mm"), default=5800.0),
                "active": self._to_bool(row.get("active"), default=True),
            }
            gamme = self._upsert_record(
                "aluminium.joinery.gamme",
                row.get("id"),
                [("code", "=", vals["code"])],
                vals,
            )
            context["gammes"] |= gamme
        context["log"].append(f"Gammes importees: {len(context['gammes'])}")

    def _import_articles(self, rows, context):
        for row in rows:
            if not any(row.values()):
                continue
            vals = {
                "default_code": row.get("default_code"),
                "name": row.get("name") or row.get("default_code"),
                "type": row.get("type") or "consu",
                "uom_id": self._resolve_m2o_xmlid("uom.uom", row.get("uom_id/id")).id,
                "uom_po_id": self._resolve_m2o_xmlid("uom.uom", row.get("uom_po_id/id")).id,
                "sale_ok": self._to_bool(row.get("sale_ok"), default=True),
                "purchase_ok": self._to_bool(row.get("purchase_ok"), default=True),
                "joinery_item_type": row.get("joinery_item_type") or False,
                "joinery_bar_length_mm": self._to_float(row.get("joinery_bar_length_mm"), default=0.0),
                "joinery_usage_role": row.get("joinery_usage_role") or False,
                "is_joinery_composite": self._to_bool(row.get("is_joinery_composite"), default=False),
                "is_placeholder_product_tmpl": self._to_bool(row.get("is_placeholder_product_tmpl"), default=False),
                "manufacturing_mode": row.get("manufacturing_mode") or "standard",
            }
            template = self._upsert_record(
                "product.template",
                row.get("id"),
                [("default_code", "=", vals["default_code"])],
                vals,
            )
            context["products"] |= template
        context["log"].append(f"Articles importes: {len(context['products'])}")

    def _import_series_complete(self, rows, context):
        structures = self._group_series_rows(rows)
        for serie_data in structures:
            gamme = self._resolve_m2o_xmlid("aluminium.joinery.gamme", serie_data["gamme_id/id"])
            serie_vals = {
                "gamme_id": gamme.id,
                "code": serie_data["code"],
                "name": serie_data["name"] or serie_data["code"],
                "opening_type": serie_data.get("opening_type") or False,
                "active": self._to_bool(serie_data.get("active"), default=True),
            }
            serie = self._upsert_record(
                "aluminium.joinery.serie",
                serie_data.get("id"),
                [("code", "=", serie_vals["code"]), ("gamme_id", "=", gamme.id)],
                serie_vals,
            )
            context["series"] |= serie
            imported_model_xids = set()
            for model_data in serie_data["models"]:
                model_vals = {
                    "serie_id": serie.id,
                    "code": model_data["code"],
                    "name": model_data["name"] or model_data["code"],
                    "panel_count": self._to_int(model_data.get("panel_count")),
                    "rail_count": self._to_int(model_data.get("rail_count")),
                    "active": self._to_bool(model_data.get("active"), default=True),
                    "x_import_key": model_data.get("x_import_key") or False,
                    "sale_product_tmpl_id": self._resolve_optional_xmlid("product.template", model_data.get("sale_product_tmpl_id/id")).id,
                    "manufactured_product_tmpl_id": self._resolve_optional_xmlid("product.template", model_data.get("manufactured_product_tmpl_id/id")).id,
                    "project_template_id": self._resolve_optional_xmlid("project.project", model_data.get("project_template_id/id")).id,
                }
                modele = self._upsert_record(
                    "aluminium.joinery.modele",
                    model_data.get("id"),
                    [("code", "=", model_vals["code"]), ("serie_id", "=", serie.id)],
                    model_vals,
                )
                context["modeles"] |= modele
                if model_data.get("id"):
                    imported_model_xids.add(self._full_xmlid(model_data["id"]))

                imported_rule_xids = set()
                for rule_data in model_data["rules"]:
                    product = self._resolve_product_for_rule(rule_data)
                    rule_vals = {
                        "modele_id": modele.id,
                        "sequence": self._to_int(rule_data.get("sequence"), default=10),
                        "name": rule_data.get("name") or rule_data.get("ref_text") or _("Regle"),
                        "category": rule_data.get("category"),
                        "product_id": product.product_variant_id.id if product else False,
                        "ref_text": rule_data.get("ref_text") or (product.default_code if product else False),
                        "designation_override": rule_data.get("designation_override") or False,
                        "profile_role": rule_data.get("profile_role") or False,
                        "cut_type": rule_data.get("cut_type") or False,
                        "formula_family": rule_data.get("formula_family") or False,
                        "target_measure": rule_data.get("target_measure") or "qty",
                        "multiplier": self._to_float(rule_data.get("multiplier"), default=1.0),
                        "coef_l": self._to_float(rule_data.get("coef_l")),
                        "offset_l": self._to_float(rule_data.get("offset_l")),
                        "divisor_l": self._to_float(rule_data.get("divisor_l"), default=1.0),
                        "coef_h": self._to_float(rule_data.get("coef_h")),
                        "offset_h": self._to_float(rule_data.get("offset_h")),
                        "divisor_h": self._to_float(rule_data.get("divisor_h"), default=1.0),
                        "constant": self._to_float(rule_data.get("constant")),
                        "uom_kind": rule_data.get("uom_kind") or "unit",
                        "active": self._to_bool(rule_data.get("active"), default=True),
                        "rule_code": rule_data.get("rule_code") or False,
                        "product_resolution_mode": rule_data.get("product_resolution_mode") or "direct_product",
                        "rule_kind": "formula",
                    }
                    rule = self._upsert_record(
                        "aluminium.joinery.rule",
                        rule_data.get("id"),
                        [("modele_id", "=", modele.id), ("rule_code", "=", rule_vals["rule_code"] or rule_vals["name"])],
                        rule_vals,
                    )
                    context["rules"] |= rule
                    if rule_data.get("id"):
                        imported_rule_xids.add(self._full_xmlid(rule_data["id"]))

                imported_filling_xids = set()
                for filling_data in model_data["fillings"]:
                    filling_vals = {
                        "modele_id": modele.id,
                        "sequence": self._to_int(filling_data.get("sequence"), default=10),
                        "name": filling_data.get("name") or _("Remplissage"),
                        "family_width": filling_data.get("family_width") or "fill_dim",
                        "width_coef_l": self._to_float(filling_data.get("width_coef_l"), default=0.0),
                        "width_coef_h": self._to_float(filling_data.get("width_coef_h"), default=0.0),
                        "width_constant": self._to_float(filling_data.get("width_constant"), default=0.0),
                        "family_height": filling_data.get("family_height") or "fill_dim",
                        "height_coef_l": self._to_float(filling_data.get("height_coef_l"), default=0.0),
                        "height_coef_h": self._to_float(filling_data.get("height_coef_h"), default=0.0),
                        "height_constant": self._to_float(filling_data.get("height_constant"), default=0.0),
                        "family_qty": filling_data.get("family_qty") or "qty_only",
                        "qty_multiplier": self._to_float(filling_data.get("qty_multiplier"), default=1.0),
                        "qty_constant": self._to_float(filling_data.get("qty_constant"), default=0.0),
                        "active": self._to_bool(filling_data.get("active"), default=True),
                        "rule_code": filling_data.get("rule_code") or False,
                    }
                    filling = self._upsert_record(
                        "aluminium.joinery.filling.rule",
                        filling_data.get("id"),
                        [("modele_id", "=", modele.id), ("rule_code", "=", filling_vals["rule_code"] or filling_vals["name"])],
                        filling_vals,
                    )
                    context["filling_rules"] |= filling
                    if filling_data.get("id"):
                        imported_filling_xids.add(self._full_xmlid(filling_data["id"]))

                if self.sync_missing_children:
                    self._sync_missing_children(modele, "aluminium.joinery.rule", modele.rule_ids, imported_rule_xids)
                    self._sync_missing_children(
                        modele,
                        "aluminium.joinery.filling.rule",
                        modele.filling_rule_ids,
                        imported_filling_xids,
                    )
            if self.sync_missing_children:
                self._sync_missing_children(serie, "aluminium.joinery.modele", serie.modele_ids, imported_model_xids)
        context["log"].append(
            f"Catalogue importe: {len(context['series'])} series, {len(context['modeles'])} modeles, "
            f"{len(context['rules'])} regles, {len(context['filling_rules'])} remplissages."
        )

    def _import_boms_complete(self, rows, context):
        structures = self._group_bom_rows(rows)
        for bom_data in structures:
            product_tmpl = self._resolve_optional_xmlid("product.template", bom_data.get("product_tmpl_id/id"))
            if not product_tmpl and bom_data.get("product_default_code"):
                product_tmpl = self._resolve_product_template(bom_data.get("product_default_code"))
            bom_vals = {
                "product_tmpl_id": product_tmpl.id,
                "product_default_code": bom_data.get("product_default_code") or product_tmpl.default_code,
                "type": bom_data.get("type") or "normal",
                "product_qty": self._to_float(bom_data.get("product_qty"), default=1.0),
                "product_uom_id": self._resolve_m2o_xmlid("uom.uom", bom_data.get("product_uom_id/id")).id,
                "code": bom_data.get("code") or False,
            }
            bom = self._upsert_record(
                "mrp.bom",
                bom_data.get("id"),
                [("product_tmpl_id", "=", product_tmpl.id), ("code", "=", bom_vals["code"])],
                bom_vals,
            )
            context["boms"] |= bom
            imported_line_xids = set()
            for line_data in bom_data["lines"]:
                product = self._resolve_product_template(
                    line_data.get("product_default_code") or line_data.get("component_default_code")
                ).product_variant_id
                line_vals = {
                    "bom_id": bom.id,
                    "product_id": product.id,
                    "component_default_code": line_data.get("product_default_code") or line_data.get("component_default_code"),
                    "product_qty": self._to_float(line_data.get("product_qty"), default=1.0),
                    "product_uom_id": self._resolve_m2o_xmlid("uom.uom", line_data.get("product_uom_id/id")).id,
                    "sequence": self._to_int(line_data.get("sequence"), default=10),
                }
                bom_line = self._upsert_record(
                    "mrp.bom.line",
                    line_data.get("id"),
                    [("bom_id", "=", bom.id), ("product_id", "=", product.id), ("sequence", "=", line_vals["sequence"])],
                    line_vals,
                )
                if line_data.get("id"):
                    imported_line_xids.add(self._full_xmlid(line_data["id"]))
            if self.sync_missing_children:
                self._sync_missing_children(bom, "mrp.bom.line", bom.bom_line_ids, imported_line_xids)
        context["log"].append(f"BOMs importees: {len(context['boms'])}")

    def _group_series_rows(self, rows):
        structures = []
        current_serie = None
        current_model = None
        serie_fields = ["id", "gamme_id/id", "code", "name", "opening_type", "active"]
        model_fields = [
            "modele_ids/id",
            "modele_ids/code",
            "modele_ids/name",
            "modele_ids/panel_count",
            "modele_ids/rail_count",
            "modele_ids/active",
            "modele_ids/x_import_key",
            "modele_ids/sale_product_tmpl_id/id",
            "modele_ids/manufactured_product_tmpl_id/id",
            "modele_ids/project_template_id/id",
        ]
        rule_prefix = "modele_ids/rule_ids/"
        filling_prefix = "modele_ids/filling_rule_ids/"

        for row in rows:
            if not any(row.values()):
                continue
            if any(row.get(field) for field in serie_fields):
                current_serie = {field: row.get(field, "") for field in serie_fields}
                current_serie["models"] = []
                structures.append(current_serie)
                current_model = None
            if current_serie is None:
                raise UserError(_("Une ligne enfant Series Complete apparait avant toute serie parente."))
            if any(row.get(field) for field in model_fields):
                current_model = {
                    "id": row.get("modele_ids/id"),
                    "code": row.get("modele_ids/code"),
                    "name": row.get("modele_ids/name"),
                    "panel_count": row.get("modele_ids/panel_count"),
                    "rail_count": row.get("modele_ids/rail_count"),
                    "active": row.get("modele_ids/active"),
                    "x_import_key": row.get("modele_ids/x_import_key"),
                    "sale_product_tmpl_id/id": row.get("modele_ids/sale_product_tmpl_id/id"),
                    "manufactured_product_tmpl_id/id": row.get("modele_ids/manufactured_product_tmpl_id/id"),
                    "project_template_id/id": row.get("modele_ids/project_template_id/id"),
                    "rules": [],
                    "fillings": [],
                }
                current_serie["models"].append(current_model)
            if current_model is None:
                raise UserError(_("Une ligne enfant de modele apparait avant toute definition de modele."))
            rule_data = {key[len(rule_prefix):]: value for key, value in row.items() if key.startswith(rule_prefix) and value}
            if rule_data:
                current_model["rules"].append(rule_data)
            filling_data = {
                key[len(filling_prefix):]: value for key, value in row.items() if key.startswith(filling_prefix) and value
            }
            if filling_data:
                current_model["fillings"].append(filling_data)
        return structures

    def _group_bom_rows(self, rows):
        structures = []
        current = None
        bom_fields = ["id", "product_tmpl_id/id", "product_default_code", "type", "product_qty", "product_uom_id/id", "code"]
        line_prefix = "bom_line_ids/"
        for row in rows:
            if not any(row.values()):
                continue
            if any(row.get(field) for field in bom_fields):
                current = {field: row.get(field, "") for field in bom_fields}
                current["lines"] = []
                structures.append(current)
            if current is None:
                raise UserError(_("Une ligne enfant de BOM apparait avant toute BOM parente."))
            line_data = {key[len(line_prefix):]: value for key, value in row.items() if key.startswith(line_prefix) and value}
            if line_data:
                current["lines"].append(line_data)
        return structures

    def _resolve_product_for_rule(self, rule_data):
        product_xmlid = rule_data.get("product_id/id")
        if product_xmlid:
            product = self._resolve_optional_xmlid("product.product", product_xmlid)
            if product:
                return product.product_tmpl_id
        code = rule_data.get("product_default_code") or rule_data.get("product_id/default_code")
        if not code:
            return False
        return self._resolve_product_template(code)

    def _resolve_product_template(self, default_code):
        code = (default_code or "").strip()
        if not code:
            return self.env["product.template"]
        template = self.env["product.template"].search([("default_code", "=", code)], limit=1)
        if not template:
            raise UserError(_("Aucun article ne correspond a la reference '%s'.") % code)
        return template

    def _upsert_record(self, model_name, xmlid, domain, vals):
        record = self._resolve_optional_xmlid(model_name, xmlid)
        if not record and domain:
            record = self.env[model_name].search(domain, limit=1)
        clean_vals = {}
        for key, value in vals.items():
            if value is None:
                continue
            clean_vals[key] = False if value == "" else value
        if record:
            record.write(clean_vals)
        else:
            record = self.env[model_name].create(clean_vals)
        if xmlid:
            self._ensure_xmlid(model_name, record.id, xmlid)
        return record

    def _sync_missing_children(self, parent, model_name, records, imported_xids):
        if not imported_xids:
            return
        current = {
            xmlid: rec_id
            for rec_id, xmlid in self._get_xmlids(records, model_name).items()
            if xmlid.startswith(f"{MODULE_NAME}.")
        }
        missing_ids = [rec_id for xmlid, rec_id in current.items() if xmlid not in imported_xids]
        if missing_ids:
            self.env[model_name].browse(missing_ids).unlink()

    def _resolve_m2o_xmlid(self, model_name, xmlid):
        record = self._resolve_optional_xmlid(model_name, xmlid)
        if not record:
            raise UserError(_("Reference externe introuvable sur %s: %s") % (model_name, xmlid))
        return record

    def _resolve_optional_xmlid(self, model_name, xmlid):
        if not xmlid:
            return self.env[model_name]
        full = self._full_xmlid(xmlid)
        try:
            record = self.env.ref(full)
        except ValueError:
            return self.env[model_name]
        if record._name != model_name:
            raise UserError(_("La reference %s ne pointe pas vers %s.") % (full, model_name))
        return record

    def _ensure_xmlid(self, model_name, res_id, xmlid):
        full = self._full_xmlid(xmlid)
        module, name = full.split(".", 1)
        data = self.env["ir.model.data"].sudo().search(
            [("module", "=", module), ("name", "=", name)],
            limit=1,
        )
        vals = {"module": module, "name": name, "model": model_name, "res_id": res_id, "noupdate": True}
        if data:
            data.write(vals)
        else:
            self.env["ir.model.data"].sudo().create(vals)

    def _full_xmlid(self, xmlid):
        return xmlid if "." in xmlid else f"{MODULE_NAME}.{xmlid}"

    def _get_xmlids(self, records, model_name):
        if not records:
            return {}
        self.env.cr.execute(
            """
            SELECT res_id, module || '.' || name
            FROM ir_model_data
            WHERE model = %s AND res_id = ANY(%s)
            """,
            [model_name, list(records.ids)],
        )
        return dict(self.env.cr.fetchall())

    def _to_bool(self, value, default=False):
        if value in (None, ""):
            return default
        if isinstance(value, bool):
            return value
        return str(value).strip().lower() in {"1", "true", "yes", "y", "oui"}

    def _to_float(self, value, default=0.0):
        if value in (None, ""):
            return default
        return float(value)

    def _to_int(self, value, default=False):
        if value in (None, ""):
            return default
        return int(float(value))
