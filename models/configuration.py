import hashlib
from collections import defaultdict
from markupsafe import Markup, escape

from odoo import _, api, fields, models
from odoo.exceptions import AccessError, ValidationError

from ..services.engine import JoineryEngine
from ..services.manufacturing_builder import ManufacturingBuilder
from ..services.project_builder import ProjectBuilder
from ..services.quotation_builder import QuotationBuilder


class AluminiumJoineryConfiguration(models.Model):
    _name = "aluminium.joinery.configuration"
    _description = "Configuration de menuiserie"
    _inherit = ["mail.thread", "mail.activity.mixin"]
    _order = "id desc"

    name = fields.Char(string="Reference", default="Nouveau", copy=False, tracking=True)
    project_name = fields.Char(string="Nom du projet", tracking=True)
    partner_id = fields.Many2one("res.partner", string="Client", tracking=True)
    state = fields.Selection(
        [
            ("draft", "Brouillon"),
            ("calculated", "Calcule"),
            ("quoted", "Devis cree"),
            ("confirmed", "Confirme"),
            ("cancelled", "Annule"),
        ],
        string="Statut",
        default="draft",
        tracking=True,
    )
    sale_order_id = fields.Many2one("sale.order", string="Devis", readonly=True, copy=False)
    project_project_id = fields.Many2one("project.project", string="Projet", readonly=True, copy=False)
    bom_ids = fields.One2many("mrp.bom", "joinery_configuration_id", string="Nomenclatures", readonly=True)
    mrp_production_ids = fields.One2many("mrp.production", "joinery_configuration_id", string="Ordres de fabrication", readonly=True)
    currency_id = fields.Many2one(
        "res.currency",
        default=lambda self: self.env.company.currency_id.id,
        required=True,
    )
    company_id = fields.Many2one("res.company", string="Societe", default=lambda self: self.env.company.id, required=True)
    date = fields.Date(string="Date", default=fields.Date.context_today, required=True)
    note = fields.Text(string="Note")
    calculation_note = fields.Text(string="Journal de calcul", readonly=True)
    last_calculated_at = fields.Datetime(string="Dernier calcul", readonly=True)
    line_ids = fields.One2many("aluminium.joinery.configuration.line", "configuration_id", string="Lignes de configuration", copy=True)
    result_line_ids = fields.One2many("aluminium.joinery.result.line", "configuration_id", string="Resultats", readonly=True)
    summary_ids = fields.One2many("aluminium.joinery.material.summary", "configuration_id", string="Synthese matiere", readonly=True)
    results_structured_html = fields.Html(
        string="Vue structuree des resultats",
        compute="_compute_results_structured_html",
        sanitize=False,
        readonly=True,
    )
    summary_structured_html = fields.Html(
        string="Vue structuree de la synthese",
        compute="_compute_summary_structured_html",
        sanitize=False,
        readonly=True,
    )
    production_structured_html = fields.Html(
        string="Vue atelier",
        compute="_compute_production_structured_html",
        sanitize=False,
        readonly=True,
    )
    line_count = fields.Integer(string="Nombre de lignes", compute="_compute_counts")
    summary_count = fields.Integer(string="Nombre de syntheses", compute="_compute_counts")
    quote_count = fields.Integer(string="Nombre de devis", compute="_compute_counts")
    bom_count = fields.Integer(string="Nombre de nomenclatures", compute="_compute_counts")
    project_service_product_id = fields.Many2one(
        "product.product",
        domain=[("type", "=", "service")],
        string="Produit de service projet",
    )
    project_service_qty = fields.Float(string="Quantite service projet", default=1.0)

    @api.depends("line_ids", "summary_ids", "sale_order_id", "bom_ids")
    def _compute_counts(self):
        for rec in self:
            rec.line_count = len(rec.line_ids)
            rec.summary_count = len(rec.summary_ids)
            rec.quote_count = 1 if rec.sale_order_id else 0
            rec.bom_count = len(rec.bom_ids)

    @api.model_create_multi
    def create(self, vals_list):
        for vals in vals_list:
            if vals.get("name", "Nouveau") == "Nouveau":
                vals["name"] = self.env["ir.sequence"].next_by_code("aluminium.joinery.configuration") or "Nouveau"
        return super().create(vals_list)

    def action_calculate(self):
        for configuration in self:
            JoineryEngine(self.env).calculate_configuration(configuration)
        return True

    def action_generate_quotation(self):
        self.ensure_one()
        order = QuotationBuilder(self.env).build_for_configuration(self)
        return {
            "type": "ir.actions.act_window",
            "res_model": "sale.order",
            "view_mode": "form",
            "res_id": order.id,
        }

    def action_create_project(self):
        self.ensure_one()
        project = ProjectBuilder(self.env).build_for_configuration(self)
        return {
            "type": "ir.actions.act_window",
            "res_model": "project.project",
            "view_mode": "form",
            "res_id": project.id,
        }

    def action_create_boms(self):
        self.ensure_one()
        boms = ManufacturingBuilder(self.env).build_boms_for_configuration(self)
        action = self.env.ref("mrp.mrp_bom_form_action").read()[0]
        if len(boms) == 1:
            action["view_mode"] = "form"
            action["res_id"] = boms.id
        else:
            action["domain"] = [("id", "in", boms.ids)]
        return action

    def action_view_quote(self):
        self.ensure_one()
        if not self.sale_order_id:
            return False
        return {
            "type": "ir.actions.act_window",
            "res_model": "sale.order",
            "view_mode": "form",
            "res_id": self.sale_order_id.id,
        }

    def action_view_project(self):
        self.ensure_one()
        if not self.project_project_id:
            return False
        return {
            "type": "ir.actions.act_window",
            "res_model": "project.project",
            "view_mode": "form",
            "res_id": self.project_project_id.id,
        }

    def get_portal_url(self):
        self.ensure_one()
        return f"/my/configurateur/{self.id}"

    def get_portal_results_url(self):
        self.ensure_one()
        return f"/my/configurateur/{self.id}/resultats"

    def get_portal_material_summary_pdf_url(self):
        self.ensure_one()
        return f"/my/configurateur/{self.id}/synthese-matiere/pdf"

    def _get_portal_form_values(self):
        self.ensure_one()
        return {
            "configuration_id": self.id,
            "project_name": self.project_name or "",
            "lines": [
                {
                    "sequence": line.sequence,
                    "gamme_id": line.gamme_id.id,
                    "serie_id": line.serie_id.id,
                    "modele_id": line.modele_id.id,
                    "qty": int(line.qty) if line.qty else 1,
                    "width_mm": int(line.width_mm) if line.width_mm else False,
                    "height_mm": int(line.height_mm) if line.height_mm else False,
                }
                for line in self.line_ids.sorted(lambda rec: (rec.sequence, rec.id))
            ] or [
                {
                    "sequence": 10,
                    "gamme_id": False,
                    "serie_id": False,
                    "modele_id": False,
                    "qty": 1,
                    "width_mm": False,
                    "height_mm": False,
                }
            ],
        }

    @api.model
    def portal_upsert_single_line_configuration(self, partner, payload, configuration=None):
        return self.portal_upsert_configuration(partner, [payload], configuration=configuration, project_name=payload.get("project_name"))

    @api.model
    def portal_upsert_configuration(self, partner, line_payloads, configuration=None, project_name=None):
        partner = partner.sudo()
        Configuration = self.sudo()
        Line = self.env["aluminium.joinery.configuration.line"].sudo()
        line_vals_list = []
        if not line_payloads:
            raise ValidationError(_("Ajoutez au moins une ligne de configuration."))
        for index, payload in enumerate(line_payloads, start=1):
            gamme = self.env["aluminium.joinery.gamme"].sudo().search(
                [("id", "=", int(payload["gamme_id"])), ("active", "=", True)],
                limit=1,
            )
            serie = self.env["aluminium.joinery.serie"].sudo().search(
                [
                    ("id", "=", int(payload["serie_id"])),
                    ("gamme_id", "=", gamme.id),
                    ("active", "=", True),
                ],
                limit=1,
            )
            modele = self.env["aluminium.joinery.modele"].sudo().search(
                [
                    ("id", "=", int(payload["modele_id"])),
                    ("serie_id", "=", serie.id),
                    ("active", "=", True),
                ],
                limit=1,
            )
            if not gamme or not serie or not modele:
                raise ValidationError(_("Les selections de gamme, serie et modele sont invalides."))

            line_vals_list.append(
                {
                    "sequence": int(payload.get("sequence") or (index * 10)),
                    "gamme_id": gamme.id,
                    "serie_id": serie.id,
                    "modele_id": modele.id,
                    "qty": int(payload["qty"]),
                    "width_mm": float(payload["width_mm"]),
                    "height_mm": float(payload["height_mm"]),
                }
            )

        project_name = (project_name or "").strip()

        reusable = configuration and configuration.partner_id == partner and not configuration.sale_order_id and configuration.state in ("draft", "calculated")
        if configuration and configuration.partner_id != partner:
            raise AccessError(_("Vous ne pouvez pas modifier cette configuration."))

        if reusable:
            configuration = configuration.sudo()
            configuration.write(
                {
                    "partner_id": partner.id,
                    "project_name": project_name or False,
                }
            )
            configuration.line_ids.unlink()
        else:
            configuration = Configuration.create(
                {
                    "partner_id": partner.id,
                    "project_name": project_name or False,
                }
            )

        Line.create([{**line_vals, "configuration_id": configuration.id} for line_vals in line_vals_list])
        return configuration

    def portal_duplicate_for_partner(self, partner):
        self.ensure_one()
        partner = partner.sudo()
        if self.partner_id != partner:
            raise AccessError(_("Vous ne pouvez pas dupliquer cette configuration."))
        duplicated = self.copy(
            {
                "name": "Nouveau",
                "project_name": _("%s - copie") % (self.project_name or self.name),
                "partner_id": partner.id,
                "state": "draft",
                "sale_order_id": False,
                "project_project_id": False,
                "calculation_note": False,
                "last_calculated_at": False,
            }
        )
        duplicated.line_ids.write(
            {
                "state": "draft",
                "calculation_hash": False,
                "last_calculated_at": False,
                "calculation_message": False,
            }
        )
        return duplicated

    @api.depends(
        "line_ids.sequence",
        "line_ids.gamme_id",
        "line_ids.serie_id",
        "line_ids.modele_id",
        "line_ids.qty",
        "line_ids.width_mm",
        "line_ids.height_mm",
        "line_ids.result_line_ids",
        "line_ids.summary_ids",
    )
    def _compute_results_structured_html(self):
        for rec in self:
            rec.results_structured_html = rec._render_structured_blocks(mode="result")

    @api.depends(
        "line_ids.sequence",
        "line_ids.gamme_id",
        "line_ids.serie_id",
        "line_ids.modele_id",
        "line_ids.qty",
        "line_ids.width_mm",
        "line_ids.height_mm",
        "line_ids.summary_ids",
    )
    def _compute_summary_structured_html(self):
        for rec in self:
            rec.summary_structured_html = rec._render_structured_blocks(mode="summary")

    @api.depends(
        "line_ids.sequence",
        "line_ids.gamme_id",
        "line_ids.serie_id",
        "line_ids.modele_id",
        "line_ids.qty",
        "line_ids.width_mm",
        "line_ids.height_mm",
        "line_ids.result_line_ids",
        "line_ids.bom_id",
        "line_ids.bom_id.code",
        "line_ids.bom_id.product_tmpl_id",
    )
    def _compute_production_structured_html(self):
        for rec in self:
            rec.production_structured_html = rec._render_structured_blocks(mode="production")

    def _render_structured_blocks(self, mode):
        self.ensure_one()
        block_html = []
        for line in self.line_ids.sorted(lambda rec: (rec.sequence, rec.id)):
            datasets = line._get_structured_display_datasets(mode=mode)
            if not any(rows for _, _, rows in datasets):
                continue
            header = Markup(
                """
                <div class="o_ajc_block card mb-3">
                    <div class="card-header">
                        <strong>{label}</strong>
                    </div>
                    <div class="card-body">
                        <div class="row mb-3">
                            <div class="col-6 col-lg-2"><strong>Gamme</strong><br/>{gamme}</div>
                            <div class="col-6 col-lg-3"><strong>Serie</strong><br/>{serie}</div>
                            <div class="col-12 col-lg-3"><strong>Modele</strong><br/>{modele}</div>
                            <div class="col-4 col-lg-1"><strong>Qte</strong><br/>{qty}</div>
                            <div class="col-4 col-lg-1"><strong>Larg. (mm)</strong><br/>{width}</div>
                            <div class="col-4 col-lg-2"><strong>Haut. (mm)</strong><br/>{height}</div>
                        </div>
                        {extra}
                        {tables}
                    </div>
                </div>
                """
            ).format(
                label=escape(line.line_label or f"Ligne {line.sequence}"),
                gamme=escape(line.gamme_id.display_name or "-"),
                serie=escape(line.serie_id.display_name or "-"),
                modele=escape(line.modele_id.display_name or "-"),
                qty=escape(line._format_display_number(line.qty)),
                width=escape(line._format_display_number(line.width_mm)),
                height=escape(line._format_display_number(line.height_mm)),
                extra=line._render_structured_extra(mode),
                tables=Markup("").join(
                    line._render_structured_category_table(title, headers, rows) for title, headers, rows in datasets
                ),
            )
            block_html.append(header)
        if not block_html:
            return Markup(
                '<div class="alert alert-info">Aucun resultat calcule a afficher pour cette configuration.</div>'
            )
        return Markup("").join(block_html)

    def _get_report_groups(self, mode):
        self.ensure_one()
        groups = []
        groups_by_gamme = {}
        for line in self.line_ids.sorted(lambda rec: (rec.gamme_id.name or "", rec.sequence, rec.id)):
            sections = line._get_report_category_sections(mode=mode)
            sections = [section for section in sections if section["rows"]]
            if not sections:
                continue
            gamme_key = line.gamme_id.id or 0
            if gamme_key not in groups_by_gamme:
                group = {
                    "gamme_id": line.gamme_id.id,
                    "gamme_name": line.gamme_id.display_name or "Sans gamme",
                    "lines": [],
                }
                groups.append(group)
                groups_by_gamme[gamme_key] = group
            groups_by_gamme[gamme_key]["lines"].append(
                {
                    "sequence": line.sequence,
                    "label": line.line_label or f"Ligne {line.sequence}",
                    "serie_name": line.serie_id.display_name or "-",
                    "modele_name": line.modele_id.display_name or "-",
                    "qty": line._format_display_number(line.qty),
                    "width_mm": line._format_display_number(line.width_mm),
                    "height_mm": line._format_display_number(line.height_mm),
                    "sections": sections,
                }
            )
        return groups


class AluminiumJoineryConfigurationLine(models.Model):
    _name = "aluminium.joinery.configuration.line"
    _description = "Ligne de configuration"
    _order = "sequence, id"

    configuration_id = fields.Many2one("aluminium.joinery.configuration", string="Configuration", required=True, ondelete="cascade")
    sequence = fields.Integer(string="Sequence", default=10)
    gamme_id = fields.Many2one("aluminium.joinery.gamme", string="Gamme", required=True)
    serie_id = fields.Many2one("aluminium.joinery.serie", string="Serie", required=True)
    modele_id = fields.Many2one("aluminium.joinery.modele", string="Modele", required=True)
    qty = fields.Integer(string="Quantite", default=1, required=True)
    width_mm = fields.Float(string="Largeur (mm)", required=True)
    height_mm = fields.Float(string="Hauteur (mm)", required=True)
    state = fields.Selection(
        [
            ("draft", "Brouillon"),
            ("calculated", "Calcule"),
            ("error", "Erreur"),
        ],
        string="Statut",
        default="draft",
        readonly=True,
    )
    calculation_hash = fields.Char(string="Empreinte de calcul", readonly=True)
    last_calculated_at = fields.Datetime(string="Dernier calcul", readonly=True)
    calculation_message = fields.Char(string="Message de calcul", readonly=True)
    result_line_ids = fields.One2many("aluminium.joinery.result.line", "configuration_line_id", string="Resultats", readonly=True)
    summary_ids = fields.One2many("aluminium.joinery.material.summary", "configuration_line_id", string="Synthese", readonly=True)
    bom_id = fields.Many2one("mrp.bom", string="Nomenclature", readonly=True, copy=False)
    line_label = fields.Char(string="Libelle", compute="_compute_line_label")

    @api.depends("gamme_id", "serie_id", "modele_id", "qty", "width_mm", "height_mm")
    def _compute_line_label(self):
        for rec in self:
            parts = [p for p in [rec.gamme_id.name, rec.serie_id.name, rec.modele_id.name] if p]
            rec.line_label = " / ".join(parts) or "Ligne de configuration"

    @api.onchange("gamme_id")
    def _onchange_gamme_id(self):
        for rec in self:
            if rec.serie_id and rec.serie_id.gamme_id != rec.gamme_id:
                rec.serie_id = False
                rec.modele_id = False

    @api.onchange("serie_id")
    def _onchange_serie_id(self):
        for rec in self:
            if rec.modele_id and rec.modele_id.serie_id != rec.serie_id:
                rec.modele_id = False

    @api.constrains("serie_id", "gamme_id")
    def _check_serie_belongs_to_gamme(self):
        for rec in self:
            if rec.serie_id.gamme_id != rec.gamme_id:
                raise ValidationError("La serie selectionnee n'appartient pas a la gamme choisie.")

    @api.constrains("modele_id", "serie_id")
    def _check_modele_belongs_to_serie(self):
        for rec in self:
            if rec.modele_id.serie_id != rec.serie_id:
                raise ValidationError("Le modele selectionne n'appartient pas a la serie choisie.")

    @api.constrains("qty", "width_mm", "height_mm")
    def _check_input_ranges(self):
        for rec in self:
            if not 1 <= rec.qty <= 500:
                raise ValidationError("La quantite doit etre comprise entre 1 et 500.")
            if not 0.1 <= rec.width_mm <= 50000:
                raise ValidationError("La largeur doit etre comprise entre 0.1 et 50000 mm.")
            if not 0.1 <= rec.height_mm <= 50000:
                raise ValidationError("La hauteur doit etre comprise entre 0.1 et 50000 mm.")

    @api.model_create_multi
    def create(self, vals_list):
        records = super().create(vals_list)
        records._mark_dirty()
        return records

    def write(self, vals):
        tracked_fields = {"gamme_id", "serie_id", "modele_id", "qty", "width_mm", "height_mm"}
        res = super().write(vals)
        if tracked_fields.intersection(vals):
            self._mark_dirty()
        return res

    def _mark_dirty(self):
        self.write(
            {
                "state": "draft",
                "calculation_hash": False,
                "calculation_message": False,
            }
        )
        self.mapped("configuration_id").write({"state": "draft"})

    def _compute_calculation_hash_value(self):
        self.ensure_one()
        payload = f"{self.gamme_id.id}|{self.serie_id.id}|{self.modele_id.id}|{self.qty}|{self.width_mm}|{self.height_mm}"
        return hashlib.sha1(payload.encode("utf-8")).hexdigest()

    def _format_display_number(self, value):
        number = value or 0.0
        if abs(number - round(number)) < 1e-9:
            return str(int(round(number)))
        return f"{number:.2f}".rstrip("0").rstrip(".")

    def _get_product_preparation_metadata(self, product):
        self.ensure_one()
        if not product:
            return {
                "nature": "Article non lie",
                "action": "A definir",
                "bom_code": "-",
                "has_bom": False,
                "is_composite": False,
            }
        tmpl = product.product_tmpl_id
        bom = self.env["mrp.bom"].search(
            [("product_tmpl_id", "=", tmpl.id), ("type", "=", "normal")],
            order="id desc",
            limit=1,
        )
        is_composite = bool(tmpl.is_joinery_composite or tmpl.manufacturing_mode == "manufactured_composite")
        if product.is_placeholder_product:
            nature = "Article temporaire"
        elif is_composite:
            nature = "Sous-ensemble composite"
        else:
            nature = "Composant simple"
        if not product:
            action = "A definir"
        elif is_composite:
            action = "A fabriquer"
        elif tmpl.purchase_ok:
            action = "A sortir / approvisionner"
        else:
            action = "A sortir du stock"
        return {
            "nature": nature,
            "action": action,
            "bom_code": bom.code if bom else "-",
            "has_bom": bool(bom),
            "is_composite": is_composite,
        }

    def _get_structured_display_datasets(self, mode):
        self.ensure_one()
        if mode == "result":
            source = self.result_line_ids.sorted(lambda rec: (rec.category, rec.sequence, rec.id))
            data_by_category = {
                "profile": [
                    [
                        record.ref_text or "",
                        record.designation or "",
                        self._format_display_number(record.length_mm),
                        self._format_display_number(record.bar_length_mm),
                        self._format_display_number(record.bars_required),
                        self._format_display_number(record.billed_length_mm),
                        record.cut_type or "-",
                    ]
                    for record in source.filtered(lambda rec: rec.category == "profile")
                ],
                "accessoire": [
                    [
                        record.ref_text or "",
                        record.designation or "",
                        self._format_display_number(record.qty),
                    ]
                    for record in source.filtered(lambda rec: rec.category == "accessoire")
                ],
                "joint": [
                    [
                        record.ref_text or "",
                        record.designation or "",
                        self._format_display_number(record.length_mm),
                    ]
                    for record in source.filtered(lambda rec: rec.category == "joint")
                ],
                "filling": [
                    [
                        record.ref_text or "",
                        record.designation or "",
                        self._format_display_number(record.width_mm),
                        self._format_display_number(record.height_mm),
                        self._format_display_number(record.qty),
                    ]
                    for record in source.filtered(lambda rec: rec.category == "filling")
                ],
            }
            return [
                ("Profiles", ["Reference", "Designation", "Longueur requise (mm)", "Longueur standard (mm)", "Barres", "Longueur facturable (mm)", "Section / Coupe"], data_by_category["profile"]),
                ("Accessoires / Quincailleries", ["Reference", "Designation", "Qte"], data_by_category["accessoire"]),
                ("Joints", ["Reference", "Designation", "Longueur (mm)"], data_by_category["joint"]),
                ("Remplissage / Vitrages et panneaux", ["Reference", "Designation", "Largeur (mm)", "Hauteur (mm)", "Qte"], data_by_category["filling"]),
            ]

        if mode == "production":
            source = self.result_line_ids.sorted(lambda rec: (rec.category, rec.sequence, rec.id))
            profile_summary_map = {
                (summary.ref_text or "", summary.designation or ""): summary
                for summary in self.summary_ids.filtered(lambda rec: rec.category == "profile")
            }

            profile_data = defaultdict(
                lambda: {
                    "ref_text": "",
                    "designation": "",
                    "pieces": 0.0,
                    "unit_length_mm": 0.0,
                    "total_length_mm": 0.0,
                    "bar_length_mm": 0.0,
                    "billed_length_mm": 0.0,
                    "cut_type": "-",
                    "bars_required": 0,
                    "nature": "",
                    "action": "",
                    "bom_code": "-",
                }
            )
            for record in source.filtered(lambda rec: rec.category == "profile"):
                meta = self._get_product_preparation_metadata(record.product_id)
                key = (
                    record.ref_text or "",
                    record.designation or "",
                    record.length_mm or 0.0,
                    record.cut_type or "-",
                    meta["nature"],
                    meta["action"],
                    meta["bom_code"],
                )
                row = profile_data[key]
                row["ref_text"] = record.ref_text or ""
                row["designation"] = record.designation or ""
                row["pieces"] += record.qty or 1.0
                row["unit_length_mm"] = record.length_mm or 0.0
                row["total_length_mm"] += (record.length_mm or 0.0) * (record.qty or 1.0)
                row["bar_length_mm"] = max(row["bar_length_mm"], record.bar_length_mm or 0.0)
                row["billed_length_mm"] += record.billed_length_mm or 0.0
                row["cut_type"] = record.cut_type or "-"
                row["nature"] = meta["nature"]
                row["action"] = meta["action"]
                row["bom_code"] = meta["bom_code"]
                summary = profile_summary_map.get((record.ref_text or "", record.designation or ""))
                row["bars_required"] = max(row["bars_required"], summary.bars_required if summary else 0)
            profile_rows = [
                [
                    row["ref_text"],
                    row["designation"],
                    self._format_display_number(row["pieces"]),
                    self._format_display_number(row["unit_length_mm"]),
                    self._format_display_number(row["total_length_mm"]),
                    self._format_display_number(row["bar_length_mm"]),
                    row["cut_type"],
                    self._format_display_number(row["bars_required"]) if row["bars_required"] else "-",
                    self._format_display_number(row["billed_length_mm"]),
                    row["nature"],
                    row["action"],
                    row["bom_code"],
                ]
                for row in sorted(
                    profile_data.values(),
                    key=lambda item: (
                        item["nature"],
                        item["action"],
                        item["ref_text"],
                        item["designation"],
                        item["unit_length_mm"],
                    ),
                )
            ]

            accessory_data = defaultdict(
                lambda: {
                    "ref_text": "",
                    "designation": "",
                    "qty": 0.0,
                    "nature": "",
                    "action": "",
                    "bom_code": "-",
                }
            )
            for record in source.filtered(lambda rec: rec.category == "accessoire"):
                meta = self._get_product_preparation_metadata(record.product_id)
                key = (
                    record.ref_text or "",
                    record.designation or "",
                    meta["nature"],
                    meta["action"],
                    meta["bom_code"],
                )
                row = accessory_data[key]
                row["ref_text"] = record.ref_text or ""
                row["designation"] = record.designation or ""
                row["qty"] += record.qty or 0.0
                row["nature"] = meta["nature"]
                row["action"] = meta["action"]
                row["bom_code"] = meta["bom_code"]
            accessory_rows = [
                [
                    row["ref_text"],
                    row["designation"],
                    self._format_display_number(row["qty"]),
                    row["nature"],
                    row["action"],
                    row["bom_code"],
                ]
                for row in sorted(
                    accessory_data.values(),
                    key=lambda item: (item["nature"], item["action"], item["ref_text"], item["designation"]),
                )
            ]

            joint_data = defaultdict(
                lambda: {
                    "ref_text": "",
                    "designation": "",
                    "total_length_mm": 0.0,
                    "nature": "",
                    "action": "",
                    "bom_code": "-",
                }
            )
            for record in source.filtered(lambda rec: rec.category == "joint"):
                meta = self._get_product_preparation_metadata(record.product_id)
                key = (
                    record.ref_text or "",
                    record.designation or "",
                    meta["nature"],
                    meta["action"],
                    meta["bom_code"],
                )
                row = joint_data[key]
                row["ref_text"] = record.ref_text or ""
                row["designation"] = record.designation or ""
                row["total_length_mm"] += record.length_mm or 0.0
                row["nature"] = meta["nature"]
                row["action"] = meta["action"]
                row["bom_code"] = meta["bom_code"]
            joint_rows = [
                [
                    row["ref_text"],
                    row["designation"],
                    self._format_display_number(row["total_length_mm"]),
                    row["nature"],
                    row["action"],
                    row["bom_code"],
                ]
                for row in sorted(
                    joint_data.values(),
                    key=lambda item: (item["nature"], item["action"], item["ref_text"], item["designation"]),
                )
            ]

            filling_data = defaultdict(
                lambda: {
                    "ref_text": "",
                    "designation": "",
                    "width_mm": 0.0,
                    "height_mm": 0.0,
                    "qty": 0.0,
                    "nature": "",
                    "action": "",
                    "bom_code": "-",
                }
            )
            for record in source.filtered(lambda rec: rec.category == "filling"):
                meta = self._get_product_preparation_metadata(record.product_id)
                key = (
                    record.ref_text or "",
                    record.designation or "",
                    record.width_mm or 0.0,
                    record.height_mm or 0.0,
                    meta["nature"],
                    meta["action"],
                    meta["bom_code"],
                )
                row = filling_data[key]
                row["ref_text"] = record.ref_text or ""
                row["designation"] = record.designation or ""
                row["width_mm"] = record.width_mm or 0.0
                row["height_mm"] = record.height_mm or 0.0
                row["qty"] += record.qty or 0.0
                row["nature"] = meta["nature"]
                row["action"] = meta["action"]
                row["bom_code"] = meta["bom_code"]
            filling_rows = [
                [
                    row["ref_text"],
                    row["designation"],
                    self._format_display_number(row["width_mm"]),
                    self._format_display_number(row["height_mm"]),
                    self._format_display_number(row["qty"]),
                    row["nature"],
                    row["action"],
                    row["bom_code"],
                ]
                for row in sorted(
                    filling_data.values(),
                    key=lambda item: (
                        item["nature"],
                        item["action"],
                        item["ref_text"],
                        item["designation"],
                        item["width_mm"],
                        item["height_mm"],
                    ),
                )
            ]
            return [
                (
                    "Pieces / Profiles a couper",
                    ["Reference", "Designation", "Pieces", "Long. req. unit. (mm)", "Long. req. totale (mm)", "Long. standard (mm)", "Section / Coupe", "Barres", "Long. facturable (mm)", "Nature", "Action", "Nomenclature"],
                    profile_rows,
                ),
                (
                    "Quincaillerie a sortir",
                    ["Reference", "Designation", "Qte", "Nature", "Action", "Nomenclature"],
                    accessory_rows,
                ),
                (
                    "Joints a preparer",
                    ["Reference", "Designation", "Longueur totale (mm)", "Nature", "Action", "Nomenclature"],
                    joint_rows,
                ),
                (
                    "Vitrages / Panneaux",
                    ["Reference", "Designation", "Largeur (mm)", "Hauteur (mm)", "Qte", "Nature", "Action", "Nomenclature"],
                    filling_rows,
                ),
            ]

        source = self.summary_ids.sorted(lambda rec: (rec.category, rec.ref_text or "", rec.id))
        data_by_category = {
            "profile": [
                [
                    record.ref_text or "",
                    record.designation or "",
                    self._format_display_number(record.total_qty),
                    self._format_display_number(record.total_length_mm / 1000.0),
                    self._format_display_number(record.bars_required),
                ]
                for record in source.filtered(lambda rec: rec.category == "profile")
            ],
            "accessoire": [
                [
                    record.ref_text or "",
                    record.designation or "",
                    self._format_display_number(record.total_qty),
                ]
                for record in source.filtered(lambda rec: rec.category == "accessoire")
            ],
            "joint": [
                [
                    record.ref_text or "",
                    record.designation or "",
                    self._format_display_number(record.total_length_mm / 1000.0),
                ]
                for record in source.filtered(lambda rec: rec.category == "joint")
            ],
            "filling": [
                [
                    record.ref_text or "",
                    record.designation or "",
                    self._format_display_number(record.width_mm),
                    self._format_display_number(record.height_mm),
                    self._format_display_number(record.total_qty),
                ]
                for record in source.filtered(lambda rec: rec.category == "filling")
            ],
        }
        return [
            ("Profiles", ["Reference", "Designation", "Qte", "Longueur totale (m)", "Barres necessaires"], data_by_category["profile"]),
            ("Accessoires / Quincailleries", ["Reference", "Designation", "Qte"], data_by_category["accessoire"]),
            ("Joints", ["Reference", "Designation", "Longueur totale (m)"], data_by_category["joint"]),
            ("Remplissage / Vitrages et panneaux", ["Reference", "Designation", "Largeur (mm)", "Hauteur (mm)", "Qte"], data_by_category["filling"]),
        ]

    def _render_structured_category_table(self, title, headers, rows):
        title_markup = Markup('<h5 class="mt-3 mb-2">{}</h5>').format(escape(title))
        if not rows:
            return title_markup + Markup(
                '<div class="text-muted mb-3"><em>Aucune ligne pour cette categorie.</em></div>'
            )
        header_cells = Markup("").join(
            Markup("<th>{}</th>").format(escape(header)) for header in headers
        )
        body_rows = []
        for row in rows:
            body_rows.append(
                Markup("<tr>{}</tr>").format(
                    Markup("").join(Markup("<td>{}</td>").format(escape(value or "")) for value in row)
                )
            )
        return title_markup + Markup(
            """
            <table class="table table-sm table-striped table-bordered mb-3">
                <thead><tr>{headers}</tr></thead>
                <tbody>{rows}</tbody>
            </table>
            """
        ).format(
            headers=header_cells,
            rows=Markup("").join(body_rows),
        )

    def _render_structured_extra(self, mode):
        self.ensure_one()
        if mode != "production":
            return Markup("")
        bom = self.bom_id
        composite_count = len(
            self.result_line_ids.filtered(
                lambda result: result.product_id
                and (
                    result.product_id.product_tmpl_id.is_joinery_composite
                    or result.product_id.product_tmpl_id.manufacturing_mode == "manufactured_composite"
                )
            ).mapped("product_id")
        )
        placeholder_count = len(
            self.result_line_ids.filtered(lambda result: result.product_id and result.product_id.is_placeholder_product).mapped("product_id")
        )
        if not bom:
            return Markup(
                """
                <div class="alert alert-warning mb-3">
                    Nomenclature non generee pour cette ligne. Utilisez "Creer les nomenclatures" si vous voulez preparer la production.
                    <br/>Sous-ensembles composites detectes : {composites}
                    <br/>Articles temporaires restants : {placeholders}
                </div>
                """
            ).format(
                composites=escape(self._format_display_number(composite_count)),
                placeholders=escape(self._format_display_number(placeholder_count)),
            )
        return Markup(
            """
            <div class="mb-3">
                <table class="table table-sm table-bordered mb-0">
                    <tbody>
                        <tr>
                            <td><strong>Nomenclature</strong></td>
                            <td>{bom_code}</td>
                            <td><strong>Produit fabrique</strong></td>
                            <td>{product}</td>
                        </tr>
                        <tr>
                            <td><strong>Sous-ensembles composites</strong></td>
                            <td>{composites}</td>
                            <td><strong>Articles temporaires</strong></td>
                            <td>{placeholders}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            """
        ).format(
            bom_code=escape(bom.code or "-"),
            product=escape(bom.product_tmpl_id.display_name or "-"),
            composites=escape(self._format_display_number(composite_count)),
            placeholders=escape(self._format_display_number(placeholder_count)),
        )

    def _make_report_row(self, *cells):
        return [{"value": value or "", "class": css_class} for value, css_class in cells]

    def _get_report_category_sections(self, mode):
        self.ensure_one()
        if mode == "result":
            source = self.result_line_ids.sorted(lambda rec: (rec.category, rec.sequence, rec.id))
            profile_rows = [
                self._make_report_row(
                    (record.ref_text, ""),
                    (record.designation, ""),
                    (self._format_display_number(record.length_mm), "text-end"),
                    (self._format_display_number(record.bar_length_mm), "text-end"),
                    (self._format_display_number(record.bars_required), "text-end"),
                    (self._format_display_number(record.billed_length_mm), "text-end"),
                    (record.cut_type or "-", ""),
                )
                for record in source.filtered(lambda rec: rec.category == "profile")
            ]
            accessory_rows = [
                self._make_report_row(
                    (record.ref_text, ""),
                    (record.designation, ""),
                    (self._format_display_number(record.qty), "text-end"),
                )
                for record in source.filtered(lambda rec: rec.category == "accessoire")
            ]
            joint_rows = [
                self._make_report_row(
                    (record.ref_text, ""),
                    (record.designation, ""),
                    (self._format_display_number(record.length_mm), "text-end"),
                )
                for record in source.filtered(lambda rec: rec.category == "joint")
            ]
            filling_rows = [
                self._make_report_row(
                    (record.ref_text, ""),
                    (record.designation, ""),
                    (self._format_display_number(record.width_mm), "text-end"),
                    (self._format_display_number(record.height_mm), "text-end"),
                    (self._format_display_number(record.qty), "text-end"),
                )
                for record in source.filtered(lambda rec: rec.category == "filling")
            ]
            return [
                {
                    "title": "PROFILES",
                    "headers": ["Reference", "Designation", "Longueur requise (mm)", "Longueur standard (mm)", "Barres", "Longueur facturable (mm)", "Section / Coupe"],
                    "rows": profile_rows,
                    "totals": self._make_report_row(
                        ("", ""),
                        ("Total profiles", "fw-bold"),
                        (self._format_display_number(sum((rec.length_mm or 0.0) for rec in source.filtered(lambda rec: rec.category == "profile"))), "text-end fw-bold"),
                        ("", ""),
                        (self._format_display_number(sum((rec.bars_required or 0.0) for rec in source.filtered(lambda rec: rec.category == "profile"))), "text-end fw-bold"),
                        (self._format_display_number(sum((rec.billed_length_mm or 0.0) for rec in source.filtered(lambda rec: rec.category == "profile"))), "text-end fw-bold"),
                        ("", ""),
                    ) if profile_rows else [],
                },
                {
                    "title": "ACCESSOIRES",
                    "headers": ["Reference", "Designation", "Qte"],
                    "rows": accessory_rows,
                    "totals": self._make_report_row(("", ""), ("Total accessoires", "fw-bold"), (self._format_display_number(sum((rec.qty or 0.0) for rec in source.filtered(lambda rec: rec.category == "accessoire"))), "text-end fw-bold")) if accessory_rows else [],
                },
                {
                    "title": "JOINTS",
                    "headers": ["Reference", "Designation", "Longueur (mm)"],
                    "rows": joint_rows,
                    "totals": self._make_report_row(("", ""), ("Total joints", "fw-bold"), (self._format_display_number(sum((rec.length_mm or 0.0) for rec in source.filtered(lambda rec: rec.category == "joint"))), "text-end fw-bold")) if joint_rows else [],
                },
                {
                    "title": "REMPLISSAGES",
                    "headers": ["Reference", "Designation", "Largeur (mm)", "Hauteur (mm)", "Qte"],
                    "rows": filling_rows,
                    "totals": self._make_report_row(
                        ("", ""),
                        ("Total remplissages", "fw-bold"),
                        ("", ""),
                        ("", ""),
                        (self._format_display_number(sum((rec.qty or 0.0) for rec in source.filtered(lambda rec: rec.category == "filling"))), "text-end fw-bold"),
                    ) if filling_rows else [],
                },
            ]

        source = self.summary_ids.sorted(lambda rec: (rec.category, rec.ref_text or "", rec.id))
        profile_rows = [
            self._make_report_row(
                (record.ref_text, ""),
                (record.designation, ""),
                (self._format_display_number(record.total_length_mm / 1000.0), "text-end"),
                (self._format_display_number(record.bar_length_mm), "text-end"),
                (self._format_display_number(record.bars_required), "text-end"),
                (self._format_display_number(record.billed_length_mm / 1000.0), "text-end"),
            )
            for record in source.filtered(lambda rec: rec.category == "profile")
        ]
        accessory_rows = [
            self._make_report_row(
                (record.ref_text, ""),
                (record.designation, ""),
                (self._format_display_number(record.total_qty), "text-end"),
            )
            for record in source.filtered(lambda rec: rec.category == "accessoire")
        ]
        joint_rows = [
            self._make_report_row(
                (record.ref_text, ""),
                (record.designation, ""),
                (self._format_display_number(record.total_length_mm / 1000.0), "text-end"),
            )
            for record in source.filtered(lambda rec: rec.category == "joint")
        ]
        filling_rows = [
            self._make_report_row(
                (record.ref_text, ""),
                (record.designation, ""),
                (self._format_display_number(record.width_mm), "text-end"),
                (self._format_display_number(record.height_mm), "text-end"),
                (self._format_display_number(record.total_qty), "text-end"),
            )
            for record in source.filtered(lambda rec: rec.category == "filling")
        ]
        return [
            {
                "title": "PROFILES",
                "headers": ["Reference", "Designation", "Longueur requise (m)", "Longueur standard (mm)", "Barres necessaires", "Longueur facturable (m)"],
                "rows": profile_rows,
                "totals": self._make_report_row(
                    ("", ""),
                    ("Total profiles", "fw-bold"),
                    (self._format_display_number(sum((rec.total_length_mm or 0.0) for rec in source.filtered(lambda rec: rec.category == "profile")) / 1000.0), "text-end fw-bold"),
                    ("", ""),
                    (self._format_display_number(sum((rec.bars_required or 0.0) for rec in source.filtered(lambda rec: rec.category == "profile"))), "text-end fw-bold"),
                    (self._format_display_number(sum((rec.billed_length_mm or 0.0) for rec in source.filtered(lambda rec: rec.category == "profile")) / 1000.0), "text-end fw-bold"),
                ) if profile_rows else [],
            },
            {
                "title": "ACCESSOIRES",
                "headers": ["Reference", "Designation", "Qte"],
                "rows": accessory_rows,
                "totals": self._make_report_row(
                    ("", ""),
                    ("Total accessoires", "fw-bold"),
                    (self._format_display_number(sum((rec.total_qty or 0.0) for rec in source.filtered(lambda rec: rec.category == "accessoire"))), "text-end fw-bold"),
                ) if accessory_rows else [],
            },
            {
                "title": "JOINTS",
                "headers": ["Reference", "Designation", "Longueur totale (m)"],
                "rows": joint_rows,
                "totals": self._make_report_row(
                    ("", ""),
                    ("Total joints", "fw-bold"),
                    (self._format_display_number(sum((rec.total_length_mm or 0.0) for rec in source.filtered(lambda rec: rec.category == "joint")) / 1000.0), "text-end fw-bold"),
                ) if joint_rows else [],
            },
            {
                "title": "REMPLISSAGES",
                "headers": ["Reference", "Designation", "Largeur (mm)", "Hauteur (mm)", "Qte"],
                "rows": filling_rows,
                "totals": self._make_report_row(
                    ("", ""),
                    ("Total remplissages", "fw-bold"),
                    ("", ""),
                    ("", ""),
                    (self._format_display_number(sum((rec.total_qty or 0.0) for rec in source.filtered(lambda rec: rec.category == "filling"))), "text-end fw-bold"),
                ) if filling_rows else [],
            },
        ]


class AluminiumJoineryResultLine(models.Model):
    _name = "aluminium.joinery.result.line"
    _description = "Ligne de resultat"
    _order = "configuration_line_id, sequence, id"

    configuration_id = fields.Many2one("aluminium.joinery.configuration", string="Configuration", required=True, ondelete="cascade")
    configuration_line_id = fields.Many2one("aluminium.joinery.configuration.line", string="Ligne de configuration", required=True, ondelete="cascade")
    gamme_id = fields.Many2one("aluminium.joinery.gamme", string="Gamme", readonly=True)
    serie_id = fields.Many2one("aluminium.joinery.serie", string="Serie", readonly=True)
    modele_id = fields.Many2one("aluminium.joinery.modele", string="Modele", readonly=True)
    category = fields.Selection(
        [
            ("profile", "Profile"),
            ("accessoire", "Accessoire"),
            ("joint", "Joint"),
            ("filling", "Remplissage"),
        ],
        string="Categorie",
        required=True,
    )
    sequence = fields.Integer(string="Sequence", default=10)
    product_id = fields.Many2one("product.product", string="Article", ondelete="restrict")
    ref_text = fields.Char(string="Reference")
    designation = fields.Char(string="Designation")
    qty = fields.Float(string="Quantite", default=0.0)
    length_mm = fields.Float(string="Longueur (mm)")
    billed_length_mm = fields.Float(string="Longueur facturable (mm)")
    width_mm = fields.Float(string="Largeur (mm)")
    height_mm = fields.Float(string="Hauteur (mm)")
    cut_type = fields.Char(string="Type de coupe")
    bar_length_mm = fields.Float(string="Longueur barre (mm)")
    bars_required = fields.Integer(string="Barres necessaires")
    unit_price = fields.Float(string="Prix unitaire")
    computed_json = fields.Json(string="Valeurs calculees")

    def _compute_surface_m2(self):
        self.ensure_one()
        return (self.qty or 0.0) * (self.width_mm or 0.0) * (self.height_mm or 0.0) / 1000000.0

    def get_manufacturing_quantity(self):
        self.ensure_one()
        if self.product_id:
            if self.category == "profile":
                if self.product_id.uom_id == self.env.ref("uom.product_uom_meter"):
                    return (self.length_mm or 0.0) / 1000.0
                return self.qty or 0.0
            if self.category == "joint" and self.product_id.uom_id == self.env.ref("uom.product_uom_meter"):
                return (self.length_mm or 0.0) / 1000.0
            if self.category == "filling" and self.product_id.uom_id == self.env.ref("uom.product_uom_square_meter"):
                return self._compute_surface_m2()
        if self.category in ("profile", "joint"):
            return self.qty or 0.0
        return self.qty or self.length_mm or 0.0


class AluminiumJoineryMaterialSummary(models.Model):
    _name = "aluminium.joinery.material.summary"
    _description = "Synthese matiere"
    _order = "configuration_line_id, category, ref_text"

    configuration_id = fields.Many2one("aluminium.joinery.configuration", string="Configuration", required=True, ondelete="cascade")
    configuration_line_id = fields.Many2one(
        "aluminium.joinery.configuration.line",
        string="Ligne de configuration",
        ondelete="cascade",
    )
    gamme_id = fields.Many2one("aluminium.joinery.gamme", string="Gamme")
    category = fields.Selection(
        [
            ("profile", "Profile"),
            ("accessoire", "Accessoire"),
            ("joint", "Joint"),
            ("filling", "Remplissage"),
        ]
    )
    product_id = fields.Many2one("product.product", string="Article", ondelete="restrict")
    ref_text = fields.Char(string="Reference")
    designation = fields.Char(string="Designation")
    total_qty = fields.Float(string="Quantite totale")
    total_length_mm = fields.Float(string="Longueur totale (mm)")
    billed_length_mm = fields.Float(string="Longueur facturable (mm)")
    bar_length_mm = fields.Float(string="Longueur barre (mm)")
    bars_required = fields.Integer(string="Barres necessaires")
    width_mm = fields.Float(string="Largeur (mm)")
    height_mm = fields.Float(string="Hauteur (mm)")
    unit_price = fields.Float(string="Prix unitaire")
    total_price = fields.Monetary(string="Montant total", currency_field="currency_id")
    currency_id = fields.Many2one(related="configuration_id.currency_id", store=True, readonly=True)

    def _compute_surface_m2(self):
        self.ensure_one()
        return (self.total_qty or 0.0) * (self.width_mm or 0.0) * (self.height_mm or 0.0) / 1000000.0

    def get_sale_quantity(self):
        self.ensure_one()
        if self.category == "profile":
            if self.product_id and self.product_id.uom_id == self.env.ref("uom.product_uom_meter"):
                return (self.billed_length_mm or self.total_length_mm) / 1000.0
            return self.bars_required or self.total_qty
        if self.category == "joint":
            if self.product_id and self.product_id.uom_id == self.env.ref("uom.product_uom_meter"):
                return self.total_length_mm / 1000.0
            return self.total_qty
        if self.category == "filling":
            if self.product_id and self.product_id.uom_id == self.env.ref("uom.product_uom_square_meter"):
                return self._compute_surface_m2()
        return self.total_qty
