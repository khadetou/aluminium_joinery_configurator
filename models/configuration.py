import hashlib

from odoo import api, fields, models
from odoo.exceptions import ValidationError

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
    width_mm = fields.Float(string="Largeur (mm)")
    height_mm = fields.Float(string="Hauteur (mm)")
    cut_type = fields.Char(string="Type de coupe")
    bar_length_mm = fields.Float(string="Longueur barre (mm)")
    unit_price = fields.Float(string="Prix unitaire")
    computed_json = fields.Json(string="Valeurs calculees")

    def _compute_surface_m2(self):
        self.ensure_one()
        return (self.qty or 0.0) * (self.width_mm or 0.0) * (self.height_mm or 0.0) / 1000000.0

    def get_manufacturing_quantity(self):
        self.ensure_one()
        if self.category == "filling" and self.product_id:
            if self.product_id.uom_id == self.env.ref("uom.product_uom_square_meter"):
                return self._compute_surface_m2()
        return self.qty or self.length_mm or 0.0


class AluminiumJoineryMaterialSummary(models.Model):
    _name = "aluminium.joinery.material.summary"
    _description = "Synthese matiere"
    _order = "configuration_id, category, ref_text"

    configuration_id = fields.Many2one("aluminium.joinery.configuration", string="Configuration", required=True, ondelete="cascade")
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
            return self.bars_required or self.total_qty
        if self.category == "joint":
            if self.product_id and self.product_id.uom_id == self.env.ref("uom.product_uom_meter"):
                return self.total_length_mm / 1000.0
            return self.total_length_mm or self.total_qty
        if self.category == "filling":
            if self.product_id and self.product_id.uom_id == self.env.ref("uom.product_uom_square_meter"):
                return self._compute_surface_m2()
        return self.total_qty
