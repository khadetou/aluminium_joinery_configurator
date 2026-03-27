import re
import unicodedata

from odoo import _, api, fields, models
from odoo.exceptions import ValidationError


FORMULA_FAMILY_SELECTION = [
    ("qty_only", "Quantite seule"),
    ("linear_l", "Lineaire largeur"),
    ("linear_h", "Lineaire hauteur"),
    ("sum_h_l", "Somme largeur + hauteur"),
    ("perimeter", "Perimetre"),
    ("joint_combo", "Combinaison joint"),
    ("fill_dim", "Dimension remplissage"),
    ("generic_affine", "Affine generique"),
]

PRODUCT_RESOLUTION_SELECTION = [
    ("direct_product", "Produit direct"),
    ("manufactured_composite", "Produit composite fabrique"),
    ("lookup_only", "Reference texte / lookup"),
]


class AluminiumJoineryGamme(models.Model):
    _name = "aluminium.joinery.gamme"
    _description = "Gamme de menuiserie"
    _order = "name"

    name = fields.Char(string="Nom", required=True, index=True)
    code = fields.Char(string="Code", required=True, index=True)
    active = fields.Boolean(string="Actif", default=True)
    default_bar_length_mm = fields.Float(string="Longueur standard de barre (mm)", default=5800.0)
    serie_ids = fields.One2many("aluminium.joinery.serie", "gamme_id")

    _sql_constraints = [
        ("aluminium_joinery_gamme_code_uniq", "unique(code)", "Le code de la gamme doit etre unique."),
    ]

    def name_get(self):
        return [(rec.id, f"[{rec.code}] {rec.name}") for rec in self]

    @api.model
    def get_import_templates(self):
        return super().get_import_templates() + [
            {
                "label": _("Template hierarchique gamme / serie / modele"),
                "template": "/aluminium_joinery_configurator/static/xls/catalogue_hierarchie_import.xlsx",
            },
        ]


class AluminiumJoinerySerie(models.Model):
    _name = "aluminium.joinery.serie"
    _description = "Serie de menuiserie"
    _order = "gamme_id, name"

    name = fields.Char(string="Nom", required=True, index=True)
    code = fields.Char(string="Code", required=True, index=True)
    gamme_id = fields.Many2one("aluminium.joinery.gamme", string="Gamme", required=True, ondelete="cascade")
    opening_type = fields.Char(string="Type d'ouverture")
    active = fields.Boolean(string="Actif", default=True)
    modele_ids = fields.One2many("aluminium.joinery.modele", "serie_id")

    _sql_constraints = [
        (
            "aluminium_joinery_serie_code_uniq",
            "unique(code, gamme_id)",
            "Le code de la serie doit etre unique dans une gamme.",
        ),
    ]

    def name_get(self):
        return [(rec.id, f"[{rec.code}] {rec.name}") for rec in self]

    @api.model
    def get_import_templates(self):
        return super().get_import_templates() + [
            {
                "label": _("Template hierarchique gamme / serie / modele"),
                "template": "/aluminium_joinery_configurator/static/xls/catalogue_hierarchie_import.xlsx",
            },
        ]


class AluminiumJoineryModele(models.Model):
    _name = "aluminium.joinery.modele"
    _description = "Modele de menuiserie"
    _order = "serie_id, name"

    name = fields.Char(string="Nom", required=True, index=True)
    code = fields.Char(string="Code", required=True, index=True)
    serie_id = fields.Many2one("aluminium.joinery.serie", string="Serie", required=True, ondelete="cascade")
    gamme_id = fields.Many2one(related="serie_id.gamme_id", string="Gamme", store=True, readonly=True)
    active = fields.Boolean(string="Actif", default=True)
    panel_count = fields.Integer(string="Nombre de vantaux")
    rail_count = fields.Integer(string="Nombre de rails")
    x_import_key = fields.Char(string="Cle import technique", index=True, copy=False)
    sale_product_tmpl_id = fields.Many2one("product.template", string="Produit de devis")
    manufactured_product_tmpl_id = fields.Many2one("product.template", string="Produit fabrique")
    project_template_id = fields.Many2one("project.project", domain=[("is_template", "=", True)])
    rule_ids = fields.One2many("aluminium.joinery.rule", "modele_id")
    filling_rule_ids = fields.One2many("aluminium.joinery.filling.rule", "modele_id")

    _sql_constraints = [
        (
            "aluminium_joinery_modele_code_uniq",
            "unique(code, serie_id)",
            "Le code du modele doit etre unique dans une serie.",
        ),
    ]

    def name_get(self):
        return [(rec.id, f"[{rec.code}] {rec.name}") for rec in self]

    @api.model
    def get_import_templates(self):
        return super().get_import_templates() + [
            {
                "label": _("Template hierarchique gamme / serie / modele"),
                "template": "/aluminium_joinery_configurator/static/xls/catalogue_hierarchie_import.xlsx",
            },
        ]


class AluminiumJoineryRule(models.Model):
    _name = "aluminium.joinery.rule"
    _description = "Regle de menuiserie"
    _order = "modele_id, sequence, id"

    name = fields.Char(string="Nom", required=True)
    active = fields.Boolean(string="Actif", default=True)
    gamme_id = fields.Many2one("aluminium.joinery.gamme", string="Gamme", related="modele_id.gamme_id", store=True, readonly=True)
    serie_id = fields.Many2one("aluminium.joinery.serie", string="Serie", related="modele_id.serie_id", store=True, readonly=True)
    modele_id = fields.Many2one("aluminium.joinery.modele", string="Modele", required=True, ondelete="cascade")
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
    product_default_code = fields.Char(
        string="Reference article import",
        help="Reference article d'origine utilisee pendant l'import natif pour resoudre ou creer un produit temporaire.",
    )
    product_is_placeholder = fields.Boolean(
        string="Produit temporaire",
        related="product_id.is_placeholder_product",
        readonly=True,
    )
    ref_text = fields.Char(string="Reference texte")
    designation_override = fields.Char(string="Designation forcee")
    rule_code = fields.Char(string="Code technique de regle", index=True)
    profile_role = fields.Char(string="Role profil")
    cut_type = fields.Char(string="Type de coupe")
    formula_family = fields.Selection(FORMULA_FAMILY_SELECTION, string="Famille de formule")
    rule_kind = fields.Selection(
        [
            ("generic", "Generique"),
            ("formula", "Formula"),
            ("filling", "Remplissage"),
        ],
        string="Type de regle",
        default="generic",
        required=True,
    )
    target_measure = fields.Selection(
        [
            ("qty", "Quantite"),
            ("length", "Longueur"),
            ("width", "Largeur"),
            ("height", "Hauteur"),
        ],
        string="Mesure cible",
        default="qty",
        required=True,
    )
    base_dimension = fields.Selection(
        [
            ("none", "Aucune"),
            ("width", "Largeur"),
            ("height", "Hauteur"),
            ("sum_both", "Largeur + Hauteur"),
            ("perimeter", "Perimetre"),
        ],
        string="Base de calcul",
        default="none",
        required=True,
    )
    operator = fields.Selection(
        [
            ("none", "Aucun"),
            ("add", "Addition"),
            ("subtract", "Soustraction"),
        ],
        string="Operateur",
        default="none",
        required=True,
    )
    multiplier = fields.Float(string="Multiplicateur", default=1.0)
    coef_l = fields.Float(string="Coefficient largeur")
    offset_l = fields.Float(string="Decalage largeur (mm)")
    divisor_l = fields.Float(string="Diviseur largeur", default=1.0)
    coef_h = fields.Float(string="Coefficient hauteur")
    offset_h = fields.Float(string="Decalage hauteur (mm)")
    divisor_h = fields.Float(string="Diviseur hauteur", default=1.0)
    constant = fields.Float(string="Constante")
    fixed_offset_mm = fields.Float(string="Decalage fixe (mm)")
    divisor = fields.Float(string="Diviseur", default=1.0)
    apply_quantity = fields.Boolean(string="Appliquer la quantite", default=True)
    expression_value = fields.Char(
        string="Expression",
        help="Expression arithmetique securisee utilisant les variables Q, L et H.",
    )
    width_expression = fields.Char(string="Expression largeur", help="Expression de largeur du remplissage utilisant Q, L et H.")
    height_expression = fields.Char(string="Expression hauteur", help="Expression de hauteur du remplissage utilisant Q, L et H.")
    qty_expression = fields.Char(string="Expression quantite", help="Expression de quantite du remplissage utilisant Q, L et H.")
    uom_kind = fields.Selection(
        [
            ("unit", "Unite"),
            ("mm", "Millimetre"),
            ("derived", "Derivee"),
        ],
        string="Type d'unite",
        default="unit",
    )
    rounding_mode = fields.Selection(
        [
            ("none", "Aucun"),
            ("up", "Au dessus"),
            ("half_up", "Demi superieur"),
        ],
        string="Mode d'arrondi",
        default="none",
    )
    product_resolution_mode = fields.Selection(
        PRODUCT_RESOLUTION_SELECTION,
        string="Resolution produit",
        default="direct_product",
        required=True,
    )

    @api.constrains("divisor", "divisor_l", "divisor_h")
    def _check_divisors(self):
        for rule in self:
            if rule.divisor == 0 or rule.divisor_l == 0 or rule.divisor_h == 0:
                raise ValidationError("Les diviseurs doivent etre differents de zero.")

    @staticmethod
    def _slugify_placeholder_code(value):
        normalized = unicodedata.normalize("NFKD", value or "").encode("ascii", "ignore").decode("ascii").lower()
        normalized = re.sub(r"[^a-z0-9]+", "_", normalized)
        normalized = re.sub(r"_+", "_", normalized)
        return normalized.strip("_")

    @api.model
    def _product_placeholder_xmlid_name(self, default_code):
        return f"ajc_product_tmpl__{self._slugify_placeholder_code(default_code)}"

    @api.model
    def _ensure_import_xmlid(self, model_name, res_id, xmlid_name):
        module = self.env.context.get("module", "__import__")
        data = self.env["ir.model.data"].sudo().search(
            [("module", "=", module), ("name", "=", xmlid_name)],
            limit=1,
        )
        vals = {"module": module, "name": xmlid_name, "model": model_name, "res_id": res_id, "noupdate": True}
        if data:
            data.write(vals)
        else:
            self.env["ir.model.data"].sudo().create(vals)

    @api.model
    def _resolve_product_by_default_code(self, default_code, create_placeholder=False):
        code = (default_code or "").strip()
        if not code:
            return self.env["product.product"]
        product = self.env["product.product"].search([("default_code", "=", code)], limit=1)
        if product or not create_placeholder:
            return product
        placeholder_template = self.env["product.template"].create(
            {
                "name": code,
                "default_code": code,
                "type": "consu",
                "sale_ok": False,
                "purchase_ok": True,
                "is_placeholder_product_tmpl": True,
            }
        )
        self._ensure_import_xmlid(
            "product.template",
            placeholder_template.id,
            self._product_placeholder_xmlid_name(code),
        )
        return placeholder_template.product_variant_id

    @api.model
    def _prepare_product_reference_vals(self, vals):
        prepared = dict(vals)
        product_id = prepared.get("product_id")
        if product_id and not prepared.get("product_default_code"):
            product = self.env["product.product"].browse(product_id)
            prepared["product_default_code"] = product.default_code or False
        if prepared.get("product_default_code"):
            prepared["product_default_code"] = prepared["product_default_code"].strip()
        return prepared

    def _ensure_product_link_from_reference(self):
        for rule in self.filtered(lambda rec: not rec.product_id and rec.product_default_code):
            product = self._resolve_product_by_default_code(rule.product_default_code, create_placeholder=True)
            if product:
                rule.product_id = product.id

    @api.model_create_multi
    def create(self, vals_list):
        records = super().create([self._prepare_product_reference_vals(vals) for vals in vals_list])
        records._ensure_product_link_from_reference()
        return records

    def write(self, vals):
        result = super().write(self._prepare_product_reference_vals(vals))
        self._ensure_product_link_from_reference()
        return result

    @api.model
    def load(self, fields, data):
        import_fields = list(fields)
        import_data = [list(row) for row in data]
        if "product_id" in import_fields and "product_default_code" not in import_fields:
            index = import_fields.index("product_id")
            import_fields[index] = "product_default_code"
        return super().load(import_fields, import_data)

    @api.model
    def get_import_templates(self):
        return super().get_import_templates() + [
            {
                "label": _("Regles de calcul industrielles"),
                "template": "/aluminium_joinery_configurator/static/xls/rules_import.xlsx",
            },
            {
                "label": _("Regles profils"),
                "template": "/aluminium_joinery_configurator/static/xls/profiles_import.xlsx",
            },
            {
                "label": _("Regles accessoires"),
                "template": "/aluminium_joinery_configurator/static/xls/accessories_import.xlsx",
            },
            {
                "label": _("Regles joints"),
                "template": "/aluminium_joinery_configurator/static/xls/joints_import.xlsx",
            },
        ]


class AluminiumJoineryFillingRule(models.Model):
    _name = "aluminium.joinery.filling.rule"
    _description = "Regle de remplissage menuiserie"
    _order = "modele_id, sequence, id"

    name = fields.Char(string="Nom", required=True)
    active = fields.Boolean(string="Actif", default=True)
    gamme_id = fields.Many2one("aluminium.joinery.gamme", string="Gamme", related="modele_id.gamme_id", store=True, readonly=True)
    serie_id = fields.Many2one("aluminium.joinery.serie", string="Serie", related="modele_id.serie_id", store=True, readonly=True)
    modele_id = fields.Many2one("aluminium.joinery.modele", string="Modele", required=True, ondelete="cascade")
    sequence = fields.Integer(string="Sequence", default=10)
    rule_code = fields.Char(string="Code technique de regle", index=True)
    product_id = fields.Many2one("product.product", string="Article", ondelete="restrict")
    product_default_code = fields.Char(
        string="Reference article import",
        help="Reference article d'origine utilisee pendant l'import natif pour resoudre ou creer un produit temporaire.",
    )
    product_is_placeholder = fields.Boolean(
        string="Produit temporaire",
        related="product_id.is_placeholder_product",
        readonly=True,
    )
    family_width = fields.Selection(FORMULA_FAMILY_SELECTION, string="Famille largeur", default="fill_dim", required=True)
    width_coef_l = fields.Float(string="Coef largeur sur L", default=1.0)
    width_coef_h = fields.Float(string="Coef largeur sur H")
    width_constant = fields.Float(string="Constante largeur")
    family_height = fields.Selection(FORMULA_FAMILY_SELECTION, string="Famille hauteur", default="fill_dim", required=True)
    height_coef_l = fields.Float(string="Coef hauteur sur L")
    height_coef_h = fields.Float(string="Coef hauteur sur H", default=1.0)
    height_constant = fields.Float(string="Constante hauteur")
    family_qty = fields.Selection(FORMULA_FAMILY_SELECTION, string="Famille quantite", default="qty_only", required=True)
    qty_multiplier = fields.Float(string="Multiplicateur quantite", default=1.0)
    qty_constant = fields.Float(string="Constante quantite")

    @api.constrains("family_width", "family_height")
    def _check_dimension_families(self):
        for rule in self:
            if rule.family_width not in ("fill_dim", "linear_l", "linear_h", "generic_affine"):
                raise ValidationError("La famille de largeur doit rester orientee dimension.")
            if rule.family_height not in ("fill_dim", "linear_l", "linear_h", "generic_affine"):
                raise ValidationError("La famille de hauteur doit rester orientee dimension.")

    @staticmethod
    def _slugify_placeholder_code(value):
        normalized = unicodedata.normalize("NFKD", value or "").encode("ascii", "ignore").decode("ascii").lower()
        normalized = re.sub(r"[^a-z0-9]+", "_", normalized)
        normalized = re.sub(r"_+", "_", normalized)
        return normalized.strip("_")

    @api.model
    def _product_placeholder_xmlid_name(self, default_code):
        return f"ajc_product_tmpl__{self._slugify_placeholder_code(default_code)}"

    @api.model
    def _ensure_import_xmlid(self, model_name, res_id, xmlid_name):
        module = self.env.context.get("module", "__import__")
        data = self.env["ir.model.data"].sudo().search(
            [("module", "=", module), ("name", "=", xmlid_name)],
            limit=1,
        )
        vals = {"module": module, "name": xmlid_name, "model": model_name, "res_id": res_id, "noupdate": True}
        if data:
            data.write(vals)
        else:
            self.env["ir.model.data"].sudo().create(vals)

    @api.model
    def _resolve_product_by_default_code(self, default_code, create_placeholder=False):
        code = (default_code or "").strip()
        if not code:
            return self.env["product.product"]
        product = self.env["product.product"].search([("default_code", "=", code)], limit=1)
        if product or not create_placeholder:
            return product
        placeholder_template = self.env["product.template"].create(
            {
                "name": code,
                "default_code": code,
                "type": "consu",
                "sale_ok": False,
                "purchase_ok": True,
                "joinery_item_type": "filling",
                "is_placeholder_product_tmpl": True,
            }
        )
        self._ensure_import_xmlid(
            "product.template",
            placeholder_template.id,
            self._product_placeholder_xmlid_name(code),
        )
        return placeholder_template.product_variant_id

    @api.model
    def _prepare_product_reference_vals(self, vals):
        prepared = dict(vals)
        product_id = prepared.get("product_id")
        if product_id and not prepared.get("product_default_code"):
            product = self.env["product.product"].browse(product_id)
            prepared["product_default_code"] = product.default_code or False
        if prepared.get("product_default_code"):
            prepared["product_default_code"] = prepared["product_default_code"].strip()
        return prepared

    def _ensure_product_link_from_reference(self):
        for rule in self.filtered(lambda rec: not rec.product_id and rec.product_default_code):
            product = self._resolve_product_by_default_code(rule.product_default_code, create_placeholder=True)
            if product:
                rule.product_id = product.id

    @api.model_create_multi
    def create(self, vals_list):
        records = super().create([self._prepare_product_reference_vals(vals) for vals in vals_list])
        records._ensure_product_link_from_reference()
        return records

    def write(self, vals):
        result = super().write(self._prepare_product_reference_vals(vals))
        self._ensure_product_link_from_reference()
        return result

    @api.model
    def load(self, fields, data):
        import_fields = list(fields)
        import_data = [list(row) for row in data]
        if "product_id" in import_fields and "product_default_code" not in import_fields:
            index = import_fields.index("product_id")
            import_fields[index] = "product_default_code"
        return super().load(import_fields, import_data)

    @api.model
    def get_import_templates(self):
        return super().get_import_templates() + [
            {
                "label": _("Regles de remplissage industrielles"),
                "template": "/aluminium_joinery_configurator/static/xls/fillings_import.xlsx",
            },
        ]
