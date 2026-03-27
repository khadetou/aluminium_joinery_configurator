import re
import unicodedata

from odoo import _, api, fields, models


ITEM_TYPE_SELECTION = [
    ("profile", "Profile"),
    ("accessory", "Accessoire"),
    ("joint", "Joint"),
    ("filling", "Remplissage"),
    ("finished", "Produit fini"),
    ("service", "Service"),
]


def _slugify_placeholder_code(value):
    normalized = unicodedata.normalize("NFKD", value or "").encode("ascii", "ignore").decode("ascii").lower()
    normalized = re.sub(r"[^a-z0-9]+", "_", normalized)
    normalized = re.sub(r"_+", "_", normalized)
    return normalized.strip("_")


def _ensure_import_xmlid(env, model_name, res_id, xmlid_name):
    module = env.context.get("module", "__import__")
    data = env["ir.model.data"].sudo().search(
        [("module", "=", module), ("name", "=", xmlid_name)],
        limit=1,
    )
    vals = {"module": module, "name": xmlid_name, "model": model_name, "res_id": res_id, "noupdate": True}
    if data:
        data.write(vals)
    else:
        env["ir.model.data"].sudo().create(vals)


def _resolve_product_variant_by_default_code(env, default_code, *, create_placeholder=False, template_vals=None):
    code = (default_code or "").strip()
    if not code:
        return env["product.product"]
    product = env["product.product"].search([("default_code", "=", code)], limit=1)
    if product or not create_placeholder:
        return product
    placeholder_template = env["product.template"].create(
        {
            "name": code,
            "default_code": code,
            "type": "consu",
            "sale_ok": False,
            "purchase_ok": True,
            "is_placeholder_product_tmpl": True,
            **(template_vals or {}),
        }
    )
    _ensure_import_xmlid(
        env,
        "product.template",
        placeholder_template.id,
        f"ajc_product_tmpl__{_slugify_placeholder_code(code)}",
    )
    return placeholder_template.product_variant_id


class ProductTemplate(models.Model):
    _inherit = "product.template"

    ajc_item_type = fields.Selection(
        ITEM_TYPE_SELECTION,
        string="Type d'article menuiserie",
    )
    ajc_bar_length_mm = fields.Float(string="Longueur standard de barre (mm)")
    ajc_default_cut_type = fields.Char(string="Type de coupe par defaut")
    ajc_is_configurator_item = fields.Boolean(string="Article du configurateur")
    ajc_usage_role = fields.Char(string="Role d'usage menuiserie")
    ajc_can_be_sale_component = fields.Boolean(string="Peut etre devisé", default=True)
    ajc_can_be_mrp_component = fields.Boolean(string="Peut etre composant de fabrication", default=True)
    joinery_item_type = fields.Selection(ITEM_TYPE_SELECTION, string="Type d'article configurateur")
    joinery_bar_length_mm = fields.Float(string="Longueur configurateur standard (mm)")
    joinery_usage_role = fields.Char(string="Role d'usage configurateur")
    schuller_item_type = fields.Selection(ITEM_TYPE_SELECTION, string="Type d'article industriel")
    schuller_bar_length_mm = fields.Float(string="Longueur industrielle standard (mm)")
    schuller_usage_role = fields.Char(string="Role d'usage industriel")
    is_joinery_composite = fields.Boolean(string="Produit composite menuiserie")
    is_placeholder_product_tmpl = fields.Boolean(
        string="Produit temporaire configurateur",
        default=False,
        copy=False,
        help="Produit cree automatiquement pour permettre l'import des regles avant l'import complet des articles.",
    )
    manufacturing_mode = fields.Selection(
        [
            ("standard", "Standard"),
            ("manufactured_composite", "Composite fabrique"),
        ],
        string="Mode de fabrication menuiserie",
        default="standard",
    )

    @api.model
    def get_import_templates(self):
        return super().get_import_templates() + [
            {
                "label": _("Articles industriels menuiserie"),
                "template": "/aluminium_joinery_configurator/static/xls/articles_import.xlsx",
            },
        ]


class MrpBom(models.Model):
    _inherit = "mrp.bom"

    product_default_code = fields.Char(
        string="Reference produit import",
        help="Reference article utilisee pendant l'import natif pour resoudre ou creer le produit fini de la nomenclature.",
    )

    @api.model
    def _prepare_product_reference_vals(self, vals):
        prepared = dict(vals)
        product_tmpl_id = prepared.get("product_tmpl_id")
        if product_tmpl_id and not prepared.get("product_default_code"):
            prepared["product_default_code"] = self.env["product.template"].browse(product_tmpl_id).default_code or False
        if prepared.get("product_default_code"):
            prepared["product_default_code"] = prepared["product_default_code"].strip()
        if not prepared.get("product_tmpl_id") and prepared.get("product_default_code"):
            product = _resolve_product_variant_by_default_code(
                self.env,
                prepared["product_default_code"],
                create_placeholder=True,
                template_vals={
                    "joinery_item_type": "finished",
                    "is_joinery_composite": True,
                    "manufacturing_mode": "manufactured_composite",
                },
            )
            if product:
                prepared["product_tmpl_id"] = product.product_tmpl_id.id
        return prepared

    def _ensure_product_template_from_reference(self):
        for bom in self.filtered(lambda rec: not rec.product_tmpl_id and rec.product_default_code):
            product = _resolve_product_variant_by_default_code(
                self.env,
                bom.product_default_code,
                create_placeholder=True,
                template_vals={
                    "joinery_item_type": "finished",
                    "is_joinery_composite": True,
                    "manufacturing_mode": "manufactured_composite",
                },
            )
            if product:
                bom.product_tmpl_id = product.product_tmpl_id.id

    @api.model_create_multi
    def create(self, vals_list):
        records = super().create([self._prepare_product_reference_vals(vals) for vals in vals_list])
        records._ensure_product_template_from_reference()
        return records

    def write(self, vals):
        result = super().write(self._prepare_product_reference_vals(vals))
        self._ensure_product_template_from_reference()
        return result

    @api.model
    def get_import_templates(self):
        return super().get_import_templates() + [
            {
                "label": _("BOMs composites menuiserie"),
                "template": "/aluminium_joinery_configurator/static/xls/boms_import.xlsx",
            },
        ]


class MrpBomLine(models.Model):
    _inherit = "mrp.bom.line"

    component_default_code = fields.Char(
        string="Reference composant import",
        help="Reference article utilisee pendant l'import natif pour resoudre ou creer le composant de nomenclature.",
    )
    component_is_placeholder = fields.Boolean(
        related="product_id.is_placeholder_product",
        string="Composant temporaire",
        readonly=True,
    )

    @api.model
    def _prepare_component_reference_vals(self, vals):
        prepared = dict(vals)
        product_id = prepared.get("product_id")
        if product_id and not prepared.get("component_default_code"):
            prepared["component_default_code"] = self.env["product.product"].browse(product_id).default_code or False
        if prepared.get("component_default_code"):
            prepared["component_default_code"] = prepared["component_default_code"].strip()
        if not prepared.get("product_id") and prepared.get("component_default_code"):
            product = _resolve_product_variant_by_default_code(
                self.env,
                prepared["component_default_code"],
                create_placeholder=True,
            )
            if product:
                prepared["product_id"] = product.id
        return prepared

    def _ensure_component_from_reference(self):
        for line in self.filtered(lambda rec: not rec.product_id and rec.component_default_code):
            product = _resolve_product_variant_by_default_code(self.env, line.component_default_code, create_placeholder=True)
            if product:
                line.product_id = product.id

    @api.model_create_multi
    def create(self, vals_list):
        records = super().create([self._prepare_component_reference_vals(vals) for vals in vals_list])
        records._ensure_component_from_reference()
        return records

    def write(self, vals):
        result = super().write(self._prepare_component_reference_vals(vals))
        self._ensure_component_from_reference()
        return result

    @api.model
    def load(self, fields, data):
        import_fields = list(fields)
        import_data = [list(row) for row in data]
        if "product_id" in import_fields and "component_default_code" not in import_fields:
            index = import_fields.index("product_id")
            import_fields[index] = "component_default_code"
        return super().load(import_fields, import_data)

    @api.model
    def get_import_templates(self):
        return super().get_import_templates() + [
            {
                "label": _("Lignes de BOM menuiserie"),
                "template": "/aluminium_joinery_configurator/static/xls/bom_lines_import.xlsx",
            },
        ]


class ProductProduct(models.Model):
    _inherit = "product.product"

    is_placeholder_product = fields.Boolean(
        related="product_tmpl_id.is_placeholder_product_tmpl",
        string="Produit temporaire configurateur",
        store=True,
        readonly=True,
    )
