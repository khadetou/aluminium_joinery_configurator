import importlib.util
import zipfile
from io import BytesIO
from pathlib import Path

from odoo.tests.common import TransactionCase


class TestJoineryConfiguration(TransactionCase):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        script_path = Path(__file__).resolve().parents[1] / "scripts" / "generate_import_templates.py"
        spec = importlib.util.spec_from_file_location("ajc_generate_import_templates", script_path)
        cls.generator = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(cls.generator)

    def _native_import(self, model_name, rows):
        headers = rows[0]
        data = rows[1:]
        result = self.env[model_name].load(headers, data)
        messages = result.get("messages") or []
        self.assertFalse(messages, "\n".join(message.get("message", str(message)) for message in messages))
        return result

    def _import_native_industrial_dataset(self):
        rows = self.generator.build_rows()
        self._native_import("aluminium.joinery.gamme", rows["Catalogue"])
        self._native_import("product.template", rows["Articles"])
        self._native_import("aluminium.joinery.rule", rows["Rules"])
        self._native_import("aluminium.joinery.filling.rule", rows["Filling Rules"])
        self._native_import("mrp.bom", rows["BOMs Complete"])
        return rows

    def _find_rule_row(self, rows, default_code):
        header = rows["Rules"][0]
        product_index = header.index("product_id")
        for row in rows["Rules"][1:]:
            if row[product_index] == default_code:
                return [header, row]
        self.fail(f"Aucune regle trouvee pour {default_code}")

    def _find_article_row(self, rows, default_code):
        header = rows["Articles"][0]
        product_index = header.index("default_code")
        for row in rows["Articles"][1:]:
            if row[product_index] == default_code:
                return [header, row]
        self.fail(f"Aucun article trouve pour {default_code}")

    def _find_filling_row(self, rows, model_xmlid):
        header = rows["Filling Rules"][0]
        model_index = header.index("modele_id/id")
        for row in rows["Filling Rules"][1:]:
            if row[model_index] == model_xmlid:
                return [header, row]
        self.fail(f"Aucune regle de remplissage trouvee pour {model_xmlid}")

    def test_native_hierarchy_import_creates_models_with_external_ids(self):
        rows = self.generator.build_rows()
        result = self._native_import("aluminium.joinery.gamme", rows["Catalogue"])

        self.assertEqual(len(result["ids"]), len(rows["Catalogue"]) - 1)

        modele = self.env["aluminium.joinery.modele"].search(
            [("x_import_key", "=", "comfort__comfort_coulissant_300_sv__coupes_coulissant_2_vantaux_2_rails")],
            limit=1,
        )
        self.assertTrue(modele)
        self.assertEqual(modele.code, "coupes_coulissant_2_vantaux_2_rails")

    def test_catalogue_rows_remain_three_level_native_safe(self):
        rows = self.generator.build_rows()["Catalogue"]
        header = rows[0]
        self.assertNotIn("serie_ids/modele_ids/filling_rule_ids/id", header)

    def test_article_template_uses_native_product_type_field(self):
        rows = self.generator.build_rows()
        self.assertIn("type", rows["Articles"][0])
        self.assertNotIn("detailed_type", rows["Articles"][0])
        self.assertNotIn("uom_po_id/id", rows["Articles"][0])
        self.assertIn("categ_id/id", rows["Articles"][0])
        type_index = rows["Articles"][0].index("type")
        self.assertEqual(rows["Articles"][1][type_index], "consu")
        uom_index = rows["Articles"][0].index("uom_id/id")

        profile_article = self._find_article_row(rows, "KCL301 / KCL315")[1]
        self.assertEqual(profile_article[uom_index], "uom.product_uom_meter")

        filling_header, filling_row = self._find_filling_row(rows, "ajc_modele__masai_a_frappe__fenetre_fixe")
        filling_product_code = filling_row[filling_header.index("product_id")]
        filling_article = self._find_article_row(rows, filling_product_code)[1]
        self.assertEqual(filling_article[uom_index], "uom.product_uom_square_meter")

    def test_bom_templates_use_default_code_references(self):
        rows = self.generator.build_rows()
        self.assertIn("product_default_code", rows["BOMs Complete"][0])
        self.assertIn("bom_line_ids/component_default_code", rows["BOMs Complete"][0])
        self.assertIn("component_default_code", rows["BOM Lines"][0])
        self.assertNotIn("bom_line_ids/product_id", rows["BOMs Complete"][0])

    def test_native_bom_import_with_missing_products_creates_placeholders(self):
        rows = self.generator.build_rows()
        self._native_import("mrp.bom", rows["BOMs Complete"])

        bom = self.env["mrp.bom"].search([("code", "=", "BOM_019_059_019_061")], limit=1)
        self.assertTrue(bom)
        self.assertEqual(bom.product_tmpl_id.default_code, "019.059+019.061")
        self.assertTrue(bom.product_tmpl_id.is_placeholder_product_tmpl)
        self.assertEqual(set(bom.bom_line_ids.mapped("component_default_code")), {"019.059", "019.061"})
        self.assertTrue(all(line.product_id for line in bom.bom_line_ids))

    def test_native_rule_import_links_rules_to_imported_models(self):
        rows = self.generator.build_rows()
        self._native_import("aluminium.joinery.gamme", rows["Catalogue"])
        self._native_import("product.template", rows["Articles"])
        self._native_import("aluminium.joinery.rule", rows["Rules"])
        self._native_import("aluminium.joinery.filling.rule", rows["Filling Rules"])

        modele = self.env["aluminium.joinery.modele"].search(
            [("x_import_key", "=", "comfort__comfort_coulissant_300_sv__coupes_coulissant_2_vantaux_2_rails")],
            limit=1,
        )
        self.assertTrue(modele.rule_ids.filtered("active"))
        self.assertTrue(modele.filling_rule_ids.filtered("active"))
        self.assertTrue(modele.rule_ids.filtered(lambda rule: rule.product_id.default_code == "KCL301 / KCL315"))
        self.assertTrue(modele.filling_rule_ids.filtered(lambda rule: rule.product_id))

    def test_filling_rule_import_with_missing_product_creates_placeholder(self):
        rows = self.generator.build_rows()
        self._native_import("aluminium.joinery.gamme", rows["Catalogue"])

        header = rows["Filling Rules"][0][:]
        row = rows["Filling Rules"][1][:]
        product_index = header.index("product_id")
        row[product_index] = "VITRAGE.TEST"
        self._native_import("aluminium.joinery.filling.rule", [header, row])

        filling_rule = self.env["aluminium.joinery.filling.rule"].search(
            [("product_default_code", "=", "VITRAGE.TEST")],
            limit=1,
        )
        self.assertTrue(filling_rule)
        self.assertTrue(filling_rule.product_id)
        self.assertTrue(filling_rule.product_id.is_placeholder_product)
        self.assertEqual(filling_rule.product_id.uom_id, self.env.ref("uom.product_uom_square_meter"))

    def test_generated_filling_rule_matches_masai_fixed_window_vba(self):
        rows = self.generator.build_rows()
        header, row = self._find_filling_row(rows, "ajc_modele__masai_a_frappe__fenetre_fixe")
        values = dict(zip(header, row))

        self.assertEqual(values["family_width"], "fill_dim")
        self.assertEqual(values["width_coef_l"], 1.0)
        self.assertEqual(values["width_coef_h"], 0.0)
        self.assertEqual(values["width_constant"], -60.0)
        self.assertEqual(values["family_height"], "fill_dim")
        self.assertEqual(values["height_coef_l"], 0.0)
        self.assertEqual(values["height_coef_h"], 1.0)
        self.assertEqual(values["height_constant"], -60.0)
        self.assertEqual(values["family_qty"], "qty_only")
        self.assertEqual(values["qty_multiplier"], 1.0)
        self.assertEqual(values["qty_constant"], 0.0)
        self.assertTrue(values["product_id"].startswith("FILL__"))

    def test_rule_import_with_missing_product_creates_placeholder(self):
        rows = self.generator.build_rows()
        self._native_import("aluminium.joinery.gamme", rows["Catalogue"])
        self._native_import("aluminium.joinery.rule", self._find_rule_row(rows, "KCL301 / KCL315"))

        rule = self.env["aluminium.joinery.rule"].search([("product_default_code", "=", "KCL301 / KCL315")], limit=1)
        self.assertTrue(rule)
        self.assertTrue(rule.product_id)
        self.assertTrue(rule.product_id.is_placeholder_product)
        self.assertEqual(rule.product_id.default_code, "KCL301 / KCL315")

    def test_importing_article_after_placeholder_reuses_same_product(self):
        rows = self.generator.build_rows()
        self._native_import("aluminium.joinery.gamme", rows["Catalogue"])
        self._native_import("aluminium.joinery.rule", self._find_rule_row(rows, "KCL301 / KCL315"))

        rule = self.env["aluminium.joinery.rule"].search([("product_default_code", "=", "KCL301 / KCL315")], limit=1)
        placeholder_product = rule.product_id
        self.assertTrue(placeholder_product.is_placeholder_product)

        self._native_import("product.template", self._find_article_row(rows, "KCL301 / KCL315"))

        rule.invalidate_recordset()
        placeholder_product.invalidate_recordset()
        refreshed_rule = self.env["aluminium.joinery.rule"].browse(rule.id)
        refreshed_product = self.env["product.product"].browse(placeholder_product.id)
        self.assertEqual(refreshed_rule.product_id.id, refreshed_product.id)
        self.assertFalse(refreshed_product.is_placeholder_product)

    def test_native_import_end_to_end_allows_calculation_after_import(self):
        self._import_native_industrial_dataset()

        gamme = self.env["aluminium.joinery.gamme"].search([("code", "=", "comfort")], limit=1)
        serie = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "comfort_coulissant_300_sv"), ("gamme_id", "=", gamme.id)],
            limit=1,
        )
        modele = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "coupes_coulissant_2_vantaux_2_rails"), ("serie_id", "=", serie.id)],
            limit=1,
        )

        self.assertTrue(modele.rule_ids.filtered("active"))

        configuration = self.env["aluminium.joinery.configuration"].create(
            {"partner_id": self.env.ref("base.res_partner_1").id}
        )
        self.env["aluminium.joinery.configuration.line"].create(
            {
                "configuration_id": configuration.id,
                "gamme_id": gamme.id,
                "serie_id": serie.id,
                "modele_id": modele.id,
                "qty": 1,
                "width_mm": 2400,
                "height_mm": 1800,
            }
        )

        configuration.action_calculate()

        self.assertTrue(configuration.result_line_ids)
        self.assertTrue(configuration.result_line_ids.filtered(lambda line: line.category == "profile"))
        filling_line = configuration.result_line_ids.filtered(lambda line: line.category == "filling")[:1]
        self.assertTrue(filling_line)
        self.assertAlmostEqual(filling_line.qty, 2.0)
        self.assertAlmostEqual(filling_line.width_mm, 1117.5)
        self.assertAlmostEqual(filling_line.height_mm, 1656.0)
        filling_summary = configuration.summary_ids.filtered(lambda line: line.category == "filling")[:1]
        self.assertTrue(filling_summary)
        self.assertEqual(filling_summary.configuration_line_id, configuration.line_ids[:1])
        self.assertFalse(configuration.calculation_note)

    def test_masai_fixed_window_filling_matches_vba_reference(self):
        self._import_native_industrial_dataset()

        gamme = self.env["aluminium.joinery.gamme"].search([("code", "=", "masai")], limit=1)
        serie = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "masai_a_frappe"), ("gamme_id", "=", gamme.id)],
            limit=1,
        )
        modele = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "fenetre_fixe"), ("serie_id", "=", serie.id)],
            limit=1,
        )

        self.assertTrue(modele.filling_rule_ids.filtered("active"))

        configuration = self.env["aluminium.joinery.configuration"].create(
            {"partner_id": self.env.ref("base.res_partner_1").id}
        )
        self.env["aluminium.joinery.configuration.line"].create(
            {
                "configuration_id": configuration.id,
                "gamme_id": gamme.id,
                "serie_id": serie.id,
                "modele_id": modele.id,
                "qty": 1,
                "width_mm": 1200,
                "height_mm": 1200,
            }
        )

        configuration.action_calculate()

        filling_line = configuration.result_line_ids.filtered(lambda line: line.category == "filling")[:1]
        self.assertTrue(filling_line)
        self.assertAlmostEqual(filling_line.width_mm, 1140.0)
        self.assertAlmostEqual(filling_line.height_mm, 1140.0)
        self.assertAlmostEqual(filling_line.qty, 1.0)

        filling_summary = configuration.summary_ids.filtered(lambda line: line.category == "filling")[:1]
        self.assertTrue(filling_summary)
        self.assertAlmostEqual(filling_summary.width_mm, 1140.0)
        self.assertAlmostEqual(filling_summary.height_mm, 1140.0)
        self.assertAlmostEqual(filling_summary.total_qty, 1.0)
        self.assertEqual(filling_summary.configuration_line_id.modele_id, modele)
        self.assertTrue(filling_line.product_id)
        self.assertEqual(filling_line.product_id.uom_id, self.env.ref("uom.product_uom_square_meter"))

    def test_structured_views_group_results_and_summary_by_configuration_line(self):
        self._import_native_industrial_dataset()

        partner = self.env.ref("base.res_partner_1")
        configuration = self.env["aluminium.joinery.configuration"].create({"partner_id": partner.id})

        masai = self.env["aluminium.joinery.gamme"].search([("code", "=", "masai")], limit=1)
        masai_serie = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "masai_a_frappe"), ("gamme_id", "=", masai.id)],
            limit=1,
        )
        fixed_model = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "fenetre_fixe"), ("serie_id", "=", masai_serie.id)],
            limit=1,
        )

        comfort = self.env["aluminium.joinery.gamme"].search([("code", "=", "comfort")], limit=1)
        comfort_serie = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "comfort_coulissant_300_sv"), ("gamme_id", "=", comfort.id)],
            limit=1,
        )
        comfort_model = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "coupes_coulissant_2_vantaux_2_rails"), ("serie_id", "=", comfort_serie.id)],
            limit=1,
        )

        fixed_line = self.env["aluminium.joinery.configuration.line"].create(
            {
                "configuration_id": configuration.id,
                "sequence": 10,
                "gamme_id": masai.id,
                "serie_id": masai_serie.id,
                "modele_id": fixed_model.id,
                "qty": 1,
                "width_mm": 1200,
                "height_mm": 1200,
            }
        )
        sliding_line = self.env["aluminium.joinery.configuration.line"].create(
            {
                "configuration_id": configuration.id,
                "sequence": 20,
                "gamme_id": comfort.id,
                "serie_id": comfort_serie.id,
                "modele_id": comfort_model.id,
                "qty": 1,
                "width_mm": 2400,
                "height_mm": 1800,
            }
        )

        configuration.action_calculate()

        self.assertEqual(set(configuration.summary_ids.mapped("configuration_line_id").ids), {fixed_line.id, sliding_line.id})
        self.assertIn("Profiles", configuration.results_structured_html)
        self.assertIn("Accessoires / Quincailleries", configuration.results_structured_html)
        self.assertIn("Remplissage / Vitrages et panneaux", configuration.results_structured_html)
        self.assertIn("Section / Coupe", configuration.results_structured_html)
        self.assertIn(fixed_line.modele_id.name, configuration.results_structured_html)
        self.assertIn(sliding_line.modele_id.name, configuration.results_structured_html)
        self.assertIn(fixed_line.modele_id.name, configuration.summary_structured_html)
        self.assertIn(sliding_line.modele_id.name, configuration.summary_structured_html)
        self.assertNotIn(">PU<", configuration.summary_structured_html)
        self.assertNotIn(">Total<", configuration.summary_structured_html)
        self.assertIn("Barres necessaires", configuration.summary_structured_html)
        self.assertIn("Pieces / Profiles a couper", configuration.production_structured_html)
        self.assertIn("Quincaillerie a sortir", configuration.production_structured_html)
        self.assertIn("Vitrages / Panneaux", configuration.production_structured_html)
        self.assertIn("Nomenclature", configuration.production_structured_html)
        self.assertIn("Sous-ensembles composites", configuration.production_structured_html)

        result_groups = configuration._get_report_groups("result")
        self.assertEqual(len(result_groups), 2)
        self.assertEqual(result_groups[0]["gamme_name"], fixed_line.gamme_id.display_name)
        self.assertEqual(result_groups[1]["gamme_name"], sliding_line.gamme_id.display_name)
        self.assertTrue(any(section["title"] == "PROFILES" for section in result_groups[0]["lines"][0]["sections"]))
        self.assertIn("Longueur requise (mm)", result_groups[0]["lines"][0]["sections"][0]["headers"])
        self.assertIn("Barres", result_groups[0]["lines"][0]["sections"][0]["headers"])
        self.assertIn("Section / Coupe", result_groups[0]["lines"][0]["sections"][0]["headers"])

        summary_groups = configuration._get_report_groups("summary")
        self.assertEqual(len(summary_groups), 2)
        self.assertTrue(any(section["title"] == "REMPLISSAGES" for section in summary_groups[0]["lines"][0]["sections"]))
        profile_summary_section = next(section for section in summary_groups[0]["lines"][0]["sections"] if section["title"] == "PROFILES")
        self.assertNotIn("PU", profile_summary_section["headers"])
        self.assertNotIn("Total", profile_summary_section["headers"])

        grouped_payload = configuration._get_grouped_result_payload()
        self.assertTrue(grouped_payload["profiles"])
        self.assertEqual(
            sum(item["bars_required"] for item in grouped_payload["profiles"]),
            sum(configuration.summary_ids.filtered(lambda rec: rec.category == "profile").mapped("bars_required")),
        )
        self.assertEqual(
            sum(item["billed_length_mm"] for item in grouped_payload["profiles"]),
            sum(configuration.summary_ids.filtered(lambda rec: rec.category == "profile").mapped("billed_length_mm")),
        )
        grouped_sections = configuration._get_grouped_result_sections()
        grouped_profile_section = next(section for section in grouped_sections if section["title"] == "PROFILES")
        self.assertIn("Longueur facturable (mm)", grouped_profile_section["headers"])

    def test_generate_quotation_uses_business_uoms_and_converted_quantities(self):
        self._import_native_industrial_dataset()

        gamme = self.env["aluminium.joinery.gamme"].search([("code", "=", "comfort")], limit=1)
        serie = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "comfort_coulissant_300_sv"), ("gamme_id", "=", gamme.id)],
            limit=1,
        )
        modele = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "coupes_coulissant_2_vantaux_2_rails"), ("serie_id", "=", serie.id)],
            limit=1,
        )

        configuration = self.env["aluminium.joinery.configuration"].create(
            {"partner_id": self.env.ref("base.res_partner_1").id}
        )
        self.env["aluminium.joinery.configuration.line"].create(
            {
                "configuration_id": configuration.id,
                "gamme_id": gamme.id,
                "serie_id": serie.id,
                "modele_id": modele.id,
                "qty": 1,
                "width_mm": 2400,
                "height_mm": 1800,
            }
        )

        configuration.action_calculate()
        action = configuration.action_generate_quotation()
        order = self.env["sale.order"].browse(action["res_id"])

        profile_line = order.order_line.filtered(
            lambda line: line.joinery_summary_id and line.joinery_summary_id.category == "profile" and not line.display_type
        )[:1]
        self.assertTrue(profile_line)
        self.assertEqual(profile_line.product_uom_id, self.env.ref("uom.product_uom_meter"))
        self.assertGreater(profile_line.joinery_summary_id.bars_required, 0)
        self.assertGreaterEqual(
            profile_line.joinery_summary_id.billed_length_mm,
            profile_line.joinery_summary_id.total_length_mm,
        )
        self.assertAlmostEqual(profile_line.product_uom_qty, profile_line.joinery_summary_id.billed_length_mm / 1000.0)

        joint_line = order.order_line.filtered(
            lambda line: line.joinery_summary_id and line.joinery_summary_id.category == "joint" and not line.display_type
        )[:1]
        self.assertTrue(joint_line)
        self.assertEqual(joint_line.product_uom_id, self.env.ref("uom.product_uom_meter"))
        self.assertAlmostEqual(joint_line.product_uom_qty, joint_line.joinery_summary_id.total_length_mm / 1000.0)

        filling_line = order.order_line.filtered(
            lambda line: line.joinery_summary_id and line.joinery_summary_id.category == "filling" and not line.display_type
        )[:1]
        self.assertTrue(filling_line)
        self.assertEqual(filling_line.product_uom_id, self.env.ref("uom.product_uom_square_meter"))
        self.assertAlmostEqual(filling_line.product_uom_qty, filling_line.joinery_summary_id.get_sale_quantity())
        self.assertFalse(
            order.order_line.filtered(
                lambda line: line.display_type == "line_note" and line.joinery_summary_id and line.joinery_summary_id.category == "filling"
            )
        )

    def test_profile_sale_quantity_uses_standard_bar_billed_length(self):
        product_tmpl = self.env["product.template"].create(
            {
                "name": "Profile test palette",
                "default_code": "PTEST-5800",
                "joinery_item_type": "profile",
                "joinery_bar_length_mm": 5800,
                "list_price": 1000.0,
            }
        )
        summary = self.env["aluminium.joinery.material.summary"].create(
            {
                "configuration_id": self.env["aluminium.joinery.configuration"].create({}).id,
                "category": "profile",
                "product_id": product_tmpl.product_variant_id.id,
                "ref_text": "PTEST-5800",
                "designation": "Profile test palette",
                "total_qty": 1.0,
                "total_length_mm": 12000.0,
                "bar_length_mm": 5800.0,
                "bars_required": 3,
                "billed_length_mm": 17400.0,
                "unit_price": 1000.0,
                "total_price": 17400.0,
            }
        )
        self.assertAlmostEqual(summary.get_sale_quantity(), 17.4)

    def test_calculation_still_works_after_linking_placeholder_product(self):
        rows = self.generator.build_rows()
        self._native_import("aluminium.joinery.gamme", rows["Catalogue"])
        self._native_import("aluminium.joinery.rule", self._find_rule_row(rows, "KCL301 / KCL315"))
        self._native_import("product.template", self._find_article_row(rows, "KCL301 / KCL315"))

        gamme = self.env["aluminium.joinery.gamme"].search([("code", "=", "comfort")], limit=1)
        serie = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "comfort_coulissant_300_sv"), ("gamme_id", "=", gamme.id)],
            limit=1,
        )
        modele = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "coupes_coulissant_2_vantaux_2_rails"), ("serie_id", "=", serie.id)],
            limit=1,
        )

        configuration = self.env["aluminium.joinery.configuration"].create(
            {"partner_id": self.env.ref("base.res_partner_1").id}
        )
        self.env["aluminium.joinery.configuration.line"].create(
            {
                "configuration_id": configuration.id,
                "gamme_id": gamme.id,
                "serie_id": serie.id,
                "modele_id": modele.id,
                "qty": 1,
                "width_mm": 2400,
                "height_mm": 1800,
            }
        )

        configuration.action_calculate()
        self.assertTrue(configuration.result_line_ids)

    def test_portal_upsert_single_line_configuration_reuses_draft_configuration(self):
        self._import_native_industrial_dataset()

        partner = self.env.ref("base.res_partner_1")
        gamme = self.env["aluminium.joinery.gamme"].search([("code", "=", "masai")], limit=1)
        serie = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "masai_a_frappe"), ("gamme_id", "=", gamme.id)],
            limit=1,
        )
        modele = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "fenetre_fixe"), ("serie_id", "=", serie.id)],
            limit=1,
        )

        payload = {
            "project_name": "Villa client",
            "gamme_id": gamme.id,
            "serie_id": serie.id,
            "modele_id": modele.id,
            "qty": 1,
            "width_mm": 1200,
            "height_mm": 1200,
        }
        configuration = self.env["aluminium.joinery.configuration"].portal_upsert_single_line_configuration(partner, payload)

        self.assertEqual(configuration.partner_id, partner)
        self.assertEqual(configuration.project_name, "Villa client")
        self.assertEqual(len(configuration.line_ids), 1)
        self.assertEqual(configuration.line_ids.modele_id, modele)

        payload["project_name"] = "Villa client maj"
        payload["width_mm"] = 1400
        updated = self.env["aluminium.joinery.configuration"].portal_upsert_single_line_configuration(
            partner,
            payload,
            configuration=configuration,
        )

        self.assertEqual(updated.id, configuration.id)
        self.assertEqual(updated.project_name, "Villa client maj")
        self.assertEqual(len(updated.line_ids), 1)
        self.assertEqual(updated.line_ids.width_mm, 1400)

    def test_portal_upsert_configuration_supports_multiple_lines(self):
        self._import_native_industrial_dataset()

        partner = self.env.ref("base.res_partner_1")
        gamme = self.env["aluminium.joinery.gamme"].search([("code", "=", "masai")], limit=1)
        serie_frappe = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "masai_a_frappe"), ("gamme_id", "=", gamme.id)],
            limit=1,
        )
        modele_fixe = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "fenetre_fixe"), ("serie_id", "=", serie_frappe.id)],
            limit=1,
        )
        serie_coulissant = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "masai_coulissant"), ("gamme_id", "=", gamme.id)],
            limit=1,
        )
        modele_coulissant = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "fenetre_2_vantaux_2_rails"), ("serie_id", "=", serie_coulissant.id)],
            limit=1,
        )

        configuration = self.env["aluminium.joinery.configuration"].portal_upsert_configuration(
            partner,
            [
                {
                    "sequence": 10,
                    "gamme_id": gamme.id,
                    "serie_id": serie_frappe.id,
                    "modele_id": modele_fixe.id,
                    "qty": 1,
                    "width_mm": 1200,
                    "height_mm": 1200,
                },
                {
                    "sequence": 20,
                    "gamme_id": gamme.id,
                    "serie_id": serie_coulissant.id,
                    "modele_id": modele_coulissant.id,
                    "qty": 2,
                    "width_mm": 2400,
                    "height_mm": 1800,
                },
            ],
            project_name="Projet multi-lignes",
        )

        self.assertEqual(configuration.partner_id, partner)
        self.assertEqual(configuration.project_name, "Projet multi-lignes")
        self.assertEqual(len(configuration.line_ids), 2)
        ordered_lines = configuration.line_ids.sorted(lambda rec: (rec.sequence, rec.id))
        self.assertEqual(ordered_lines[0].modele_id, modele_fixe)
        self.assertEqual(ordered_lines[1].modele_id, modele_coulissant)
        self.assertEqual(configuration._get_portal_form_values()["lines"][1]["qty"], 2)

    def test_portal_duplicate_for_partner_creates_clean_draft_copy(self):
        self._import_native_industrial_dataset()

        partner = self.env.ref("base.res_partner_1")
        gamme = self.env["aluminium.joinery.gamme"].search([("code", "=", "masai")], limit=1)
        serie = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "masai_a_frappe"), ("gamme_id", "=", gamme.id)],
            limit=1,
        )
        modele = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "fenetre_fixe"), ("serie_id", "=", serie.id)],
            limit=1,
        )

        configuration = self.env["aluminium.joinery.configuration"].create(
            {
                "partner_id": partner.id,
                "project_name": "Projet original",
                "state": "quoted",
            }
        )
        self.env["aluminium.joinery.configuration.line"].create(
            {
                "configuration_id": configuration.id,
                "gamme_id": gamme.id,
                "serie_id": serie.id,
                "modele_id": modele.id,
                "qty": 1,
                "width_mm": 1200,
                "height_mm": 1200,
                "state": "calculated",
                "calculation_hash": "abc",
            }
        )

        duplicated = configuration.portal_duplicate_for_partner(partner)

        self.assertNotEqual(duplicated.id, configuration.id)
        self.assertEqual(duplicated.partner_id, partner)
        self.assertEqual(duplicated.state, "draft")
        self.assertEqual(duplicated.project_name, "Projet original - copie")
        self.assertFalse(duplicated.sale_order_id)
        self.assertFalse(duplicated.project_project_id)
        self.assertEqual(len(duplicated.line_ids), 1)
        self.assertEqual(duplicated.line_ids.modele_id, modele)
        self.assertEqual(duplicated.line_ids.state, "draft")
        self.assertFalse(duplicated.line_ids.calculation_hash)

    def test_joinery_sale_order_uses_custom_portal_url(self):
        order = self.env["sale.order"].create(
            {
                "partner_id": self.env.ref("base.res_partner_1").id,
            }
        )
        self.assertEqual(order.get_joinery_portal_url(), f"/my/joinery/quotes/{order.id}")

    def test_detailed_result_xlsx_export_builds_valid_workbook(self):
        self._import_native_industrial_dataset()

        gamme = self.env["aluminium.joinery.gamme"].search([("code", "=", "masai")], limit=1)
        serie = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "masai_a_frappe"), ("gamme_id", "=", gamme.id)],
            limit=1,
        )
        modele = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "fenetre_fixe"), ("serie_id", "=", serie.id)],
            limit=1,
        )
        configuration = self.env["aluminium.joinery.configuration"].create(
            {"partner_id": self.env.ref("base.res_partner_1").id, "project_name": "Export XLSX"}
        )
        self.env["aluminium.joinery.configuration.line"].create(
            {
                "configuration_id": configuration.id,
                "gamme_id": gamme.id,
                "serie_id": serie.id,
                "modele_id": modele.id,
                "qty": 1,
                "width_mm": 1200,
                "height_mm": 1200,
            }
        )
        configuration.action_calculate()

        content = configuration._build_detailed_result_xlsx()

        self.assertTrue(content.startswith(b"PK"))
        archive = zipfile.ZipFile(BytesIO(content))
        self.assertEqual(
            set(archive.namelist()),
            {
                "[Content_Types].xml",
                "_rels/.rels",
                "docProps/app.xml",
                "docProps/core.xml",
                "xl/workbook.xml",
                "xl/_rels/workbook.xml.rels",
                "xl/styles.xml",
                "xl/worksheets/sheet1.xml",
            },
        )
        sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode()
        self.assertIn("Resultats detailles", sheet_xml)
        self.assertIn("GAMME : Masai", sheet_xml)
        self.assertIn("PROFILES", sheet_xml)
        self.assertIn("REMPLISSAGES", sheet_xml)

    def test_detailed_result_xlsx_export_action_uses_shared_download_route(self):
        self._import_native_industrial_dataset()

        gamme = self.env["aluminium.joinery.gamme"].search([("code", "=", "masai")], limit=1)
        serie = self.env["aluminium.joinery.serie"].search(
            [("code", "=", "masai_a_frappe"), ("gamme_id", "=", gamme.id)],
            limit=1,
        )
        modele = self.env["aluminium.joinery.modele"].search(
            [("code", "=", "fenetre_fixe"), ("serie_id", "=", serie.id)],
            limit=1,
        )
        configuration = self.env["aluminium.joinery.configuration"].create(
            {"partner_id": self.env.ref("base.res_partner_1").id}
        )
        self.env["aluminium.joinery.configuration.line"].create(
            {
                "configuration_id": configuration.id,
                "gamme_id": gamme.id,
                "serie_id": serie.id,
                "modele_id": modele.id,
                "qty": 1,
                "width_mm": 1200,
                "height_mm": 1200,
            }
        )
        configuration.action_calculate()

        action = configuration.action_export_detailed_results_xlsx()

        self.assertEqual(action["type"], "ir.actions.act_url")
        self.assertEqual(action["url"], configuration.get_detailed_results_xlsx_url())
