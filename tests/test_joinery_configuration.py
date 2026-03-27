import importlib.util
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
        type_index = rows["Articles"][0].index("type")
        self.assertEqual(rows["Articles"][1][type_index], "consu")

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
        self.assertTrue(configuration.summary_ids.filtered(lambda line: line.category == "filling"))
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
