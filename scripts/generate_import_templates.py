from __future__ import annotations

import csv
import json
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ROOT_DIR = Path(__file__).resolve().parents[4]
MODULE_DIR = Path(__file__).resolve().parents[1]
SOURCE_XLSM = ROOT_DIR / "prologiciel" / "Version_MacOs.xlsm"
CSV_INVENTORY = ROOT_DIR / "prologiciel" / "vba_formula_inventory.csv"
OUTPUT_DIR = MODULE_DIR / "static" / "xls"


def col_to_num(col_letters: str) -> int:
    value = 0
    for char in col_letters:
        value = value * 26 + (ord(char) - 64)
    return value


def num_to_col(index: int) -> str:
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def split_ref(ref: str) -> tuple[int, int]:
    match = re.match(r"([A-Z]+)(\d+)$", ref)
    if not match:
        raise ValueError(f"Reference de cellule invalide: {ref}")
    return col_to_num(match.group(1)), int(match.group(2))


def parse_range(cell_range: str) -> tuple[tuple[int, int], tuple[int, int]]:
    start, end = cell_range.split(":")
    return split_ref(start), split_ref(end)


def normalize_bool(value) -> int:
    return 1 if value else 0


def derive_opening_type(serie_code: str) -> str:
    value = serie_code.lower()
    if "coulissant" in value:
        return "coulissant"
    if "frappe" in value:
        return "frappe"
    return ""


def derive_counts(modele_code: str) -> tuple[int | None, int | None]:
    lowered = modele_code.lower()
    panel_match = re.search(r"(\d+)[ _-]?vanta(?:il|ux)", lowered)
    rail_match = re.search(r"(\d+)[ _-]?rails?", lowered)
    panel_count = int(panel_match.group(1)) if panel_match else None
    rail_count = int(rail_match.group(1)) if rail_match else None
    return panel_count, rail_count


def slugify_code(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii").lower()
    normalized = re.sub(r"[^a-z0-9]+", "_", normalized)
    normalized = re.sub(r"_+", "_", normalized)
    return normalized.strip("_")


def build_modele_import_key(gamme: str, serie: str, modele: str) -> str:
    return f"{slugify_code(gamme)}__{slugify_code(serie)}__{slugify_code(modele)}"


def build_xmlid(prefix: str, *parts: str) -> str:
    suffix = "__".join(slugify_code(part) for part in parts if part)
    return f"ajc_{prefix}__{suffix}"


def build_example_rules() -> dict[str, list[list[object]]]:
    furio_gamme = slugify_code("Furio")
    furio_serie = slugify_code("Furio_coulissant")
    furio_modele = slugify_code("fenetre_2_vantaux_2_rails")
    flamingo_gamme = slugify_code("Flamingo")
    flamingo_serie = slugify_code("Flamingo_à_frappe")
    flamingo_modele = slugify_code("fenetre_2_vantaux_ouverture_française")

    article_names = {
        "012.042": "Dormant Furio",
        "012.117": "Traverse ouvrant Furio",
        "012.215": "Montant ouvrant Furio",
        "002.180": "Fermeture encastre",
        "002.225": "Equerre",
        "021.035": "Joint brosse",
        "021.045": "Joint brosse vertical",
        "013.210": "Ouvrant Flamingo",
        "019.059+019.061": "Profil composite 019.059 + 019.061",
        "019.059": "Composant 019.059",
        "019.061": "Composant 019.061",
    }

    articles = [
        [
            build_xmlid("product_tmpl", "012.042"),
            "012.042",
            article_names["012.042"],
            "product",
            "uom.product_uom_unit",
            "uom.product_uom_unit",
            0,
            1,
            "profile",
            5800,
            "dormant",
            0,
            "standard",
        ],
        [
            build_xmlid("product_tmpl", "012.117"),
            "012.117",
            article_names["012.117"],
            "product",
            "uom.product_uom_unit",
            "uom.product_uom_unit",
            0,
            1,
            "profile",
            5800,
            "traverse_ouvrant",
            0,
            "standard",
        ],
        [
            build_xmlid("product_tmpl", "012.215"),
            "012.215",
            article_names["012.215"],
            "product",
            "uom.product_uom_unit",
            "uom.product_uom_unit",
            0,
            1,
            "profile",
            5800,
            "montant_ouvrant",
            0,
            "standard",
        ],
        [
            build_xmlid("product_tmpl", "002.180"),
            "002.180",
            article_names["002.180"],
            "product",
            "uom.product_uom_unit",
            "uom.product_uom_unit",
            1,
            1,
            "accessory",
            "",
            "fermeture",
            0,
            "standard",
        ],
        [
            build_xmlid("product_tmpl", "002.225"),
            "002.225",
            article_names["002.225"],
            "product",
            "uom.product_uom_unit",
            "uom.product_uom_unit",
            1,
            1,
            "accessory",
            "",
            "equerre",
            0,
            "standard",
        ],
        [
            build_xmlid("product_tmpl", "021.035"),
            "021.035",
            article_names["021.035"],
            "product",
            "uom.product_uom_unit",
            "uom.product_uom_unit",
            1,
            1,
            "joint",
            "",
            "joint_brosse",
            0,
            "standard",
        ],
        [
            build_xmlid("product_tmpl", "021.045"),
            "021.045",
            article_names["021.045"],
            "product",
            "uom.product_uom_unit",
            "uom.product_uom_unit",
            1,
            1,
            "joint",
            "",
            "joint_vertical",
            0,
            "standard",
        ],
        [
            build_xmlid("product_tmpl", "013.210"),
            "013.210",
            article_names["013.210"],
            "product",
            "uom.product_uom_unit",
            "uom.product_uom_unit",
            0,
            1,
            "profile",
            5800,
            "ouvrant",
            0,
            "standard",
        ],
        [
            build_xmlid("product_tmpl", "019.059+019.061"),
            "019.059+019.061",
            article_names["019.059+019.061"],
            "product",
            "uom.product_uom_unit",
            "uom.product_uom_unit",
            1,
            1,
            "finished",
            "",
            "composite_profile",
            1,
            "manufactured_composite",
        ],
        [
            build_xmlid("product_tmpl", "019.059"),
            "019.059",
            article_names["019.059"],
            "product",
            "uom.product_uom_unit",
            "uom.product_uom_unit",
            0,
            1,
            "profile",
            5800,
            "component",
            0,
            "standard",
        ],
        [
            build_xmlid("product_tmpl", "019.061"),
            "019.061",
            article_names["019.061"],
            "product",
            "uom.product_uom_unit",
            "uom.product_uom_unit",
            0,
            1,
            "profile",
            5800,
            "component",
            0,
            "standard",
        ],
    ]

    profiles = [
        [
            furio_gamme,
            furio_serie,
            furio_modele,
            "fenetre_2_vantaux_2_rails",
            10,
            "Dormant",
            "profile",
            "012.042",
            "012.042",
            "Dormant",
            "dormant",
            "onglet",
            "perimeter",
            "length",
            1,
            2,
            0,
            1,
            2,
            0,
            1,
            0,
            "mm",
            1,
            "furio_dormant_perimeter",
            "direct_product",
        ],
        [
            "",
            "",
            "",
            "",
            20,
            "Traverse ouvrant",
            "profile",
            "012.117",
            "012.117",
            "Traverse ouvrant",
            "traverse_ouvrant",
            "droit",
            "linear_l",
            "length",
            4,
            1,
            -150.1,
            2,
            0,
            0,
            1,
            0,
            "mm",
            1,
            "furio_traverse_linear_l",
            "direct_product",
        ],
        [
            "",
            "",
            "",
            "",
            30,
            "Montant ouvrant",
            "profile",
            "012.215",
            "012.215",
            "Montant ouvrant",
            "montant_ouvrant",
            "droit",
            "linear_h",
            "length",
            2,
            0,
            0,
            1,
            1,
            -78.6,
            1,
            0,
            "mm",
            1,
            "furio_montant_linear_h",
            "direct_product",
        ],
        [
            flamingo_gamme,
            flamingo_serie,
            flamingo_modele,
            "fenetre_2_vantaux_ouverture_française",
            10,
            "Ouvrant",
            "profile",
            "013.210",
            "013.210",
            "Ouvrant",
            "ouvrant",
            "droit",
            "sum_h_l",
            "length",
            4,
            1,
            -41.8,
            2,
            1,
            -36.8,
            1,
            0,
            "mm",
            1,
            "flamingo_ouvrant_sum_h_l",
            "direct_product",
        ],
    ]

    accessories = [
        [
            furio_gamme,
            furio_serie,
            furio_modele,
            "fenetre_2_vantaux_2_rails",
            10,
            "Fermeture encastre",
            "accessoire",
            "002.180",
            "002.180",
            "Fermeture encastre",
            "qty_only",
            2,
            0,
            "unit",
            1,
            "furio_fermeture_qty",
            "direct_product",
        ],
        [
            "",
            "",
            "",
            "",
            20,
            "Equerre",
            "accessoire",
            "002.225",
            "002.225",
            "Equerre",
            "qty_only",
            4,
            0,
            "unit",
            1,
            "furio_equerre_qty",
            "direct_product",
        ],
    ]

    joints = [
        [
            furio_gamme,
            furio_serie,
            furio_modele,
            "fenetre_2_vantaux_2_rails",
            10,
            "Joint brosse perimetrique",
            "joint",
            "021.035",
            "021.035",
            "Joint brosse",
            "joint_combo",
            4,
            4,
            0,
            "mm",
            1,
            "furio_joint_brosse_combo",
            "direct_product",
        ],
        [
            "",
            "",
            "",
            "",
            20,
            "Joint brosse vertical",
            "joint",
            "021.045",
            "021.045",
            "Joint brosse vertical",
            "joint_combo",
            0,
            2,
            0,
            "mm",
            1,
            "furio_joint_vertical_combo",
            "direct_product",
        ],
    ]

    fillings = [
        [
            furio_gamme,
            furio_serie,
            furio_modele,
            "fenetre_2_vantaux_2_rails",
            10,
            "Vitrage standard",
            "fill_dim",
            0.5,
            0,
            -87.5,
            "fill_dim",
            0,
            1,
            -182,
            "qty_only",
            2,
            0,
            1,
            "furio_vitrage_standard",
        ],
    ]

    series_complete = [
        [
            build_xmlid("serie", "Furio_coulissant"),
            build_xmlid("gamme", "Furio"),
            furio_serie,
            "Furio_coulissant",
            "coulissant",
            1,
            build_xmlid("modele", "Furio_coulissant", "fenetre_2_vantaux_2_rails"),
            furio_modele,
            "fenetre_2_vantaux_2_rails",
            2,
            2,
            1,
            build_modele_import_key("Furio", "Furio_coulissant", "fenetre_2_vantaux_2_rails"),
            build_xmlid("rule", "Furio_coulissant", "fenetre_2_vantaux_2_rails", "Dormant"),
            "Dormant",
            "profile",
            article_names["012.042"],
            "012.042",
            "Dormant",
            "dormant",
            "onglet",
            "perimeter",
            "length",
            1,
            2,
            0,
            1,
            2,
            0,
            1,
            0,
            "mm",
            1,
            "furio_dormant_perimeter",
            "direct_product",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ],
        [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            build_xmlid("rule", "Furio_coulissant", "fenetre_2_vantaux_2_rails", "Traverse_ouvrant"),
            "Traverse ouvrant",
            "profile",
            article_names["012.117"],
            "012.117",
            "Traverse ouvrant",
            "traverse_ouvrant",
            "droit",
            "linear_l",
            "length",
            4,
            1,
            -150.1,
            2,
            0,
            0,
            1,
            0,
            "mm",
            1,
            "furio_traverse_linear_l",
            "direct_product",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ],
        [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            build_xmlid("rule", "Furio_coulissant", "fenetre_2_vantaux_2_rails", "Montant_ouvrant"),
            "Montant ouvrant",
            "profile",
            article_names["012.215"],
            "012.215",
            "Montant ouvrant",
            "montant_ouvrant",
            "droit",
            "linear_h",
            "length",
            2,
            0,
            0,
            1,
            1,
            -78.6,
            1,
            0,
            "mm",
            1,
            "furio_montant_linear_h",
            "direct_product",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ],
        [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            build_xmlid("rule", "Furio_coulissant", "fenetre_2_vantaux_2_rails", "Fermeture_encastre"),
            "Fermeture encastre",
            "accessoire",
            article_names["002.180"],
            "002.180",
            "Fermeture encastre",
            "",
            "",
            "qty_only",
            "qty",
            2,
            0,
            0,
            1,
            0,
            0,
            1,
            0,
            "unit",
            1,
            "furio_fermeture_qty",
            "direct_product",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ],
        [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            build_xmlid("rule", "Furio_coulissant", "fenetre_2_vantaux_2_rails", "Joint_brosse"),
            "Joint brosse perimetrique",
            "joint",
            article_names["021.035"],
            "021.035",
            "Joint brosse",
            "",
            "",
            "joint_combo",
            "length",
            1,
            4,
            0,
            1,
            4,
            0,
            1,
            0,
            "mm",
            1,
            "furio_joint_brosse_combo",
            "direct_product",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ],
        [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            build_xmlid("filling_rule", "Furio_coulissant", "fenetre_2_vantaux_2_rails", "Vitrage_standard"),
            "Vitrage standard",
            "fill_dim",
            0.5,
            0,
            -87.5,
            "fill_dim",
            0,
            1,
            -182,
            "qty_only",
            2,
            0,
            1,
            "furio_vitrage_standard",
        ],
        [
            build_xmlid("serie", "Flamingo_a_frappe"),
            build_xmlid("gamme", "Flamingo"),
            flamingo_serie,
            "Flamingo_à_frappe",
            "frappe",
            1,
            build_xmlid("modele", "Flamingo_a_frappe", "fenetre_2_vantaux_ouverture_francaise"),
            flamingo_modele,
            "fenetre_2_vantaux_ouverture_française",
            2,
            "",
            1,
            build_modele_import_key("Flamingo", "Flamingo_à_frappe", "fenetre_2_vantaux_ouverture_française"),
            build_xmlid("rule", "Flamingo_a_frappe", "fenetre_2_vantaux_ouverture_francaise", "Ouvrant"),
            "Ouvrant",
            "profile",
            article_names["013.210"],
            "013.210",
            "Ouvrant",
            "ouvrant",
            "droit",
            "sum_h_l",
            "length",
            4,
            1,
            -41.8,
            2,
            1,
            -36.8,
            1,
            0,
            "mm",
            1,
            "flamingo_ouvrant_sum_h_l",
            "direct_product",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ],
    ]
    series_complete = [row + [""] * (50 - len(row)) for row in series_complete]

    boms_complete = [
        [
            build_xmlid("bom", "019.059+019.061"),
            article_names["019.059+019.061"],
            "normal",
            1,
            "uom.product_uom_unit",
            "BOM_019_059_019_061",
            build_xmlid("bom_line", "019.059+019.061", "019.059"),
            article_names["019.059"],
            1,
            "uom.product_uom_unit",
            10,
        ],
        [
            "",
            "",
            "",
            "",
            "",
            "",
            build_xmlid("bom_line", "019.059+019.061", "019.061"),
            article_names["019.061"],
            1,
            "uom.product_uom_unit",
            20,
        ],
    ]

    enums = [
        ["group", "value", "label", "notes"],
        ["formula_family", "qty_only", "Quantite seule", "Q * multiplier + constant"],
        ["formula_family", "linear_l", "Lineaire largeur", "Q * multiplier * ((coef_l*L + offset_l) / divisor_l) + constant"],
        ["formula_family", "linear_h", "Lineaire hauteur", "Q * multiplier * ((coef_h*H + offset_h) / divisor_h) + constant"],
        ["formula_family", "sum_h_l", "Somme largeur + hauteur", "Combine L et H sur une seule ligne de regle"],
        ["formula_family", "perimeter", "Perimetre", "Semantique perimetre, stockage compact H + L"],
        ["formula_family", "joint_combo", "Joint combine", "Q * (a*L + b*H + c)"],
        ["formula_family", "fill_dim", "Dimension remplissage", "coef_l*L + coef_h*H + constant"],
        ["uom_kind", "unit", "Unite", "Accessoires et quantites discretes"],
        ["uom_kind", "mm", "Millimetre", "Profils et joints lineaires"],
        ["product_resolution_mode", "direct_product", "Produit direct", "La regle pointe un article standard"],
        ["product_resolution_mode", "manufactured_composite", "Composite manufacture", "La regle pointe un produit fini avec BOM"],
        ["manufacturing_mode", "standard", "Standard", "Article simple"],
        ["manufacturing_mode", "manufactured_composite", "Composite fabrique", "Produit fini a nomenclature"],
        ["joinery_item_type", "profile", "Profile", "Profil aluminium"],
        ["joinery_item_type", "accessory", "Accessoire", "Quincaillerie"],
        ["joinery_item_type", "joint", "Joint", "Joint lineaire"],
        ["joinery_item_type", "finished", "Produit fini", "Produit composite ou vendu"],
    ]

    return {
        "Article Names": article_names,
        "Articles": articles,
        "Profiles": profiles,
        "Accessories": accessories,
        "Joints": joints,
        "Fillings": fillings,
        "Series Complete": series_complete,
        "BOMs Complete": boms_complete,
        "Enums": enums,
    }


COMPOSITE_CODE_MAP = {
    "019.059 + 061": "019.059+019.061",
}


def normalize_default_code(value: str) -> str:
    cleaned = (value or "").strip()
    return COMPOSITE_CODE_MAP.get(cleaned, cleaned)


def to_display_label(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    text = text.replace("_", " ").replace("-", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text[:1].upper() + text[1:]


def category_from_element(element_type: str) -> str:
    return {
        "profil": "profile",
        "accessoire": "accessoire",
        "joint": "joint",
    }.get(element_type, "filling")


def item_type_from_element(element_type: str) -> str:
    return {
        "profil": "profile",
        "accessoire": "accessory",
        "joint": "joint",
        "remplissage": "filling",
    }.get(element_type, "finished")


def model_product_code(gamme: str, serie: str, modele: str) -> str:
    return f"CFG__{build_modele_import_key(gamme, serie, modele)}"


def model_product_xmlid(gamme: str, serie: str, modele: str) -> str:
    return build_xmlid("product_tmpl", "cfg", gamme, serie, modele)


def read_inventory_rows() -> list[dict[str, str]]:
    with CSV_INVENTORY.open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def build_inventory_rows(catalog: dict[str, object]) -> dict[str, list[list[object]]]:
    inventory_rows = read_inventory_rows()
    ordered_model_keys = [
        (gamme, serie, modele)
        for gamme in catalog["gammes"]
        for serie in catalog["series_by_gamme"].get(gamme, [])
        for modele in catalog["models_by_serie"].get(serie, [])
    ]
    inventory_model_keys = {
        (row["gamme"], row["serie"], row["modele"])
        for row in inventory_rows
    }
    ordered_model_keys = [key for key in ordered_model_keys if key in inventory_model_keys]
    remaining_model_keys = sorted(inventory_model_keys - set(ordered_model_keys))
    ordered_model_keys.extend(remaining_model_keys)

    rules_by_model = {}
    filling_by_model = {}
    article_meta = {}
    composite_codes = {normalize_default_code(code) for code in COMPOSITE_CODE_MAP.values()}

    for row in inventory_rows:
        normalized_code = normalize_default_code(row["default_code"])
        element_type = row["element_type"]
        if normalized_code:
            meta = article_meta.setdefault(
                normalized_code,
                {
                    "default_code": normalized_code,
                    "name": "",
                    "joinery_item_type": item_type_from_element(element_type),
                    "joinery_bar_length_mm": 5800 if element_type == "profil" else 0,
                    "uom_id": "uom.product_uom_meter" if element_type == "joint" else "uom.product_uom_unit",
                    "joinery_usage_role": slugify_code(row["article_label"] or normalized_code),
                    "is_joinery_composite": 0,
                    "manufacturing_mode": "standard",
                    "sale_ok": 1 if element_type == "accessoire" else 0,
                    "purchase_ok": 1,
                },
            )
            meta["name"] = meta["name"] or to_display_label(row["article_label"]) or normalized_code
            if normalized_code in composite_codes:
                meta["name"] = f"Profil composite {normalized_code.replace('+', ' + ')}"
                meta["joinery_item_type"] = "finished"
                meta["is_joinery_composite"] = 1
                meta["manufacturing_mode"] = "manufactured_composite"
                meta["sale_ok"] = 1

        model_key = (row["gamme"], row["serie"], row["modele"])
        if element_type == "remplissage":
            filling_group = filling_by_model.setdefault(model_key, {})
            filling_rule = filling_group.setdefault(row["row_index"], {})
            filling_rule[row["target_measure"]] = row
            continue

        rules_by_model.setdefault(model_key, []).append(row)

    for gamme, serie, modele in ordered_model_keys:
        code = model_product_code(gamme, serie, modele)
        article_meta[code] = {
            "default_code": code,
            "name": f"Configuration {to_display_label(modele)}",
            "joinery_item_type": "finished",
            "joinery_bar_length_mm": 0,
            "uom_id": "uom.product_uom_unit",
            "joinery_usage_role": "configured_joinery",
            "is_joinery_composite": 0,
            "manufacturing_mode": "standard",
            "sale_ok": 1,
            "purchase_ok": 0,
        }

    article_rows = [[
        "id",
        "default_code",
        "name",
        "type",
        "uom_id/id",
        "uom_po_id/id",
        "sale_ok",
        "purchase_ok",
        "joinery_item_type",
        "joinery_bar_length_mm",
        "joinery_usage_role",
        "is_joinery_composite",
        "is_placeholder_product_tmpl",
        "manufacturing_mode",
    ]]
    for code in sorted(article_meta):
        meta = article_meta[code]
        article_rows.append(
            [
                build_xmlid("product_tmpl", code),
                meta["default_code"],
                meta["name"],
                "consu",
                meta["uom_id"],
                meta["uom_id"],
                meta["sale_ok"],
                meta["purchase_ok"],
                meta["joinery_item_type"],
                meta["joinery_bar_length_mm"] or "",
                meta["joinery_usage_role"],
                meta["is_joinery_composite"],
                0,
                meta["manufacturing_mode"],
            ]
        )

    series_complete_rows = [[
        "id",
        "gamme_id/id",
        "code",
        "name",
        "opening_type",
        "active",
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
        "modele_ids/rule_ids/id",
        "modele_ids/rule_ids/sequence",
        "modele_ids/rule_ids/name",
        "modele_ids/rule_ids/category",
        "modele_ids/rule_ids/product_default_code",
        "modele_ids/rule_ids/ref_text",
        "modele_ids/rule_ids/designation_override",
        "modele_ids/rule_ids/profile_role",
        "modele_ids/rule_ids/cut_type",
        "modele_ids/rule_ids/formula_family",
        "modele_ids/rule_ids/target_measure",
        "modele_ids/rule_ids/multiplier",
        "modele_ids/rule_ids/coef_l",
        "modele_ids/rule_ids/offset_l",
        "modele_ids/rule_ids/divisor_l",
        "modele_ids/rule_ids/coef_h",
        "modele_ids/rule_ids/offset_h",
        "modele_ids/rule_ids/divisor_h",
        "modele_ids/rule_ids/constant",
        "modele_ids/rule_ids/uom_kind",
        "modele_ids/rule_ids/active",
        "modele_ids/rule_ids/rule_code",
        "modele_ids/rule_ids/product_resolution_mode",
        "modele_ids/filling_rule_ids/id",
        "modele_ids/filling_rule_ids/sequence",
        "modele_ids/filling_rule_ids/name",
        "modele_ids/filling_rule_ids/product_default_code",
        "modele_ids/filling_rule_ids/family_width",
        "modele_ids/filling_rule_ids/width_coef_l",
        "modele_ids/filling_rule_ids/width_coef_h",
        "modele_ids/filling_rule_ids/width_constant",
        "modele_ids/filling_rule_ids/family_height",
        "modele_ids/filling_rule_ids/height_coef_l",
        "modele_ids/filling_rule_ids/height_coef_h",
        "modele_ids/filling_rule_ids/height_constant",
        "modele_ids/filling_rule_ids/family_qty",
        "modele_ids/filling_rule_ids/qty_multiplier",
        "modele_ids/filling_rule_ids/qty_constant",
        "modele_ids/filling_rule_ids/active",
        "modele_ids/filling_rule_ids/rule_code",
    ]]
    rule_headers = [
        "id",
        "modele_id/id",
        "sequence",
        "name",
        "category",
        "product_id",
        "ref_text",
        "designation_override",
        "profile_role",
        "cut_type",
        "formula_family",
        "target_measure",
        "multiplier",
        "coef_l",
        "offset_l",
        "divisor_l",
        "coef_h",
        "offset_h",
        "divisor_h",
        "constant",
        "uom_kind",
        "active",
        "rule_code",
        "product_resolution_mode",
    ]
    profiles_rows = [rule_headers[:]]
    accessories_rows = [rule_headers[:]]
    joints_rows = [rule_headers[:]]
    rules_rows = [rule_headers[:]]
    filling_headers = [
        "id",
        "modele_id/id",
        "sequence",
        "name",
        "product_id",
        "family_width",
        "width_coef_l",
        "width_coef_h",
        "width_constant",
        "family_height",
        "height_coef_l",
        "height_coef_h",
        "height_constant",
        "family_qty",
        "qty_multiplier",
        "qty_constant",
        "active",
        "rule_code",
    ]
    fillings_rows = [filling_headers[:]]

    boms_complete_rows = [[
        "id",
        "product_default_code",
        "type",
        "product_qty",
        "product_uom_id/id",
        "code",
        "bom_line_ids/id",
        "bom_line_ids/component_default_code",
        "bom_line_ids/product_qty",
        "bom_line_ids/product_uom_id/id",
        "bom_line_ids/sequence",
    ]]
    bom_rows = [[
        "id",
        "product_default_code",
        "type",
        "product_qty",
        "product_uom_id/id",
        "code",
    ]]
    bom_line_rows = [[
        "id",
        "bom_id/id",
        "component_default_code",
        "product_qty",
        "product_uom_id/id",
        "sequence",
    ]]

    ordered_series = [(gamme, serie) for gamme in catalog["gammes"] for serie in catalog["series_by_gamme"].get(gamme, [])]
    for gamme, serie in ordered_series:
        model_names = [modele for current_gamme, current_serie, modele in ordered_model_keys if current_gamme == gamme and current_serie == serie]
        if not model_names:
            continue
        first_row_for_serie = True
        gamme_id = build_xmlid("gamme", gamme)
        serie_id = build_xmlid("serie", serie)
        for modele in model_names:
            model_key = (gamme, serie, modele)
            model_code = slugify_code(modele)
            panel_count, rail_count = derive_counts(modele)
            model_xmlid = build_xmlid("modele", serie, modele)
            finished_product_id = model_product_xmlid(gamme, serie, modele)
            rules = sorted(rules_by_model.get(model_key, []), key=lambda record: int(record["row_index"]))
            fillings = sorted(filling_by_model.get(model_key, {}).items(), key=lambda item: int(item[0]))
            payload_rows = []
            for rule_row in rules:
                payload_rows.append(("rule", build_rule_payload(rule_row)))
            for fill_index, fill_group in fillings:
                payload_rows.append(("filling", build_filling_payload(gamme, serie, modele, fill_index, fill_group)))
            if not payload_rows:
                payload_rows.append(("empty", {}))
            first_payload_for_model = True
            for payload_type, payload in payload_rows:
                row = [""] * len(series_complete_rows[0])
                if first_row_for_serie:
                    row[0] = serie_id
                    row[1] = gamme_id
                    row[2] = slugify_code(serie)
                    row[3] = serie
                    row[4] = derive_opening_type(serie)
                    row[5] = 1
                if first_payload_for_model:
                    row[6] = model_xmlid
                    row[7] = model_code
                    row[8] = modele
                    row[9] = panel_count or ""
                    row[10] = rail_count or ""
                    row[11] = 1
                    row[12] = build_modele_import_key(gamme, serie, modele)
                    row[13] = finished_product_id
                    row[14] = finished_product_id
                    row[15] = ""
                if payload_type == "rule":
                    rule_cells = [
                        payload["id"],
                        payload["sequence"],
                        payload["name"],
                        payload["category"],
                        payload["product_default_code"],
                        payload["ref_text"],
                        payload["designation_override"],
                        payload["profile_role"],
                        payload["cut_type"],
                        payload["formula_family"],
                        payload["target_measure"],
                        payload["multiplier"],
                        payload["coef_l"],
                        payload["offset_l"],
                        payload["divisor_l"],
                        payload["coef_h"],
                        payload["offset_h"],
                        payload["divisor_h"],
                        payload["constant"],
                        payload["uom_kind"],
                        1,
                        payload["rule_code"],
                        payload["product_resolution_mode"],
                    ]
                    row[16:39] = rule_cells
                elif payload_type == "filling":
                    fill_cells = [
                        payload["id"],
                        payload["sequence"],
                        payload["name"],
                        payload["product_default_code"],
                        payload["family_width"],
                        payload["width_coef_l"],
                        payload["width_coef_h"],
                        payload["width_constant"],
                        payload["family_height"],
                        payload["height_coef_l"],
                        payload["height_coef_h"],
                        payload["height_constant"],
                        payload["family_qty"],
                        payload["qty_multiplier"],
                        payload["qty_constant"],
                        1,
                        payload["rule_code"],
                    ]
                    row[39:56] = fill_cells
                series_complete_rows.append(row)
                first_row_for_serie = False
                first_payload_for_model = False

            for rule_row in rules:
                target = category_from_element(rule_row["element_type"])
                reference_row = build_rule_payload(rule_row)
                output = [
                    reference_row["id"],
                    model_xmlid,
                    reference_row["sequence"],
                    reference_row["name"],
                    reference_row["category"],
                    reference_row["product_default_code"],
                    reference_row["ref_text"],
                    reference_row["designation_override"],
                    reference_row["profile_role"],
                    reference_row["cut_type"],
                    reference_row["formula_family"],
                    reference_row["target_measure"],
                    reference_row["multiplier"],
                    reference_row["coef_l"],
                    reference_row["offset_l"],
                    reference_row["divisor_l"],
                    reference_row["coef_h"],
                    reference_row["offset_h"],
                    reference_row["divisor_h"],
                    reference_row["constant"],
                    reference_row["uom_kind"],
                    1,
                    reference_row["rule_code"],
                    reference_row["product_resolution_mode"],
                ]
                rules_rows.append(output)
                if target == "profile":
                    profiles_rows.append(output)
                elif target == "accessoire":
                    accessories_rows.append(output)
                else:
                    joints_rows.append(output)

            for fill_index, fill_group in fillings:
                fill_payload = build_filling_payload(gamme, serie, modele, fill_index, fill_group)
                fillings_rows.append(
                    [
                        fill_payload["id"],
                        model_xmlid,
                        fill_payload["sequence"],
                        fill_payload["name"],
                        fill_payload["product_default_code"],
                        fill_payload["family_width"],
                        fill_payload["width_coef_l"],
                        fill_payload["width_coef_h"],
                        fill_payload["width_constant"],
                        fill_payload["family_height"],
                        fill_payload["height_coef_l"],
                        fill_payload["height_coef_h"],
                        fill_payload["height_constant"],
                        fill_payload["family_qty"],
                        fill_payload["qty_multiplier"],
                        fill_payload["qty_constant"],
                        1,
                        fill_payload["rule_code"],
                    ]
                )

    composite_codes_in_inventory = sorted(code for code in article_meta if code in composite_codes)
    for code in composite_codes_in_inventory:
        components = code.split("+")
        bom_id = build_xmlid("bom", code)
        bom_rows.append(
            [bom_id, code, "normal", 1, "uom.product_uom_unit", f"BOM_{slugify_code(code)}"]
        )
        first_line = True
        for sequence, component in enumerate(components, start=1):
            line_id = build_xmlid("bom_line", code, component)
            bom_line = [
                line_id,
                bom_id,
                component,
                1,
                "uom.product_uom_unit",
                sequence * 10,
            ]
            bom_line_rows.append(bom_line)
            boms_complete_rows.append(
                [
                    bom_id if first_line else "",
                    code if first_line else "",
                    "normal" if first_line else "",
                    1 if first_line else "",
                    "uom.product_uom_unit" if first_line else "",
                    f"BOM_{slugify_code(code)}" if first_line else "",
                    line_id,
                    component,
                    1,
                    "uom.product_uom_unit",
                    sequence * 10,
                ]
            )
            first_line = False

    enums = [
        ["group", "value", "label", "notes"],
        ["formula_family", "qty_only", "Quantite seule", "Q * multiplier + constant"],
        ["formula_family", "linear_l", "Lineaire largeur", "Q * ((coef_l*L + offset_l) / divisor_l) + constant"],
        ["formula_family", "linear_h", "Lineaire hauteur", "Q * ((coef_h*H + offset_h) / divisor_h) + constant"],
        ["formula_family", "sum_h_l", "Somme largeur + hauteur", "Combine L et H sur une seule ligne de regle"],
        ["formula_family", "perimeter", "Perimetre", "Semantique perimetre, stockage compact H + L"],
        ["formula_family", "joint_combo", "Joint combine", "Q * ((coef_l*L) + (coef_h*H) + constant)"],
        ["formula_family", "fill_dim", "Dimension remplissage", "coef_l*L + coef_h*H + constant"],
        ["uom_kind", "unit", "Unite", "Quantites discretes"],
        ["uom_kind", "mm", "Millimetre", "Longueurs et dimensions"],
        ["product_resolution_mode", "direct_product", "Produit direct", "Resolution article par reference"],
        ["product_resolution_mode", "manufactured_composite", "Composite fabrique", "Produit fini avec nomenclature"],
        ["product_resolution_mode", "lookup_only", "Lookup only", "Aucun article resolu, reference libre"],
        ["manufacturing_mode", "standard", "Standard", "Article simple"],
        ["manufacturing_mode", "manufactured_composite", "Composite fabrique", "Produit fini a nomenclature"],
        ["joinery_item_type", "profile", "Profile", "Profil aluminium"],
        ["joinery_item_type", "accessory", "Accessoire", "Quincaillerie"],
        ["joinery_item_type", "joint", "Joint", "Joint lineaire"],
        ["joinery_item_type", "finished", "Produit fini", "Produit de modele ou composite"],
    ]

    return {
        "Articles": article_rows,
        "Series Complete": series_complete_rows,
        "Profiles": profiles_rows,
        "Accessories": accessories_rows,
        "Joints": joints_rows,
        "Rules": rules_rows,
        "Fillings": fillings_rows,
        "Filling Rules": fillings_rows,
        "BOMs Complete": boms_complete_rows,
        "BOMs": bom_rows,
        "BOM Lines": bom_line_rows,
        "Enums": enums,
    }


def build_rule_payload(row: dict[str, str]) -> dict[str, object]:
    params = json.loads(row["canonical_params_json"] or "{}")
    family = row["formula_family"]
    q_constant = float(params.get("q_constant") or 0.0)
    payload = {
        "id": build_xmlid("rule", row["serie"], row["modele"], row["row_index"], normalize_default_code(row["default_code"]) or row["article_label"]),
        "sequence": int(row["row_index"]),
        "name": to_display_label(row["article_label"]) or normalize_default_code(row["default_code"]) or "Regle",
        "category": category_from_element(row["element_type"]),
        "product_default_code": normalize_default_code(row["default_code"]),
        "ref_text": normalize_default_code(row["default_code"]) or row["article_label"],
        "designation_override": to_display_label(row["article_label"]) or normalize_default_code(row["default_code"]),
        "profile_role": slugify_code(row["article_label"]) if row["element_type"] == "profil" else "",
        "cut_type": row["cut_type"],
        "formula_family": family,
        "target_measure": row["target_measure"],
        "multiplier": 1,
        "coef_l": 0,
        "offset_l": 0,
        "divisor_l": 1,
        "coef_h": 0,
        "offset_h": 0,
        "divisor_h": 1,
        "constant": float(params.get("constant") or 0.0),
        "uom_kind": "mm" if row["target_measure"] in {"length", "width", "height"} else "unit",
        "rule_code": slugify_code(f"{row['gamme']} {row['serie']} {row['modele']} {row['row_index']} {normalize_default_code(row['default_code']) or row['article_label']}"),
        "product_resolution_mode": (
            "lookup_only"
            if not normalize_default_code(row["default_code"])
            else "manufactured_composite"
            if normalize_default_code(row["default_code"]) in COMPOSITE_CODE_MAP.values()
            else "direct_product"
        ),
    }
    if family == "qty_only":
        payload["multiplier"] = q_constant or 1
    elif family == "linear_l":
        payload["coef_l"] = float(params.get("coef_l") or 0.0)
        payload["offset_l"] = q_constant
    elif family == "linear_h":
        payload["coef_h"] = float(params.get("coef_h") or 0.0)
        payload["offset_h"] = q_constant
    elif family in {"sum_h_l", "perimeter"}:
        payload["coef_l"] = float(params.get("coef_l") or 0.0)
        payload["coef_h"] = float(params.get("coef_h") or 0.0)
        payload["offset_l"] = q_constant
    elif family == "joint_combo":
        payload["coef_l"] = float(params.get("coef_l") or 0.0)
        payload["coef_h"] = float(params.get("coef_h") or 0.0)
    return payload


def build_filling_payload(gamme: str, serie: str, modele: str, fill_index: str, fill_group: dict[str, dict[str, str]]) -> dict[str, object]:
    width = fill_group["width"]
    height = fill_group["height"]
    qty = fill_group["qty"]
    width_params = json.loads(width["canonical_params_json"] or "{}")
    height_params = json.loads(height["canonical_params_json"] or "{}")
    qty_params = json.loads(qty["canonical_params_json"] or "{}")
    product_default_code = (
        normalize_default_code(width.get("default_code"))
        or normalize_default_code(height.get("default_code"))
        or normalize_default_code(qty.get("default_code"))
    )
    return {
        "id": build_xmlid("filling_rule", serie, modele, fill_index),
        "sequence": int(fill_index),
        "name": f"Remplissage {fill_index}",
        "product_default_code": product_default_code,
        "family_width": width["formula_family"],
        "width_coef_l": float(width_params.get("coef_l") or 0.0),
        "width_coef_h": float(width_params.get("coef_h") or 0.0),
        "width_constant": float(width_params.get("constant") or 0.0),
        "family_height": height["formula_family"],
        "height_coef_l": float(height_params.get("coef_l") or 0.0),
        "height_coef_h": float(height_params.get("coef_h") or 0.0),
        "height_constant": float(height_params.get("constant") or 0.0),
        "family_qty": qty["formula_family"],
        "qty_multiplier": float(qty_params.get("q_constant") or 1.0),
        "qty_constant": float(qty_params.get("constant") or 0.0),
        "rule_code": slugify_code(f"{gamme} {serie} {modele} filling {fill_index}"),
    }


def get_catalog_from_xlsm(path: Path) -> dict[str, object]:
    ns = {"a": NS_MAIN, "r": NS_REL}
    with ZipFile(path) as zf:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for item in root.findall(f"{{{NS_MAIN}}}si"):
                parts = [node.text or "" for node in item.iter(f"{{{NS_MAIN}}}t")]
                shared_strings.append("".join(parts))

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        sheets = {}
        for sheet in workbook.find("a:sheets", ns):
            sheets[sheet.attrib[f"{{{NS_REL}}}id"]] = sheet.attrib["name"]

        relationships = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        sheet_targets = {}
        for rel in relationships:
            rel_id = rel.attrib["Id"]
            target = rel.attrib["Target"]
            if rel_id in sheets:
                sheet_targets[sheets[rel_id]] = f"xl/{target}" if not target.startswith("xl/") else target

        sheet_cache = {}

        def get_sheet_cells(sheet_name: str) -> dict[str, str]:
            if sheet_name in sheet_cache:
                return sheet_cache[sheet_name]
            xml_root = ET.fromstring(zf.read(sheet_targets[sheet_name]))
            values = {}
            for cell in xml_root.iter(f"{{{NS_MAIN}}}c"):
                ref = cell.attrib["r"]
                cell_type = cell.attrib.get("t")
                value_node = cell.find(f"{{{NS_MAIN}}}v")
                if value_node is None:
                    values[ref] = ""
                    continue
                value = value_node.text or ""
                if cell_type == "s":
                    value = shared_strings[int(value)]
                values[ref] = value
            sheet_cache[sheet_name] = values
            return values

        def defined_name_values(name: str) -> list[str]:
            for defined_name in workbook.find("a:definedNames", ns):
                if defined_name.attrib.get("name") != name:
                    continue
                text = (defined_name.text or "").strip()
                if "!" not in text:
                    return []
                sheet_name, cell_range = text.split("!", 1)
                sheet_name = sheet_name.strip("'")
                cell_range = cell_range.replace("$", "")
                cells = get_sheet_cells(sheet_name)
                if ":" not in cell_range:
                    value = cells.get(cell_range, "")
                    return [value] if value else []
                (c1, r1), (c2, r2) = parse_range(cell_range)
                result = []
                for row in range(r1, r2 + 1):
                    for col in range(c1, c2 + 1):
                        value = cells.get(f"{num_to_col(col)}{row}", "")
                        if value:
                            result.append(value)
                return result
            return []

        gammes = defined_name_values("Gamme")
        series_by_gamme = {gamme: defined_name_values(gamme) for gamme in gammes}
        models_by_serie = {}
        for series in series_by_gamme.values():
            for serie in series:
                models_by_serie[serie] = defined_name_values(serie)

    return {
        "gammes": gammes,
        "series_by_gamme": series_by_gamme,
        "models_by_serie": models_by_serie,
    }


def build_rows() -> dict[str, list[list[object]]]:
    catalog = get_catalog_from_xlsm(SOURCE_XLSM)
    gammes = catalog["gammes"]
    series_by_gamme = catalog["series_by_gamme"]
    models_by_serie = catalog["models_by_serie"]
    inventory = build_inventory_rows(catalog)

    instructions = [
        ["instruction", "details"],
        ["usage", "Utilisez l'import natif Odoo modele par modele. Les fichiers telechargeables du module sont deja alignes avec chaque modele."],
        ["ordre", "Ordre recommande: 1. Catalogue (gammes/séries/modèles)  2. Articles  3. Rules  4. Filling Rules  5. BOMs"],
        ["source", "Les gammes, series et modeles de cet exemplaire proviennent directement du classeur VBA."],
        ["important", "Les cellules parent vides signifient: meme parent que la ligne precedente."],
        ["regles", "Les regles restent orientees familles de formule + parametres numeriques, pas formules texte libres."],
        ["composites", "Une reference composite telle que 019.059 + 019.061 doit etre geree comme produit fini fabrique avec BOM."],
        ["ids", "Les feuilles utilisent des ids stables sur les parents et enfants pour permettre les mises a jour natives Odoo."],
        ["codes", "Les colonnes code sont des cles techniques ASCII stables ; les noms restent metier et lisibles."],
        ["reimport", "Les reimports natifs mettent a jour les enregistrements portant le meme External ID. Les suppressions restent a traiter explicitement si necessaire."],
        ["enums", "L'onglet Enums est un dictionnaire de reference et n'est pas importe comme modele Odoo."],
    ]

    gammes_rows = [[
        "id",
        "code",
        "name",
        "default_bar_length_mm",
        "active",
    ]]
    for gamme in gammes:
        gammes_rows.append(
            [
                build_xmlid("gamme", gamme),
                slugify_code(gamme),
                gamme,
                5800,
                1,
            ]
        )

    catalogue_rows = [[
        "id",
        "code",
        "name",
        "default_bar_length_mm",
        "active",
        "serie_ids/id",
        "serie_ids/code",
        "serie_ids/name",
        "serie_ids/opening_type",
        "serie_ids/active",
        "serie_ids/modele_ids/id",
        "serie_ids/modele_ids/code",
        "serie_ids/modele_ids/name",
        "serie_ids/modele_ids/panel_count",
        "serie_ids/modele_ids/rail_count",
        "serie_ids/modele_ids/active",
        "serie_ids/modele_ids/x_import_key",
    ]]
    inventory_model_keys = sorted(
        {
            (row["gamme"], row["serie"], row["modele"])
            for row in read_inventory_rows()
        }
    )
    inventory_model_set = set(inventory_model_keys)
    written_model_keys = set()
    for gamme in gammes:
        first_row_for_gamme = True
        gamme_code = slugify_code(gamme)
        for serie in series_by_gamme.get(gamme, []):
            first_row_for_serie = True
            serie_code = slugify_code(serie)
            for modele in models_by_serie.get(serie, []):
                if (gamme, serie, modele) not in inventory_model_set:
                    continue
                panel_count, rail_count = derive_counts(modele)
                modele_code = slugify_code(modele)
                catalogue_rows.append(
                    [
                        build_xmlid("gamme", gamme) if first_row_for_gamme else "",
                        gamme_code if first_row_for_gamme else "",
                        gamme if first_row_for_gamme else "",
                        5800 if first_row_for_gamme else "",
                        1 if first_row_for_gamme else "",
                        build_xmlid("serie", serie) if first_row_for_serie else "",
                        serie_code if first_row_for_serie else "",
                        serie if first_row_for_serie else "",
                        derive_opening_type(serie) if first_row_for_serie else "",
                        1 if first_row_for_serie else "",
                        build_xmlid("modele", serie, modele),
                        modele_code,
                        modele,
                        panel_count or "",
                        rail_count or "",
                        1,
                        build_modele_import_key(gamme, serie, modele),
                    ]
                )
                first_row_for_gamme = False
                first_row_for_serie = False
                written_model_keys.add((gamme, serie, modele))
    for gamme, serie, modele in inventory_model_keys:
        if (gamme, serie, modele) in written_model_keys:
            continue
        panel_count, rail_count = derive_counts(modele)
        catalogue_rows.append(
            [
                build_xmlid("gamme", gamme),
                slugify_code(gamme),
                gamme,
                5800,
                1,
                build_xmlid("serie", serie),
                slugify_code(serie),
                serie,
                derive_opening_type(serie),
                1,
                build_xmlid("modele", serie, modele),
                slugify_code(modele),
                modele,
                panel_count or "",
                rail_count or "",
                1,
                build_modele_import_key(gamme, serie, modele),
            ]
        )

    return {
        "Consignes": instructions,
        "Enums": inventory["Enums"],
        "Gammes": gammes_rows,
        "Catalogue": catalogue_rows,
        "Articles": inventory["Articles"],
        "Series Complete": inventory["Series Complete"],
        "Rules": inventory["Rules"],
        "Profiles": inventory["Profiles"],
        "Accessories": inventory["Accessories"],
        "Joints": inventory["Joints"],
        "Fillings": inventory["Fillings"],
        "Filling Rules": inventory["Filling Rules"],
        "BOMs Complete": inventory["BOMs Complete"],
        "BOMs": inventory["BOMs"],
        "BOM Lines": inventory["BOM Lines"],
    }


def xml_cell(value) -> tuple[str, bool]:
    if value is None:
        return "", False
    if isinstance(value, bool):
        return f"<v>{1 if value else 0}</v>", True
    if isinstance(value, (int, float)):
        return f"<v>{value}</v>", True
    text = str(value)
    if text == "":
        return "", False
    return f"<is><t>{escape(text)}</t></is>", False


def sheet_xml(rows: list[list[object]]) -> str:
    max_col = max(len(row) for row in rows)
    last_cell = f"{num_to_col(max_col)}{len(rows)}"
    row_xml = []
    for row_index, row in enumerate(rows, start=1):
        cells_xml = []
        for col_index, value in enumerate(row, start=1):
            payload, numeric = xml_cell(value)
            if not payload:
                continue
            cell_type = "" if numeric else ' t="inlineStr"'
            cells_xml.append(f'<c r="{num_to_col(col_index)}{row_index}"{cell_type}>{payload}</c>')
        row_xml.append(f'<row r="{row_index}">{"".join(cells_xml)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{NS_MAIN}">'
        f"<dimension ref=\"A1:{last_cell}\"/>"
        "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>"
        "<sheetFormatPr defaultRowHeight=\"15\"/>"
        f"<sheetData>{''.join(row_xml)}</sheetData>"
        "</worksheet>"
    )


def workbook_xml(sheet_names: list[str]) -> str:
    sheets_xml = []
    for index, name in enumerate(sheet_names, start=1):
        sheets_xml.append(
            f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">'
        "<bookViews><workbookView activeTab=\"0\"/></bookViews>"
        f"<sheets>{''.join(sheets_xml)}</sheets>"
        "</workbook>"
    )


def workbook_rels_xml(sheet_count: int) -> str:
    rels = []
    for index in range(1, sheet_count + 1):
        rels.append(
            '<Relationship '
            f'Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    rels.append(
        '<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'.format(sheet_count + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{''.join(rels)}"
        "</Relationships>"
    )


def root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def content_types_xml(sheet_count: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            '<Override PartName="/xl/worksheets/sheet{0}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(index)
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f"{''.join(overrides)}"
        "</Types>"
    )


def styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{NS_MAIN}">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )


def app_xml(sheet_names: list[str]) -> str:
    parts = "".join(f"<vt:lpstr>{escape(name)}</vt:lpstr>" for name in sheet_names)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Codex</Application>'
        f'<TitlesOfParts><vt:vector size="{len(sheet_names)}" baseType="lpstr">{parts}</vt:vector></TitlesOfParts>'
        f"<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant></vt:vector></HeadingPairs>"
        '</Properties>'
    )


def core_xml() -> str:
    created = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:creator>Codex</dc:creator>'
        '<cp:lastModifiedBy>Codex</cp:lastModifiedBy>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>'
        '</cp:coreProperties>'
    )


def write_workbook(path: Path, sheets: list[tuple[str, list[list[object]]]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    sheet_names = [name for name, _rows in sheets]
    with ZipFile(path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheets)))
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("docProps/app.xml", app_xml(sheet_names))
        zf.writestr("docProps/core.xml", core_xml())
        zf.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets)))
        zf.writestr("xl/styles.xml", styles_xml())
        for index, (_name, rows) in enumerate(sheets, start=1):
            zf.writestr(f"xl/worksheets/sheet{index}.xml", sheet_xml(rows))


def main() -> None:
    rows = build_rows()
    native_bundle_sheets = [
        ("Consignes", rows["Consignes"]),
        ("Enums", rows["Enums"]),
        ("Catalogue", rows["Catalogue"]),
        ("Articles", rows["Articles"]),
        ("Rules", rows["Rules"]),
        ("Filling Rules", rows["Filling Rules"]),
        ("BOMs", rows["BOMs Complete"]),
    ]
    reference_sheets = [
        ("Consignes", rows["Consignes"]),
        ("Enums", rows["Enums"]),
        ("Gammes", rows["Gammes"]),
        ("Articles", rows["Articles"]),
        ("Catalogue", rows["Catalogue"]),
        ("Rules", rows["Rules"]),
        ("Filling Rules", rows["Filling Rules"]),
        ("Profiles", rows["Profiles"]),
        ("Accessories", rows["Accessories"]),
        ("Joints", rows["Joints"]),
        ("Fillings", rows["Fillings"]),
        ("Series Complete", rows["Series Complete"]),
        ("BOMs Complete", rows["BOMs Complete"]),
        ("BOMs", rows["BOMs"]),
        ("BOM Lines", rows["BOM Lines"]),
    ]
    write_workbook(
        OUTPUT_DIR / "catalogue_vba_exemple.xlsx",
        reference_sheets,
    )
    write_workbook(
        OUTPUT_DIR / "catalogue_industriel_menuiserie.xlsx",
        native_bundle_sheets,
    )
    write_workbook(
        OUTPUT_DIR / "catalogue_hierarchie_import.xlsx",
        [("Catalogue", rows["Catalogue"])],
    )
    write_workbook(
        OUTPUT_DIR / "gammes_import.xlsx",
        [("Gammes", rows["Gammes"])],
    )
    write_workbook(
        OUTPUT_DIR / "articles_import.xlsx",
        [("Articles", rows["Articles"])],
    )
    write_workbook(
        OUTPUT_DIR / "enums_import.xlsx",
        [("Enums", rows["Enums"])],
    )
    write_workbook(
        OUTPUT_DIR / "rules_import.xlsx",
        [("Rules", rows["Rules"])],
    )
    write_workbook(
        OUTPUT_DIR / "profiles_import.xlsx",
        [("Profiles", rows["Profiles"])],
    )
    write_workbook(
        OUTPUT_DIR / "accessories_import.xlsx",
        [("Accessories", rows["Accessories"])],
    )
    write_workbook(
        OUTPUT_DIR / "joints_import.xlsx",
        [("Joints", rows["Joints"])],
    )
    write_workbook(
        OUTPUT_DIR / "fillings_import.xlsx",
        [("Filling Rules", rows["Filling Rules"])],
    )
    write_workbook(
        OUTPUT_DIR / "boms_import.xlsx",
        [("BOMs", rows["BOMs Complete"])],
    )
    write_workbook(
        OUTPUT_DIR / "boms_complete_import.xlsx",
        [("BOMs Complete", rows["BOMs Complete"])],
    )
    write_workbook(
        OUTPUT_DIR / "bom_lines_import.xlsx",
        [("BOM Lines", rows["BOM Lines"])],
    )
    write_workbook(
        OUTPUT_DIR / "series_complete_import.xlsx",
        [("Series Complete", rows["Series Complete"])],
    )
    write_workbook(
        OUTPUT_DIR / "series_import.xlsx",
        [("Series Complete", rows["Series Complete"])],
    )
    write_workbook(
        OUTPUT_DIR / "modeles_import.xlsx",
        [("Series Complete", rows["Series Complete"])],
    )
    write_workbook(
        OUTPUT_DIR / "regles_import.xlsx",
        [("Rules", rows["Rules"])],
    )


if __name__ == "__main__":
    main()
