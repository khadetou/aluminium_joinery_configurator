from __future__ import annotations

from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _col_to_index(cell_ref: str) -> int:
    letters = []
    for char in cell_ref:
        if char.isalpha():
            letters.append(char)
        else:
            break
    value = 0
    for char in letters:
        value = value * 26 + (ord(char.upper()) - 64)
    return value - 1


def _cell_value(cell, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value_node = cell.find(f"{{{NS_MAIN}}}v")
    if value_node is None:
        inline = cell.find(f"{{{NS_MAIN}}}is")
        if inline is None:
            return ""
        return "".join(node.text or "" for node in inline.iter(f"{{{NS_MAIN}}}t"))
    value = value_node.text or ""
    if cell_type == "s":
        return shared_strings[int(value)]
    return value


def read_xlsx_sheets(binary_content: bytes) -> dict[str, list[dict[str, str]]]:
    ns = {"a": NS_MAIN, "r": NS_REL}
    with ZipFile(BytesIO(binary_content)) as zf:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            shared_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for item in shared_root.findall("a:si", ns):
                shared_strings.append("".join(node.text or "" for node in item.iter(f"{{{NS_MAIN}}}t")))

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        relationships = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in relationships}

        result = {}
        for sheet in workbook.find("a:sheets", ns):
            rel_id = sheet.attrib[f"{{{NS_REL}}}id"]
            target = targets[rel_id]
            if not target.startswith("xl/"):
                target = f"xl/{target}"
            root = ET.fromstring(zf.read(target))
            rows = []
            for row in root.findall(".//a:sheetData/a:row", ns):
                cells = []
                for cell in row.findall("a:c", ns):
                    col_index = _col_to_index(cell.attrib["r"])
                    while len(cells) <= col_index:
                        cells.append("")
                    cells[col_index] = _cell_value(cell, shared_strings).strip()
                rows.append(cells)
            result[sheet.attrib["name"]] = rows

        for name, rows in list(result.items()):
            if not rows:
                result[name] = []
                continue
            headers = rows[0]
            parsed_rows = []
            for values in rows[1:]:
                padded = values + [""] * (len(headers) - len(values))
                parsed_rows.append({header: padded[index] for index, header in enumerate(headers) if header})
            result[name] = parsed_rows
        return result
