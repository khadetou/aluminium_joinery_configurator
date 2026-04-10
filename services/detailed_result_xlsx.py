from __future__ import annotations

from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


def _column_letter(index: int) -> str:
    letter = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


class DetailedResultXlsxBuilder:
    SHEET_NAME = "Resultats detailles"
    COLUMN_WIDTHS = [20, 34, 22, 22, 12, 24, 18]

    def __init__(self, payload: dict):
        self.payload = payload
        self.max_cols = 7
        self.rows: list[dict] = []
        self.merges: list[tuple[int, int, int, int]] = []
        self.current_row = 1

    def build(self) -> bytes:
        self._populate_rows()
        output = BytesIO()
        with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", self._content_types_xml())
            archive.writestr("_rels/.rels", self._root_rels_xml())
            archive.writestr("docProps/app.xml", self._app_xml())
            archive.writestr("docProps/core.xml", self._core_xml())
            archive.writestr("xl/workbook.xml", self._workbook_xml())
            archive.writestr("xl/_rels/workbook.xml.rels", self._workbook_rels_xml())
            archive.writestr("xl/styles.xml", self._styles_xml())
            archive.writestr("xl/worksheets/sheet1.xml", self._sheet_xml())
        return output.getvalue()

    def _populate_rows(self):
        config_name = self.payload.get("configuration_name") or "-"
        title = "Resultats detailles"
        subtitle_parts = [config_name]
        if self.payload.get("project_name"):
            subtitle_parts.append(self.payload["project_name"])
        if self.payload.get("partner_name"):
            subtitle_parts.append(self.payload["partner_name"])
        self._add_merged_text_row(title, style=1, height=24)
        self._add_merged_text_row(" - ".join(part for part in subtitle_parts if part), style=2)
        self._blank_row()

        self._add_meta_row(
            ("Configuration", config_name),
            ("Projet", self.payload.get("project_name") or "-"),
            ("Client", self.payload.get("partner_name") or "-"),
        )
        self._add_meta_row(
            ("Date", self.payload.get("date") or "-"),
            ("Statut", self.payload.get("state_label") or "-"),
            ("Lignes", self.payload.get("line_count") or "0"),
        )
        self._blank_row()

        for group in self.payload.get("groups", []):
            self._add_merged_text_row(f"GAMME : {group['gamme_name']}", style=5)
            for line in group.get("lines", []):
                self._blank_row()
                self._add_merged_text_row(line["label"], style=6)
                self._add_meta_row(
                    ("Serie", line.get("serie_name") or "-"),
                    ("Modele", line.get("modele_name") or "-"),
                    ("Quantite", line.get("qty") or "-"),
                )
                self._add_meta_row(
                    ("Largeur (mm)", line.get("width_mm") or "-"),
                    ("Hauteur (mm)", line.get("height_mm") or "-"),
                    ("", ""),
                )
                self._blank_row()
                for section in line.get("sections", []):
                    if not section.get("rows"):
                        continue
                    self._add_merged_text_row(section["title"], style=7)
                    self._add_table_header(section.get("headers") or [])
                    for row in section.get("rows") or []:
                        self._add_table_row(row)
                    if section.get("totals"):
                        self._add_table_row(section["totals"], is_total=True)
                    self._blank_row()
            self._blank_row()

    def _blank_row(self):
        self.rows.append({"index": self.current_row, "cells": [], "height": 8})
        self.current_row += 1

    def _add_merged_text_row(self, value, style, height=20):
        self.rows.append(
            {
                "index": self.current_row,
                "cells": [{"col": 1, "value": value or "", "style": style, "kind": "string"}],
                "height": height,
            }
        )
        self.merges.append((self.current_row, 1, self.current_row, self.max_cols))
        self.current_row += 1

    def _add_meta_row(self, first, second, third):
        cells = []
        for offset, (label, value) in enumerate((first, second, third)):
            base_col = (offset * 2) + 1
            if label:
                cells.append({"col": base_col, "value": label, "style": 3, "kind": "string"})
                cells.append({"col": base_col + 1, "value": value or "-", "style": 4, "kind": "string"})
        self.rows.append({"index": self.current_row, "cells": cells, "height": 18})
        self.current_row += 1

    def _add_table_header(self, headers):
        cells = [
            {"col": index + 1, "value": header, "style": 8, "kind": "string"}
            for index, header in enumerate(headers[: self.max_cols])
        ]
        self.rows.append({"index": self.current_row, "cells": cells, "height": 20})
        self.current_row += 1

    def _add_table_row(self, row, is_total=False):
        cells = []
        for index, cell in enumerate(row[: self.max_cols], start=1):
            value = cell.get("value", "")
            css_class = cell.get("class") or ""
            is_numeric = "text-end" in css_class and self._is_number(value)
            style = 12 if is_total and is_numeric else 11 if is_total else 10 if is_numeric else 9
            kind = "number" if is_numeric else "string"
            cells.append({"col": index, "value": value, "style": style, "kind": kind})
        self.rows.append({"index": self.current_row, "cells": cells, "height": 18})
        self.current_row += 1

    def _is_number(self, value) -> bool:
        try:
            float(value)
            return True
        except (TypeError, ValueError):
            return False

    def _sheet_xml(self) -> str:
        cols_xml = "".join(
            f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>'
            for index, width in enumerate(self.COLUMN_WIDTHS, start=1)
        )
        rows_xml = []
        for row in self.rows:
            attrs = [f'r="{row["index"]}"']
            if row.get("height"):
                attrs.append(f'ht="{row["height"]}"')
                attrs.append('customHeight="1"')
            cells_xml = "".join(self._cell_xml(row["index"], cell) for cell in row["cells"])
            rows_xml.append(f'<row {" ".join(attrs)}>{cells_xml}</row>')
        merge_xml = ""
        if self.merges:
            merge_refs = "".join(
                f'<mergeCell ref="{_column_letter(start_col)}{start_row}:{_column_letter(end_col)}{end_row}"/>'
                for start_row, start_col, end_row, end_col in self.merges
            )
            merge_xml = f'<mergeCells count="{len(self.merges)}">{merge_refs}</mergeCells>'
        last_row = max((row["index"] for row in self.rows), default=1)
        dimension = f"A1:{_column_letter(self.max_cols)}{last_row}"
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="{dimension}"/>'
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'<cols>{cols_xml}</cols>'
            f'<sheetData>{"".join(rows_xml)}</sheetData>'
            f"{merge_xml}"
            "</worksheet>"
        )

    def _cell_xml(self, row_index: int, cell: dict) -> str:
        ref = f"{_column_letter(cell['col'])}{row_index}"
        style = cell["style"]
        value = cell["value"] or ""
        if cell["kind"] == "number":
            return f'<c r="{ref}" s="{style}"><v>{float(value)}</v></c>'
        text = escape(str(value))
        return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'

    def _content_types_xml(self) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/docProps/core.xml" '
            'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            '</Types>'
        )

    def _root_rels_xml(self) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
            'Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
            'Target="docProps/app.xml"/>'
            '</Relationships>'
        )

    def _workbook_xml(self) -> str:
        name = escape(self.SHEET_NAME)
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/></bookViews>'
            f'<sheets><sheet name="{name}" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>'
        )

    def _workbook_rels_xml(self) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
            '</Relationships>'
        )

    def _styles_xml(self) -> str:
        return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="5">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="16"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
    <font><sz val="11"/><color rgb="FF6B7280"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF5F7FA"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF163A70"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEFF3F8"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color auto="1"/></left>
      <right style="thin"><color auto="1"/></right>
      <top style="thin"><color auto="1"/></top>
      <bottom style="thin"><color auto="1"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color auto="1"/></left>
      <right style="thin"><color auto="1"/></right>
      <top style="medium"><color auto="1"/></top>
      <bottom style="thin"><color auto="1"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="13">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="5" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="5" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"""

    def _app_xml(self) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>Odoo Aluminium Joinery</Application>'
            '</Properties>'
        )

    def _core_xml(self) -> str:
        title = escape(self.payload.get("configuration_name") or self.SHEET_NAME)
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{title}</dc:title>'
            '<dc:creator>Odoo</dc:creator>'
            '<cp:lastModifiedBy>Odoo</cp:lastModifiedBy>'
            '</cp:coreProperties>'
        )
