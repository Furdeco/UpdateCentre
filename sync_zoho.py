from __future__ import annotations

import csv
import html
import json
import re
import sys
import xml.etree.ElementTree as ET
from datetime import date, datetime, time
from pathlib import Path

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.styles.colors import COLOR_INDEX

SOURCE = Path("/tmp/zoho-update-centre.xlsx")
JSON_OUT = Path("updates.json")
CSV_OUT = Path("updates.csv")

BODY_HEADERS = {
    "body", "updatebody", "updatecontent", "content", "updateinformation",
    "details", "description", "message", "instructions", "notes",
    "announcement", "information", "text", "update"
}

def norm(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())

def plain_value(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, (date, time)):
        return value.isoformat()
    if isinstance(value, CellRichText):
        return "".join(part.text if isinstance(part, TextBlock) else str(part) for part in value)
    return str(value)

def theme_palette(workbook) -> dict[int, str]:
    theme = getattr(workbook, "loaded_theme", None)
    if not theme:
        return {}
    try:
        root = ET.fromstring(theme)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        scheme = root.find(".//a:clrScheme", ns)
        names = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2",
                 "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
        result = {}
        if scheme is not None:
            for idx, name in enumerate(names):
                node = scheme.find(f"a:{name}", ns)
                if node is not None and len(node):
                    child = node[0]
                    value = child.attrib.get("lastClr") or child.attrib.get("val")
                    if value and re.fullmatch(r"[0-9A-Fa-f]{6}", value):
                        result[idx] = "#" + value
        return result
    except Exception:
        return {}

def tint_hex(hex_color: str, tint: float) -> str:
    try:
        rgb = [int(hex_color[i:i+2], 16) for i in (1, 3, 5)]
        adjusted = []
        for c in rgb:
            if tint < 0:
                adjusted.append(round(c * (1 + tint)))
            else:
                adjusted.append(round(c * (1 - tint) + 255 * tint))
        return "#" + "".join(f"{max(0, min(255, c)):02X}" for c in adjusted)
    except Exception:
        return hex_color

def color_css(color, palette: dict[int, str]) -> str | None:
    if color is None:
        return None
    try:
        if color.type == "rgb" and color.rgb:
            raw = str(color.rgb)
            if len(raw) == 8:
                raw = raw[-6:]
            if re.fullmatch(r"[0-9A-Fa-f]{6}", raw):
                if raw.upper() in {"000000", "010101"}:
                    return None
                return "#" + raw
        if color.type == "theme" and color.theme in palette:
            value = palette[color.theme]
            tint = getattr(color, "tint", 0) or 0
            return tint_hex(value, tint) if tint else value
        if color.type == "indexed" and color.indexed is not None:
            raw = COLOR_INDEX[color.indexed]
            if raw and len(raw) == 8:
                return "#" + raw[-6:]
    except Exception:
        pass
    return None

def font_css(font, fallback, palette: dict[int, str]) -> str:
    if font is None:
        font = fallback
    rules = []
    def merged_attr(name):
        value = getattr(font, name, None)
        return getattr(fallback, name, None) if value is None and fallback is not None else value

    bold = merged_attr("b")
    italic = merged_attr("i")
    underline = merged_attr("u")
    strike = merged_attr("strike")
    size = merged_attr("sz")
    family = getattr(font, "rFont", None) or getattr(font, "name", None)
    if not family and fallback is not None:
        family = getattr(fallback, "name", None)
    color = color_css(getattr(font, "color", None), palette)
    if not color and fallback is not None:
        color = color_css(getattr(fallback, "color", None), palette)

    if bold:
        rules.append("font-weight:700")
    if italic:
        rules.append("font-style:italic")
    decorations = []
    if underline:
        decorations.append("underline")
    if strike:
        decorations.append("line-through")
    if decorations:
        rules.append("text-decoration:" + " ".join(decorations))
    if color:
        rules.append("color:" + color)
    if size:
        try:
            rules.append(f"font-size:{float(size):g}pt")
        except (TypeError, ValueError):
            pass
    if family:
        family = str(family).replace("\\", "").replace('"', "").replace("'", "")
        if re.fullmatch(r"[\w\s,-]{1,100}", family):
            rules.append('font-family:"' + family + '"')
    return ";".join(rules)

def escaped_text(value: object) -> str:
    return html.escape(str(value), quote=False).replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")

def rich_cell_html(cell, palette: dict[int, str]) -> str:
    value = cell.value
    if value is None:
        return ""
    fallback = cell.font
    if isinstance(value, CellRichText):
        pieces = []
        for part in value:
            if isinstance(part, TextBlock):
                run_font = part.font
                text = part.text
            else:
                run_font = fallback
                text = str(part)
            style = font_css(run_font, fallback, palette)
            content = escaped_text(text)
            pieces.append(f'<span style="{html.escape(style, quote=True)}">{content}</span>' if style else content)
        return "".join(pieces)
    text = plain_value(value)
    style = font_css(fallback, fallback, palette)
    content = escaped_text(text)
    return f'<span style="{html.escape(style, quote=True)}">{content}</span>' if style else content

def main() -> None:
    if not SOURCE.exists():
        raise SystemExit(f"Missing downloaded workbook: {SOURCE}")
    if not SOURCE.read_bytes().startswith(b"PK"):
        raise SystemExit("Downloaded Zoho response is not a valid XLSX file.")

    # rich_text=True is essential: without it, openpyxl flattens rich-text runs.
    wb = openpyxl.load_workbook(SOURCE, data_only=True, rich_text=True, read_only=False)
    if not wb.sheetnames:
        raise SystemExit("The downloaded workbook contains no worksheets.")
    ws = wb[wb.sheetnames[0]]
    palette = theme_palette(wb)

    header_cells = list(ws[1])
    headers = [plain_value(c.value).strip() for c in header_cells]
    if not any(headers):
        raise SystemExit("The first worksheet row does not contain column headers.")
    body_idx = next((i for i, h in enumerate(headers) if norm(h) in BODY_HEADERS), None)
    if body_idx is None:
        raise SystemExit("Could not find the update body/content column in the worksheet headers.")

    records = []
    csv_rows = [headers]
    rich_cells = 0
    for row in ws.iter_rows(min_row=2):
        values = []
        record = {}
        for idx, cell in enumerate(row[:len(headers)]):
            if idx == body_idx:
                value = rich_cell_html(cell, palette)
                if isinstance(cell.value, CellRichText):
                    rich_cells += 1
                plain = plain_value(cell.value)
            else:
                value = plain_value(cell.value)
                plain = value
            values.append(plain)
            record[headers[idx]] = value
        while values and not values[-1].strip():
            values.pop()
        if not any(v.strip() for v in values):
            continue
        records.append(record)
        csv_rows.append(values)

    wb.close()
    if not records:
        raise SystemExit("No non-empty update records were found; refusing to overwrite published data.")

    JSON_OUT.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    with CSV_OUT.open("w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(csv_rows)

    check = json.loads(JSON_OUT.read_text(encoding="utf-8"))
    html_values = [str(v) for row in check for k, v in row.items() if norm(k) in BODY_HEADERS]
    tags = sum(v.count("<span") for v in html_values)
    print(f"Exported {len(records)} updates to {JSON_OUT} and {CSV_OUT}.")
    print(f"Preserved rich-text cells: {rich_cells}; styled HTML spans: {tags}.")
    if rich_cells and tags == 0:
        raise SystemExit("Rich-text cells were detected but no styled spans were generated.")

if __name__ == "__main__":
    main()
