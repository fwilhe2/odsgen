#!/usr/bin/env python3
"""Write `functions/`: one recalculated flat ODS document per category of `function_samples`.

Each document has the category's data sheets and a `Samples` sheet, one row per call:

    A  the function    B  the formula as typed    C  the formula    D  what it shows

Column C is written with no cached value and LibreOffice recalculates it, so the value saved
beside each formula is LibreOffice's answer — which makes the documents a test harness for
another evaluator as well as samples a person can open. Column B is the formula as
LibreOffice's formula bar shows it, read back from LibreOffice's own output in a first pass.

Beside the documents:

- `samples.json` — every cell of every document as LibreOffice saved it, each formula with
  its OpenFormula text, its Excel spelling (where Excel has one) and its value, for a
  generator of `.xlsx` versions such as ooxmlgen's `functions` family;
- `README.md` — the samples as a table per category, the documentation half.

Needs LibreOffice (`soffice`) on PATH. Settings every document carries, because they change
answers: wildcards on and regular expressions off in criteria, and case-insensitive text
comparison.
"""

import json
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from html import escape
from pathlib import Path

from function_samples import CATEGORIES, D, F, Tm

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "functions"

NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    "of": "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
    "calcext": "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
}

HEADINGS = ["Function", "Formula", "Result", "What it shows"]


def q(value):
    return escape(str(value), quote=True)


def column_letters(index):
    letters = ""
    index += 1
    while index:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


# --- writing ---------------------------------------------------------------------------

def xml_cell(value, style=None):
    attrs = f' table:style-name="{style}"' if style else ""
    if value is None:
        return f"<table:table-cell{attrs}/>"
    if isinstance(value, F):
        return f'<table:table-cell{attrs} table:formula="of:{q(value.formula)}"/>'
    if isinstance(value, bool):
        word = "true" if value else "false"
        return (f'<table:table-cell{attrs} office:value-type="boolean" office:boolean-value="{word}">'
                f"<text:p>{word.upper()}</text:p></table:table-cell>")
    if isinstance(value, (int, float)):
        return (f'<table:table-cell{attrs} office:value-type="float" office:value="{value!r}">'
                f"<text:p>{value}</text:p></table:table-cell>")
    if isinstance(value, D):
        return (f'<table:table-cell table:style-name="date" office:value-type="date" '
                f'office:date-value="{value.iso}"><text:p>{value.iso}</text:p></table:table-cell>')
    if isinstance(value, Tm):
        h, m, s = value.hms.split(":")
        return (f'<table:table-cell table:style-name="time" office:value-type="time" '
                f'office:time-value="PT{h}H{m}M{s}S"><text:p>{value.hms}</text:p></table:table-cell>')
    return (f'<table:table-cell{attrs} office:value-type="string">'
            f"<text:p>{q(value)}</text:p></table:table-cell>")


def xml_table(name, rows, hidden=None, widths=None):
    hidden = hidden or {}
    columns = "".join(f'<table:table-column table:style-name="{w}"/>' for w in widths or [])
    if not columns:
        count = max(len(row) for row in rows)
        columns = f'<table:table-column table:style-name="wide" table:number-columns-repeated="{count}"/>'
    body = []
    for number, row in enumerate(rows, start=1):
        visibility = f' table:visibility="{hidden[number]}"' if number in hidden else ""
        style = "heading" if number == 1 else None
        cells = "".join(xml_cell(value, style) for value in row)
        body.append(f"<table:table-row{visibility}>{cells}</table:table-row>")
    return f'<table:table table:name="{q(name)}">{columns}{"".join(body)}</table:table>'


def database_ranges(category):
    ranges = []
    for index, sheet in enumerate(category.sheets):
        if not sheet.filter:
            continue
        last_row, field, keep = sheet.filter
        last_col = column_letters(len(sheet.rows[0]) - 1)
        items = "".join(f'<table:filter-set-item table:value="{q(v)}"/>' for v in keep)
        ranges.append(
            f'<table:database-range table:name="__Anonymous_Sheet_DB__{index}" '
            f'table:target-range-address="{sheet.name}.A1:{sheet.name}.{last_col}{last_row}" '
            f'table:display-filter-buttons="true"><table:filter>'
            f'<table:filter-condition table:field-number="{field}" table:value="{q(keep[0])}" '
            f'table:operator="=">{items}</table:filter-condition></table:filter></table:database-range>')
    return f"<table:database-ranges>{''.join(ranges)}</table:database-ranges>" if ranges else ""


def flat_document(category, typed):
    """The document to hand LibreOffice. `typed` is column B, one text per sample."""
    samples = [HEADINGS]
    for (function, formula, what), text in zip(category.samples, typed):
        samples.append([function, text, F(formula), what])
    tables = [xml_table(s.name, s.rows, s.hidden) for s in category.sheets]
    tables.append(xml_table("Samples", samples, widths=["narrow", "formula", "result", "formula"]))
    attributes = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in NS.items())
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document {attributes} office:version="1.4" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
<office:meta><meta:generator xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">odsgen function samples</meta:generator></office:meta>
<office:automatic-styles>
<number:date-style style:name="iso-date"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>
<number:time-style style:name="clock"><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:time-style>
<style:style style:name="date" style:family="table-cell" style:data-style-name="iso-date"/>
<style:style style:name="time" style:family="table-cell" style:data-style-name="clock"/>
<style:style style:name="heading" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>
<style:style style:name="narrow" style:family="table-column"><style:table-column-properties style:column-width="1.6in"/></style:style>
<style:style style:name="wide" style:family="table-column"><style:table-column-properties style:column-width="1.3in"/></style:style>
<style:style style:name="result" style:family="table-column"><style:table-column-properties style:column-width="1.4in"/></style:style>
<style:style style:name="formula" style:family="table-column"><style:table-column-properties style:column-width="4.4in"/></style:style>
</office:automatic-styles>
<office:body><office:spreadsheet>
<table:calculation-settings table:case-sensitive="false" table:use-wildcards="true" table:use-regular-expressions="false"/>
{''.join(tables)}
{database_ranges(category)}
</office:spreadsheet></office:body></office:document>
"""


def recalculate(xml, name, work):
    """LibreOffice's own flat save of `xml`, every formula recalculated."""
    source = work / f"{name}.fods"
    source.write_text(xml, encoding="utf-8")
    out = work / "out"
    out.mkdir(exist_ok=True)
    subprocess.run(
        ["soffice", f"-env:UserInstallation=file://{work}/profile", "--headless",
         "--convert-to", "fods", "--outdir", str(out), str(source)],
        check=True, capture_output=True)
    return (out / f"{name}.fods").read_text(encoding="utf-8")


# --- reading LibreOffice's answer --------------------------------------------------------

def tag(prefix, local):
    return f"{{{NS[prefix]}}}{local}"


def cell_of(element):
    """One saved cell: its kind, its value as text, its formula and what it displays."""
    a = element.attrib
    formula = a.get(tag("table", "formula"))
    shown = "\n".join("".join(p.itertext()) for p in element.findall("text:p", NS))
    kind = a.get(tag("calcext", "value-type")) or a.get(tag("office", "value-type"))
    value = {
        "float": a.get(tag("office", "value")),
        "percentage": a.get(tag("office", "value")),
        "currency": a.get(tag("office", "value")),
        "date": a.get(tag("office", "date-value")),
        "time": a.get(tag("office", "time-value")),
        "boolean": a.get(tag("office", "boolean-value")),
        "string": a.get(tag("office", "string-value"), shown),
        "error": shown,
    }.get(kind)
    if kind is None and formula is None:
        return None
    cell = {"kind": kind, "value": value, "shown": shown}
    if formula is not None:
        cell["formula"] = formula.removeprefix("of:")
    return cell


def read_sheets(xml):
    """Every sheet LibreOffice saved: name → {"A1": cell}."""
    root = ET.fromstring(xml)
    sheets = {}
    for table in root.iter(tag("table", "table")):
        cells = {}
        row = 0
        for element in table.iter():
            if element.tag != tag("table", "table-row"):
                continue
            repeat = int(element.get(tag("table", "number-rows-repeated"), "1"))
            col = 0
            for child in element:
                if child.tag not in (tag("table", "table-cell"), tag("table", "covered-table-cell")):
                    continue
                span = int(child.get(tag("table", "number-columns-repeated"), "1"))
                cell = cell_of(child)
                if cell is not None:
                    for c in range(col, col + min(span, 64)):
                        cells[f"{column_letters(c)}{row + 1}"] = cell
                col += span
            row += min(repeat, 1) if repeat > 1000 else repeat
        sheets[table.get(tag("table", "name"))] = cells
    return sheets


# --- spellings ---------------------------------------------------------------------------

def split_strings(formula):
    """`formula` in pieces, each `(is_string, text)`, so nothing rewrites inside a literal."""
    pieces = re.split(r'("(?:[^"]|"")*")', formula)
    return [(i % 2 == 1, piece) for i, piece in enumerate(pieces)]


def displayed(formula):
    """An OpenFormula expression as LibreOffice's formula bar shows it: `[$Data.A1:.B2]` is
    `$Data.A1:B2`, `[.A1]` is `A1`."""
    def reference(match):
        parts = match.group(1).split(":")
        return ":".join(p[1:] if p.startswith(".") else p for p in parts)
    out = []
    for literal, piece in split_strings(formula):
        out.append(piece if literal else re.sub(r"\[([^\]]*)\]", reference, piece))
    return "".join(out)


EXCEL_RENAMED = {"COM.MICROSOFT.CEILING": "CEILING", "COM.MICROSOFT.FLOOR": "FLOOR"}
# OpenFormula's own CEILING and FLOOR take a Mode, and Excel's under the same name do not:
# a formula calling one has no Excel spelling at all.
ODF_ONLY = ("CEILING", "FLOOR")


def excel(formula):
    """The same formula in Excel's A1 syntax, or None where Excel cannot say it."""
    out = []
    for literal, piece in split_strings(formula):
        if literal:
            out.append(piece)
            continue
        if re.search(r"(?<![.\w])(%s)\(" % "|".join(ODF_ONLY), piece):
            return None
        for odf, xl in EXCEL_RENAMED.items():
            piece = piece.replace(odf + "(", xl + "(")

        def reference(match):
            parts = match.group(1).split(":")
            sheet = None
            spelled = []
            for p in parts:
                p = p.lstrip("$") if "." in p.lstrip(".") and not p.startswith(".") else p
                if p.startswith("."):
                    spelled.append(p[1:])
                else:
                    name, _, cell = p.rpartition(".")
                    sheet = sheet or name
                    spelled.append(cell)
            return (f"{sheet}!" if sheet else "") + ":".join(spelled)
        piece = re.sub(r"\[([^\]]*)\]", reference, piece)
        piece = piece.replace("TRUE()", "TRUE").replace("FALSE()", "FALSE").replace(";", ",")
        out.append(piece)
    return "".join(out)


# --- the run -------------------------------------------------------------------------------

def generate(category, work):
    first = read_sheets(recalculate(flat_document(category, [""] * len(category.samples)),
                                    category.slug, work))
    typed = ["=" + displayed(first["Samples"][f"C{row}"]["formula"].removeprefix("="))
             for row in range(2, len(category.samples) + 2)]
    saved = recalculate(flat_document(category, typed), category.slug, work)
    (OUT / f"{category.slug}.fods").write_text(saved, encoding="utf-8")
    sheets = read_sheets(saved)

    def source_formula(sheet, ref, saved_cell):
        # The formula as written here rather than as LibreOffice re-spelt it (`[$Data.A1]`):
        # what an importer is expected to produce.
        if sheet == "Samples":
            return category.samples[int(ref[1:]) - 2][1]
        row = int(re.sub(r"\D", "", ref)) - 1
        col = ord(ref[0]) - 65
        value = next(s for s in category.sheets if s.name == sheet).rows[row][col]
        return value.formula if isinstance(value, F) else saved_cell["formula"]

    described = []
    for sheet in [s.name for s in category.sheets] + ["Samples"]:
        cells = []
        for ref, cell in sheets[sheet].items():
            cell = dict(cell, ref=ref)
            if "formula" in cell:
                written = source_formula(sheet, ref, cell)
                cell["odf"] = written.removeprefix("=")
                xl = excel(written.removeprefix("="))
                cell["excel"] = xl
            cells.append(cell)
        spec = next((s for s in category.sheets if s.name == sheet), None)
        described.append({
            "name": sheet,
            "hidden": {str(k): v for k, v in (spec.hidden if spec else {}).items()},
            "filter": list(spec.filter) if spec and spec.filter else None,
            "cells": cells,
        })
    samples = []
    for row, (function, formula, what) in enumerate(category.samples, start=2):
        result = sheets["Samples"][f"C{row}"]
        samples.append({
            "function": function, "formula": formula, "typed": typed[row - 2],
            "excel": excel(formula.removeprefix("=")), "what": what,
            "kind": result["kind"], "value": result["value"], "shown": result["shown"],
        })
    return {
        "file": f"{category.slug}.fods", "slug": category.slug, "title": category.title,
        "sections": category.sections, "story": category.story,
        "sheets": described, "samples": samples,
    }


def readme(documents):
    lines = [
        "# Function samples",
        "",
        "Every function grind implements, called the way a real spreadsheet calls it. One document per",
        "category of ODF 1.4 Part 4 §6; each has realistic data and a `Samples` sheet whose rows are",
        "the function, the formula as typed, the formula itself and what it shows. Every result was",
        "computed by LibreOffice. Generated by `generate_function_samples.py` from",
        "`function_samples.py` — edit those, never these files. All data is fictional.",
        "",
    ]
    for doc in documents:
        lines += [f"## {doc['title']} — `{doc['file']}`", "", f"{doc['story']} ({doc['sections']})", "",
                  "| Function | Formula | Result | What it shows |", "|---|---|---|---|"]
        for s in doc["samples"]:
            typed = s["typed"].replace("|", "\\|")
            lines.append(f"| `{s['function']}` | `{typed}` | {s['shown']} | {s['what']} |")
        lines.append("")
    return "\n".join(lines)


def main():
    if shutil.which("soffice") is None:
        sys.exit("needs LibreOffice: soffice is not on PATH")
    OUT.mkdir(exist_ok=True)
    documents = []
    with tempfile.TemporaryDirectory() as tmp:
        for category in CATEGORIES:
            documents.append(generate(category, Path(tmp)))
            print(f"{category.slug}.fods: {len(category.samples)} samples")
    (OUT / "samples.json").write_text(
        json.dumps({"generator": "odsgen", "documents": documents},
                   indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    (OUT / "README.md").write_text(readme(documents), encoding="utf-8")


if __name__ == "__main__":
    main()
