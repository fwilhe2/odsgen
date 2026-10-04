#!/usr/bin/env python3
"""Generate fictional spreadsheet fixtures using LibreOffice's flat ODS support."""

from html import escape
from pathlib import Path
import subprocess
import tempfile


ROOT = Path(__file__).resolve().parent
FIXTURE_DIR = ROOT / "fixtures"
NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    "of": "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
}


def cell(value, style="", kind="string"):
    return {"value": value, "style": style, "kind": kind}


def formula(expression, result, style="currency"):
    return {"value": result, "style": style, "kind": "formula", "formula": expression}


def xml_cell(item):
    if item is None:
        return "<table:table-cell/>"
    attrs = []
    if item["style"]:
        attrs.append(f'table:style-name="{escape(item["style"], quote=True)}"')
    if item["kind"] == "formula":
        attrs.extend([
            f'table:formula="{escape(item["formula"], quote=True)}"',
            'office:value-type="float"',
            f'office:value="{item["value"]}"',
        ])
        display = str(item["value"])
    elif item["kind"] == "date":
        attrs.extend([
            'office:value-type="date"',
            f'office:date-value="{item["value"]}"',
        ])
        display = item["value"]
    elif item["kind"] == "float":
        attrs.extend([
            'office:value-type="float"',
            f'office:value="{item["value"]}"',
        ])
        display = str(item["value"])
    else:
        attrs.append('office:value-type="string"')
        display = str(item["value"])
    attr_text = " " + " ".join(attrs) if attrs else ""
    return f"<table:table-cell{attr_text}><text:p>{escape(display)}</text:p></table:table-cell>"


def flat_document(title, headers, rows, summary_rows=()):
    all_rows = [[cell(title, "title")], [cell("FICTIONAL TEST DATA | All names, organizations, and transactions are invented.", "note")],
                [cell(value, "header") for value in headers]]
    all_rows.extend(rows)
    if summary_rows:
        all_rows.append([])
        all_rows.extend(summary_rows)

    row_xml = []
    for row in all_rows:
        row_xml.append("<table:table-row>" + "".join(xml_cell(item) for item in row) + "</table:table-row>")
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="{NS['office']}" xmlns:style="{NS['style']}"
 xmlns:text="{NS['text']}" xmlns:table="{NS['table']}" xmlns:fo="{NS['fo']}"
 xmlns:number="{NS['number']}" xmlns:of="{NS['of']}" office:version="1.3"
 office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
 <office:automatic-styles>
    <number:number-style style:name="N2" number:language="en" number:country="US"><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:number-style>
    <number:currency-style style:name="Money" number:language="en" number:country="US"><number:currency-symbol>$</number:currency-symbol><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:currency-style>
    <number:percentage-style style:name="Pct" number:language="en" number:country="US"><number:number number:decimal-places="1"/></number:percentage-style>
    <number:date-style style:name="ISODate" number:language="en" number:country="US"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="2-digit"/><number:text>-</number:text><number:day number:style="2-digit"/></number:date-style>
  <style:style style:name="title" style:family="table-cell"><style:table-cell-properties fo:background-color="#17324D" fo:padding="0.12in"/><style:text-properties fo:font-size="16pt" fo:font-weight="bold" fo:color="#FFFFFF"/></style:style>
  <style:style style:name="note" style:family="table-cell"><style:table-cell-properties fo:background-color="#EAF1F5" fo:padding="0.06in"/><style:text-properties fo:font-style="italic" fo:color="#425466"/></style:style>
  <style:style style:name="header" style:family="table-cell"><style:table-cell-properties fo:background-color="#147D75" fo:padding="0.08in" fo:border-bottom="0.02in solid #0B514D"/><style:text-properties fo:font-weight="bold" fo:color="#FFFFFF"/></style:style>
  <style:style style:name="text" style:family="table-cell"><style:table-cell-properties fo:padding="0.05in" fo:border-bottom="0.01in solid #D8E1E8"/></style:style>
  <style:style style:name="date" style:family="table-cell" style:data-style-name="ISODate"><style:table-cell-properties fo:padding="0.05in" fo:border-bottom="0.01in solid #D8E1E8"/></style:style>
  <style:style style:name="number" style:family="table-cell" style:data-style-name="N2"><style:table-cell-properties fo:padding="0.05in" fo:border-bottom="0.01in solid #D8E1E8"/></style:style>
  <style:style style:name="currency" style:family="table-cell" style:data-style-name="Money"><style:table-cell-properties fo:padding="0.05in" fo:border-bottom="0.01in solid #D8E1E8"/></style:style>
  <style:style style:name="percent" style:family="table-cell" style:data-style-name="Pct"><style:table-cell-properties fo:padding="0.05in" fo:border-bottom="0.01in solid #D8E1E8"/></style:style>
  <style:style style:name="total" style:family="table-cell" style:data-style-name="Money"><style:table-cell-properties fo:background-color="#DDF2EC" fo:padding="0.07in" fo:border-top="0.02in solid #147D75"/><style:text-properties fo:font-weight="bold" fo:color="#17324D"/></style:style>
  <style:style style:name="label" style:family="table-cell"><style:table-cell-properties fo:background-color="#DDF2EC" fo:padding="0.07in" fo:border-top="0.02in solid #147D75"/><style:text-properties fo:font-weight="bold" fo:color="#17324D"/></style:style>
 </office:automatic-styles>
 <office:body><office:spreadsheet><table:table table:name="Data">
 {''.join(row_xml)}
 </table:table></office:spreadsheet></office:body>
</office:document>'''


def save_ods(filename, title, headers, rows, summary_rows=()):
    FIXTURE_DIR.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ods-fixture-") as temp_dir:
        source = Path(temp_dir) / f"{filename}.fods"
        source.write_text(flat_document(title, headers, rows, summary_rows), encoding="utf-8")
        subprocess.run([
            "libreoffice", "--headless", "--convert-to", "ods", "--outdir",
            str(FIXTURE_DIR), str(source),
        ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)


def budget_fixture():
    entries = [
        ("2026-09-01", "Income", "Example Works LLC", "September salary", 4200.00, "Income"),
        ("2026-09-03", "Housing", "Maple Key Rentals", "Monthly rent", 1450.00, "Expense"),
        ("2026-09-04", "Groceries", "North Market #018", "Weekly groceries", 86.42, "Expense"),
        ("2026-09-05", "Transit", "Metro Sample Pass", "Monthly transit pass", 72.00, "Expense"),
        ("2026-09-08", "Utilities", "Fictional Power Co.", "Electricity and water", 118.36, "Expense"),
        ("2026-09-10", "Dining", "Cafe Example", "Lunch with team", 24.50, "Expense"),
        ("2026-09-12", "Groceries", "North Market #018", "Pantry restock", 54.19, "Expense"),
        ("2026-09-14", "Health", "Sample Pharmacy", "Prescription refill", 18.75, "Expense"),
        ("2026-09-16", "Phone", "Orbit Mobile (Demo)", "Monthly plan", 45.00, "Expense"),
        ("2026-09-18", "Dining", "Noodle House 404", "Dinner", 38.20, "Expense"),
        ("2026-09-21", "Income", "Example Works LLC", "Expense reimbursement", 63.40, "Income"),
        ("2026-09-23", "Household", "Secondhand Shelf", "Desk lamp", 31.99, "Expense"),
    ]
    rows = [[cell(day, "date", "date"), cell(category, "text"), cell(merchant, "text"),
             cell(memo, "text"), cell(amount, "currency", "float"), cell(kind, "text")]
            for day, category, merchant, memo, amount, kind in entries]
    summary = [
        [cell("September 2026 summary", "label"), None, None, None, None, None],
        [cell("Income", "text"), formula('of:=SUMIF([.F4:.F15];"Income";[.E4:.E15])', 4263.40)],
        [cell("Expenses", "text"), formula('of:=SUMIF([.F4:.F15];"Expense";[.E4:.E15])', 1939.41)],
        [cell("Net", "label"), formula('of:=[.B18]-[.B19]', 2323.99, "total")],
    ]
    save_ods("fictional_monthly_budget", "Fictional Monthly Budget | September 2026",
             ["Date", "Category", "Merchant", "Description", "Amount", "Type"], rows, summary)


def orders_fixture():
    orders = [
        ("ORD-DEMO-1001", "2026-10-01", "Avery Example", "Desk organizer", 3, 18.50, 0.00, 0.0825),
        ("ORD-DEMO-1002", "2026-10-01", "Jordan Sample", "Reusable bottle", 2, 24.00, 0.10, 0.0825),
        ("ORD-DEMO-1003", "2026-10-02", "Morgan Fiction", "Notebook set", 5, 7.25, 0.05, 0.0000),
        ("ORD-DEMO-1004", "2026-10-03", "Casey Placeholder", "LED task light", 1, 42.95, 0.00, 0.0825),
        ("ORD-DEMO-1005", "2026-10-04", "Riley Example", "Cable clips", 4, 3.75, 0.00, 0.0825),
        ("ORD-DEMO-1006", "2026-10-05", "Quinn Sample", "Desk organizer", 2, 18.50, 0.15, 0.0825),
        ("ORD-DEMO-1007", "2026-10-05", "Taylor Fiction", "Notebook set", 1, 7.25, 0.00, 0.0000),
        ("ORD-DEMO-1008", "2026-10-06", "Drew Placeholder", "Reusable bottle", 3, 24.00, 0.05, 0.0825),
    ]
    rows = []
    grand_total = 0.0
    total_quantity = 0
    for index, (order_id, day, customer, product, quantity, price, discount, tax) in enumerate(orders, 4):
        subtotal = round(quantity * price * (1 - discount), 2)
        tax_amount = round(subtotal * tax, 2)
        grand_total += subtotal + tax_amount
        total_quantity += quantity
        rows.append([
            cell(order_id, "text"), cell(day, "date", "date"), cell(customer, "text"),
            cell(product, "text"), cell(quantity, "number", "float"),
            cell(price, "currency", "float"), cell(discount, "percent", "float"),
            formula(f"of:=[.E{index}]*[.F{index}]*(1-[.G{index}])", subtotal),
            cell(tax, "percent", "float"),
            formula(f"of:=[.H{index}]*[.I{index}]", tax_amount),
            formula(f"of:=[.H{index}]+[.J{index}]", round(subtotal + tax_amount, 2)),
        ])
    summary = [
        [cell("Order batch totals", "label")],
        [cell("Units", "text"), formula("of:=SUM([.E4:.E11])", total_quantity, "number")],
        [cell("Net sales", "text"), formula("of:=SUM([.H4:.H11])", 0, "total")],
        [cell("Tax collected", "text"), formula("of:=SUM([.J4:.J11])", 0, "total")],
        [cell("Order total", "label"), formula("of:=SUM([.K4:.K11])", round(grand_total, 2), "total")],
    ]
    net_sales = sum(round(qty * price * (1 - discount), 2) for _, _, _, _, qty, price, discount, _ in orders)
    taxes = round(grand_total - net_sales, 2)
    summary[2][1]["value"] = net_sales
    summary[3][1]["value"] = taxes
    save_ods("fictional_order_register", "Fictional Online Orders | October 2026",
             ["Order ID", "Date", "Customer", "Product", "Qty", "Unit price", "Discount", "Net", "Tax rate", "Tax", "Total"], rows, summary)


def timesheet_fixture():
    entries = [
        ("2026-09-01", "Alex Example", "Website refresh", "Wireframe review", 2.0, 88.00, "Yes"),
        ("2026-09-01", "Sam Sample", "Data cleanup", "Import validation", 3.5, 76.00, "Yes"),
        ("2026-09-02", "Alex Example", "Website refresh", "Product page layout", 4.0, 88.00, "Yes"),
        ("2026-09-02", "Jamie Fiction", "Internal", "Team planning", 1.0, 64.00, "No"),
        ("2026-09-03", "Sam Sample", "Data cleanup", "Duplicate review", 2.5, 76.00, "Yes"),
        ("2026-09-04", "Alex Example", "Website refresh", "Accessibility pass", 3.0, 88.00, "Yes"),
        ("2026-09-04", "Jamie Fiction", "Internal", "Documentation", 2.0, 64.00, "No"),
        ("2026-09-07", "Sam Sample", "Data cleanup", "Reconciliation", 4.0, 76.00, "Yes"),
        ("2026-09-08", "Alex Example", "Website refresh", "Client revisions", 2.5, 88.00, "Yes"),
        ("2026-09-09", "Jamie Fiction", "Internal", "Onboarding notes", 1.5, 64.00, "No"),
    ]
    rows = []
    total_hours = 0.0
    billable_hours = 0.0
    total_amount = 0.0
    for index, (day, person, project, task, hours, rate, billable) in enumerate(entries, 4):
        amount = round(hours * rate, 2)
        total_hours += hours
        total_amount += amount if billable == "Yes" else 0
        billable_hours += hours if billable == "Yes" else 0
        rows.append([
            cell(day, "date", "date"), cell(person, "text"), cell(project, "text"),
            cell(task, "text"), cell(hours, "number", "float"), cell(rate, "currency", "float"),
            formula(f"of:=[.E{index}]*[.F{index}]", amount), cell(billable, "text"),
        ])
    summary = [
        [cell("September 2026 totals", "label")],
        [cell("Hours", "text"), formula("of:=SUM([.E4:.E13])", total_hours, "number")],
        [cell("Billable hours", "text"), formula('of:=SUMIF([.H4:.H13];"Yes";[.E4:.E13])', billable_hours, "number")],
        [cell("Billable value", "text"), formula('of:=SUMIF([.H4:.H13];"Yes";[.G4:.G13])', total_amount, "total")],
    ]
    save_ods("fictional_project_timesheet", "Fictional Project Timesheet | September 2026",
             ["Date", "Consultant", "Project", "Work item", "Hours", "Rate", "Line value", "Billable"], rows, summary)


if __name__ == "__main__":
    budget_fixture()
    orders_fixture()
    timesheet_fixture()
    print(f"Generated fixtures in {FIXTURE_DIR}")