"""Every spreadsheet function grind implements, called the way a real spreadsheet calls it.

One `Category` per group of ODF 1.4 Part 4 §6, each with the data a person would actually have
— a household budget, an invoice, a class's marks, a price list, a staff roster, a mortgage —
and a list of samples: the function, the formula in ODF's own syntax, and one line saying what
it shows. `generate_function_samples.py` turns each category into a flat ODS document that
LibreOffice recalculates, so the value beside every formula is LibreOffice's answer.

All people, companies, addresses and numbers are fictional.

Formulas are written in OpenFormula's syntax (Part 4 §5): `[Sheet.A1:.B2]` for a range on
another sheet, `[.A1]` on the same one, `;` between arguments, `TRUE()` for a logical.
"""

from dataclasses import dataclass, field


@dataclass(frozen=True)
class D:
    """A date, ISO 8601."""

    iso: str


@dataclass(frozen=True)
class Tm:
    """A time of day, `HH:MM:SS`."""

    hms: str


@dataclass(frozen=True)
class F:
    """A formula in a data sheet, OpenFormula syntax with the leading `=`."""

    formula: str


@dataclass
class Sheet:
    name: str
    rows: list
    # Row number (1-based) → "filter" or "collapse".
    hidden: dict = field(default_factory=dict)
    # (last row, field number, values kept) — an autofilter over A1 to the last column.
    filter: tuple | None = None


@dataclass
class Category:
    slug: str
    title: str
    sections: str
    story: str
    sheets: list
    samples: list  # (function, formula, what it shows)


MATH = Category(
    slug="math",
    title="Mathematical",
    sections="Part 4 §6.16",
    story="A household's September budget, a garden plan and an order of supplies for it.",
    sheets=[
        Sheet("Ledger", [
            ["Date", "Category", "Payee", "Amount", "Account", "Cleared"],
            [D("2026-09-01"), "Income", "Northwind Payroll", 3250.00, "Checking", True],
            [D("2026-09-01"), "Housing", "Riverside Lettings", -1180.00, "Checking", True],
            [D("2026-09-03"), "Groceries", "Market Hall", -86.40, "Card", True],
            [D("2026-09-05"), "Transport", "City Transit", -49.00, "Card", True],
            [D("2026-09-08"), "Groceries", "Corner Bakery", -12.75, "Cash", True],
            [D("2026-09-10"), "Utilities", "Volt Energy", -94.30, "Checking", True],
            [D("2026-09-12"), "Dining", "Trattoria Sole", -58.20, "Card", True],
            [D("2026-09-15"), "Groceries", "Market Hall", -102.15, "Card", True],
            [D("2026-09-18"), "Income", "Invoice 2026-017", 640.00, "Checking", False],
            [D("2026-09-20"), "Transport", "Fuel Stop", -61.80, "Card", True],
            [D("2026-09-24"), "Utilities", "Aqua Water", -38.60, "Checking", False],
            [D("2026-09-27"), "Groceries", "Market Hall", -74.95, "Card", True],
            [D("2026-09-30"), "Savings", "Rainy-day fund", -400.00, "Transfer", True],
        ]),
        Sheet("Receipts", [
            ["Date", "Shop", "Amount"],
            [D("2026-09-03"), "Market Hall", 86.40],
            [D("2026-09-08"), "Corner Bakery", 12.75],
            [D("2026-09-15"), "Market Hall", 102.15],
            [D("2026-09-19"), "Corner Bakery", 9.30],
            [D("2026-09-22"), "Green Grocer", 23.60],
            [D("2026-09-27"), "Market Hall", 74.95],
            [D("2026-09-29"), "Green Grocer", 18.20],
            [D("2026-09-30"), "Corner Bakery", 6.10],
        ],
            # The filter keeps two shops; one Green Grocer receipt is also hidden by hand.
            hidden={3: "filter", 5: "filter", 8: "collapse", 9: "filter"},
            filter=(9, 1, ["Green Grocer", "Market Hall"])),
        Sheet("Garden", [
            ["Item", "Value", "Unit"],
            ["Flowerbed radius", 1.8, "m"],
            ["Fence run", 12.5, "m"],
            ["Fence height", 1.6, "m"],
            ["Ramp rise", 0.45, "m"],
            ["Ramp run", 5.4, "m"],
            ["Roof pitch", 35, "degrees"],
            ["Rafter length", 4.2, "m"],
            ["Plant varieties for the border", 6, "varieties"],
            ["Paving slabs needed", 47, "slabs"],
            ["Tiles per box", 12, "tiles"],
            ["Tiles needed", 130, "tiles"],
            ["Seedlings planted", 40, "plants"],
            ["Monthly growth", 0.12, "per month"],
            ["Soil H⁺ concentration", 0.00000032, "mol/L"],
            ["Shed, east of the house", 8, "m"],
            ["Shed, north of the house", 6, "m"],
        ]),
        Sheet("Order", [
            ["Item", "Quantity", "Unit price"],
            ["Topsoil, 40 L bag", 6, 7.49],
            ["Compost, 50 L bag", 4, 5.99],
            ["Gravel, 25 kg bag", 10, 3.25],
            ["Wildflower seed mix", 3, 4.50],
            ["Lawn edging, per metre", 12, 2.80],
        ]),
    ],
    samples=[
        ("SUM", "=SUM([Ledger.D2:.D14])", "What the month left over: income less spending"),
        ("SUM", "=SUM([Order.B2:.B4])", "Bags to carry home: topsoil, compost and gravel"),
        ("SUMIF", '=SUMIF([Ledger.B2:.B14];"Groceries";[Ledger.D2:.D14])', "Spent on groceries"),
        ("SUMIF", '=SUMIF([Ledger.D2:.D14];">0")', "Income only: every positive amount"),
        ("SUMIFS", '=SUMIFS([Ledger.D2:.D14];[Ledger.B2:.B14];"Groceries";[Ledger.E2:.E14];"Card")',
         "Groceries paid by card"),
        ("SUMIFS", '=SUMIFS([Ledger.D2:.D14];[Ledger.A2:.A14];">="&DATE(2026;9;15);[Ledger.D2:.D14];"<0")',
         "Spending in the second half of the month"),
        ("SUBTOTAL", "=SUBTOTAL(9;[Receipts.C2:.C9])",
         "Receipts the filter shows, added up — rows it hides are left out"),
        ("SUBTOTAL", "=SUBTOTAL(109;[Receipts.C2:.C9])",
         "The same, also leaving out the row hidden by hand"),
        ("SUBTOTAL", "=SUBTOTAL(2;[Receipts.C2:.C9])", "How many receipts are showing"),
        ("SUBTOTAL", "=SUBTOTAL(1;[Receipts.C2:.C9])", "Average of the receipts showing"),
        ("SUMPRODUCT", "=SUMPRODUCT([Order.B2:.B6];[Order.C2:.C6])",
         "The order's total: each quantity times its price"),
        ("SUMPRODUCT", '=SUMPRODUCT(([Ledger.B2:.B14]="Groceries")*([Ledger.E2:.E14]="Card");[Ledger.D2:.D14])',
         "Groceries by card, as it was written before SUMIFS"),
        ("PRODUCT", "=PRODUCT([Garden.B3];[Garden.B4])", "Area of the fence, to buy stain for"),
        ("ABS", '=ABS(SUMIF([Ledger.D2:.D14];"<0"))', "Total spending, as a positive figure"),
        ("SIGN", "=SIGN(SUM([Ledger.D2:.D14]))", "1 if the month ended in surplus, -1 in deficit"),
        ("QUOTIENT", "=QUOTIENT([Garden.B12];[Garden.B11])", "Full boxes of tiles"),
        ("MOD", "=MOD([Garden.B12];[Garden.B11])", "Tiles needed beyond the full boxes"),
        ("POWER", "=[Garden.B13]*POWER(1+[Garden.B14];6)", "Plants after six months of 12% monthly growth"),
        ("SQRT", "=SQRT([Garden.B5]^2+[Garden.B6]^2)", "Length of the ramp's surface"),
        ("EXP", "=[Garden.B13]*EXP([Garden.B14]*6)", "The same growth, compounded continuously"),
        ("LN", "=LN(2)/[Garden.B14]", "Months to double, growing continuously"),
        ("LOG", "=LOG(2;1+[Garden.B14])", "Months to double, growing once a month"),
        ("LOG10", "=-LOG10([Garden.B15])", "The soil's pH"),
        ("FACT", "=FACT([Garden.B9])", "Orders in which six varieties can be planted along the border"),
        ("EVEN", "=EVEN([Garden.B10])", "Paving slabs to order when they come in pairs"),
        ("ODD", "=ODD([Garden.B3]/2)", "Fence posts, two metres apart, an odd count so the gate is central"),
        ("PI", "=PI()*[Garden.B2]^2", "Area of the round flowerbed"),
        ("RADIANS", "=RADIANS([Garden.B7])", "The roof pitch in radians"),
        ("SIN", "=[Garden.B8]*SIN(RADIANS([Garden.B7]))", "Height a rafter rises"),
        ("COS", "=[Garden.B8]*COS(RADIANS([Garden.B7]))", "Width a rafter spans"),
        ("TAN", "=TAN(RADIANS([Garden.B7]))", "Rise of the roof per metre of run"),
        ("ATAN", "=ATAN([Garden.B5]/[Garden.B6])", "The ramp's slope, in radians"),
        ("DEGREES", "=DEGREES(ATAN([Garden.B5]/[Garden.B6]))", "The ramp's slope, in degrees"),
        ("ASIN", "=DEGREES(ASIN([Garden.B5]/SQRT([Garden.B5]^2+[Garden.B6]^2)))",
         "The same slope, from the rise and the surface length"),
        ("ACOS", "=DEGREES(ACOS([Garden.B6]/SQRT([Garden.B5]^2+[Garden.B6]^2)))",
         "The same slope, from the run and the surface length"),
        ("ATAN2", "=DEGREES(ATAN2([Garden.B16];[Garden.B17]))",
         "Bearing of the shed above due east — ATAN2 takes x first"),
    ],
)

ROUNDING = Category(
    slug="rounding",
    title="Rounding",
    sections="Part 4 §6.17",
    story="A consultant's invoice, a day's timesheet and a shipment.",
    sheets=[
        Sheet("Invoice", [
            ["Item", "Quantity", "Unit price", "Net"],
            ["Consulting hours", 7.25, 95.00, F("=[.B2]*[.C2]")],
            ["Travel, per km", 183, 0.42, F("=[.B3]*[.C3]")],
            ["Printed handouts", 35, 1.85, F("=[.B4]*[.C4]")],
            ["Room hire, half day", 1, 240.00, F("=[.B5]*[.C5]")],
            ["Catering, per person", 18, 12.40, F("=[.B6]*[.C6]")],
            [None],
            ["VAT rate", 0.19],
            ["Credit note", -12.34],
        ]),
        Sheet("Times", [
            ["Task", "Minutes"],
            ["Call with the client", 47],
            ["Drafting the report", 133],
            ["Review", 22],
            ["Site visit", 188],
            ["Boxes per pallet", 40],
            ["Boxes to ship", 257],
        ]),
    ],
    samples=[
        ("ROUND", "=ROUND(SUM([Invoice.D2:.D6])*[Invoice.B8];2)", "VAT on the invoice, to the cent"),
        ("ROUND", "=ROUND(SUM([Invoice.D2:.D6]);-2)", "The invoice to the nearest hundred, for a forecast"),
        ("INT", "=INT([Times.B3]/60)", "Whole hours in the drafting time"),
        ("INT", "=INT([Invoice.B9])", "INT steps down, so a credit of -12.34 becomes -13"),
        ("TRUNC", "=TRUNC([Invoice.B9])", "TRUNC cuts towards zero: -12"),
        ("TRUNC", "=TRUNC((DATE(2026;9;30)-DATE(2019;3;15))/365.25)", "Whole years of service"),
        ("ROUNDUP", "=ROUNDUP([Times.B2]/15;0)*15", "The call, billed in started quarter hours"),
        ("ROUNDDOWN", "=ROUNDDOWN(SUM([Invoice.D2:.D6]);0)", "Loyalty points: one per whole euro"),
        ("MROUND", "=MROUND([Times.B5];5)", "The site visit to the nearest five minutes"),
        ("MROUND", "=MROUND([Invoice.C6]*1.07;0.05)", "Catering with 7% added, to the nearest five cents"),
        ("CEILING", "=CEILING([Times.B7]/[Times.B6])", "Pallets needed for every box"),
        ("CEILING", "=CEILING([Invoice.C4]*1.1;0.25)", "Handout price plus 10%, up to the next quarter"),
        ("CEILING", "=CEILING([Invoice.B9];-0.5;1)", "With a Mode, the credit rounds away from zero: -12.5"),
        ("FLOOR", "=FLOOR([Times.B7]/[Times.B6])", "Pallets that go out full"),
        ("FLOOR", "=FLOOR([Invoice.C6]*0.9;0.1)", "Catering with a 10% discount, down to the ten cents"),
        ("COM.MICROSOFT.CEILING", "=COM.MICROSOFT.CEILING([Invoice.C2]*1.03;5)",
         "Excel's CEILING: next year's hourly rate, up to a step of five"),
        ("COM.MICROSOFT.CEILING", "=COM.MICROSOFT.CEILING([Invoice.B9];1)",
         "Excel's CEILING on a credit goes towards zero: -12"),
        ("COM.MICROSOFT.FLOOR", "=COM.MICROSOFT.FLOOR([Times.B5];30)", "The site visit in whole half hours"),
        ("COM.MICROSOFT.FLOOR", "=COM.MICROSOFT.FLOOR([Invoice.B9];1)",
         "Excel's FLOOR on a credit goes away from zero: -13"),
    ],
)

STATISTICAL = Category(
    slug="statistical",
    title="Statistical",
    sections="Part 4 §6.18",
    story="End-of-term marks for two classes. One English mark is missing and one was recorded as a note.",
    sheets=[
        Sheet("Marks", [
            ["Student", "Class", "Maths", "English", "Science"],
            ["Ann Becker", "7A", 78, 85, 81],
            ["Ben Okafor", "7A", 64, 72, 70],
            ["Chloe Martin", "7B", 91, 88, 94],
            ["Dev Patel", "7B", 55, "absent", 59],
            ["Ella Novak", "7A", 83, 79, 86],
            ["Finn Larsen", "7B", 70, None, 74],
            ["Grace Liu", "7A", 88, 92, 90],
            ["Hugo Silva", "7B", 47, 58, 52],
            ["Isla Moreau", "7A", 95, 90, 97],
            ["Jonah Weiss", "7B", 72, 68, 75],
        ]),
    ],
    samples=[
        ("AVERAGE", "=AVERAGE([Marks.C2:.C11])", "Mean maths mark"),
        ("AVERAGE", "=AVERAGE([Marks.D2:.D11])", "Mean English mark: the blank and the note are skipped, not zeros"),
        ("AVERAGEIF", '=AVERAGEIF([Marks.B2:.B11];"7A";[Marks.C2:.C11])', "Class 7A's mean maths mark"),
        ("AVERAGEIF", '=AVERAGEIF([Marks.E2:.E11];">=80")', "Mean of the science marks of 80 and above"),
        ("AVERAGEIFS", '=AVERAGEIFS([Marks.E2:.E11];[Marks.B2:.B11];"7B";[Marks.C2:.C11];">=60")',
         "Science in 7B, among those with at least 60 in maths"),
        ("MAX", "=MAX([Marks.C2:.E11])", "The best mark in any subject"),
        ("MIN", "=MIN([Marks.C2:.C11])", "The lowest maths mark"),
        ("LARGE", "=LARGE([Marks.C2:.C11];3)", "Third-best maths mark"),
        ("SMALL", "=SMALL([Marks.E2:.E11];2)", "Second-lowest science mark"),
        ("MEDIAN", "=MEDIAN([Marks.C2:.C11])", "The middle maths mark"),
        ("RANK", "=RANK([Marks.C4];[Marks.C2:.C11])", "Chloe's place in maths, best first"),
        ("RANK", "=RANK([Marks.C9];[Marks.C2:.C11];1)", "Hugo's place in maths, counted from the bottom"),
        ("STDEV", "=STDEV([Marks.C2:.C11])", "Spread of the maths marks, as a sample"),
        ("STDEVP", "=STDEVP([Marks.C2:.C11])", "The same, treating the classes as the whole population"),
        ("VAR", "=VAR([Marks.E2:.E11])", "Variance of the science marks, as a sample"),
        ("VARP", "=VARP([Marks.E2:.E11])", "The same, as a population"),
    ],
)

LOGICAL = Category(
    slug="logical",
    title="Logical",
    sections="Part 4 §6.15",
    story="A small shop's open orders, and the checks it runs on them.",
    sheets=[
        Sheet("Orders", [
            ["Order", "Customer", "Amount", "Paid", "Shipped", "Due", "Express", "Items"],
            ["SO-1041", "Harbor Café", 420.00, True, True, D("2026-10-02"), False, 12],
            ["SO-1042", "Lindqvist Design", 1280.50, False, False, D("2026-10-09"), True, 40],
            ["SO-1043", "Mason & Rowe", 96.80, True, False, D("2026-10-05"), False, 3],
            ["SO-1044", "Pine Street School", 2310.00, False, True, D("2026-09-28"), False, 75],
            ["SO-1045", "Quayside Gym", 0, True, True, D("2026-10-12"), True, 0],
        ]),
    ],
    samples=[
        ("IF", '=IF([Orders.D2];"Paid";"Open")', "SO-1041's payment status"),
        ("IF", "=IF([Orders.C3]>1000;[Orders.C3]*0.03;0)", "A 3% discount on orders over 1000"),
        ("AND", "=AND([Orders.D2];[Orders.E2])", "SO-1041 is paid and shipped: done"),
        ("AND", '=IF(AND([Orders.D4];[Orders.E4]);"Done";"Follow up")', "SO-1043 still needs attention"),
        ("OR", "=OR([Orders.G3];[Orders.C3]>1000)", "SO-1042 needs a signature: express, or over 1000"),
        ("NOT", "=NOT([Orders.E3])", "SO-1042 has still to ship"),
        ("TRUE", "=COUNTIF([Orders.D2:.D6];TRUE())", "Orders paid"),
        ("FALSE", "=COUNTIF([Orders.E2:.E6];FALSE())", "Orders not yet shipped"),
        ("IFERROR", '=IFERROR([Orders.C2]/[Orders.H2];"no items")', "Price per item on SO-1041"),
        ("IFERROR", '=IFERROR([Orders.C6]/[Orders.H6];"no items")', "SO-1045 has no items, so a note instead of #DIV/0!"),
        ("IFNA", '=IFNA(VLOOKUP("SO-1050";[Orders.A2:.C6];3;FALSE());"not on file")', "An order number that does not exist"),
        ("IFNA", '=IFNA(MATCH("Quayside Gym";[Orders.B2:.B6];0);0)', "Where Quayside Gym's order is in the list"),
    ],
)

INFORMATION = Category(
    slug="information",
    title="Information",
    sections="Part 4 §6.13",
    story="A stock-take in a hardware store: one count is missing, one was written down as a note, and the "
          "weeks-of-cover column has two errors in it.",
    sheets=[
        Sheet("Stock", [
            ["SKU", "Item", "On hand", "Reorder at", "Supplier", "Counted", "Discontinued", "Weekly use",
             "Weeks of cover"],
            ["A-100", "Hex bolts M8", 340, 200, "Boltworks", D("2026-09-28"), False, 60, F("=[.C2]/[.H2]")],
            ["A-101", "Hex nuts M8", 95, 200, "Boltworks", D("2026-09-28"), False, 80, F("=[.C3]/[.H3]")],
            ["A-102", "Washers M8", None, 500, "Boltworks", None, False, 120, F("=[.C4]/[.H4]")],
            ["B-200", "Wood screws 4×40", 1200, 400, "Timber & Co", D("2026-09-29"), False, 300, F("=[.C5]/[.H5]")],
            ["B-201", "Wall plugs 6 mm", "see bin 4", 300, "Timber & Co", D("2026-09-29"), False, 25,
             F("=[.C6]/[.H6]")],
            ["C-300", "Cable ties 200 mm", 0, 250, "Wirewell", D("2026-09-30"), True, 0, F("=[.C7]/[.H7]")],
            ["C-301", "Heat-shrink 6 mm", 48, 50, "Wirewell", D("2026-09-30"), False, 10, F("=[.C8]/[.H8]")],
        ]),
    ],
    samples=[
        ("COUNT", "=COUNT([Stock.C2:.C8])", "Items with a count: the note and the blank are not numbers"),
        ("COUNTA", "=COUNTA([Stock.C2:.C8])", "Items with anything written down"),
        ("COUNTBLANK", "=COUNTBLANK([Stock.C2:.C8])", "Items not counted yet"),
        ("COUNTIF", '=COUNTIF([Stock.E2:.E8];"Boltworks")', "Items bought from Boltworks"),
        ("COUNTIF", '=COUNTIF([Stock.A2:.A8];"A-*")', "SKUs in range A, by wildcard"),
        ("COUNTIFS", '=COUNTIFS([Stock.E2:.E8];"Wirewell";[Stock.G2:.G8];FALSE())', "Wirewell items still sold"),
        ("ISBLANK", "=ISBLANK([Stock.C4])", "The washers have not been counted"),
        ("ISNUMBER", "=ISNUMBER([Stock.C6])", "The wall plugs' count is a note, not a number"),
        ("ISTEXT", "=ISTEXT([Stock.C6])", "…which ISTEXT confirms"),
        ("ISNONTEXT", "=ISNONTEXT([Stock.C4])", "A blank is not text"),
        ("ISLOGICAL", "=ISLOGICAL([Stock.G2])", "The discontinued flag is a logical value"),
        ("ISERROR", "=ISERROR([Stock.I7])", "Cover for the cable ties cannot be worked out"),
        ("ISERR", "=ISERR([Stock.I6])", "Dividing a note is an error ISERR counts"),
        ("ISERR", '=ISERR(VLOOKUP("Z-999";[Stock.A2:.C8];3;FALSE()))', "A missing SKU is #N/A, which ISERR does not count"),
        ("ISNA", '=ISNA(MATCH("D-400";[Stock.A2:.A8];0))', "D-400 is not stocked"),
        ("NA", "=IF(ISBLANK([Stock.C4]);NA();[Stock.C4])", "#N/A for an uncounted item, so a chart leaves a gap"),
        ("N", "=N([Stock.C6])", "N turns the note into 0"),
        ("N", "=N([Stock.G7])", "…and TRUE into 1"),
        ("VALUE", "=VALUE(MID([Stock.A5];3;3))", "The number inside SKU B-200"),
        ("ROW", "=ROW([Stock.A5])", "The row SKU B-200 is in"),
        ("ROWS", "=ROWS([Stock.A2:.A8])", "SKUs on the list"),
        ("COLUMN", "=COLUMN([Stock.E1])", "Supplier is the fifth column"),
        ("COLUMNS", "=COLUMNS([Stock.A1:.I1])", "Fields kept per item"),
    ],
)

TEXT = Category(
    slug="text",
    title="Text",
    sections="Part 4 §6.20 and §6.11",
    story="A customer list as it arrives from a web form: stray spaces, mixed case, phone numbers in every format.",
    sheets=[
        Sheet("Customers", [
            ["ID", "Name", "Email", "Phone", "City", "Postcode", "Code", "Since"],
            ["C-0042", "  maria  ESPOSITO ", "maria.esposito@example.org", "+49 30 1234567", "berlin", "10115",
             "ESP-BER-42", 2019],
            ["C-0043", "Tom O'Neill", "tom.oneill@example.com", "+44 20 7946 0958", "London", "SW1A 1AA",
             "ONE-LON-43", 2021],
            ["C-0044", "Aiko Tanaka", "aiko@example.net", "+81 3 1234 5678", "Tokyo", "100-0001",
             "TAN-TOK-44", 2023],
            ["C-0045", "Lars Eriksen", "lars.eriksen@example.org", "+45 33 12 34 56", "Copenhagen", "1050",
             "ERI-COP-45", 2018],
        ]),
    ],
    samples=[
        ("TRIM", "=TRIM([Customers.B2])", "Maria's name without the stray spaces"),
        ("PROPER", "=PROPER(TRIM([Customers.B2]))", "…and in the case a letter is addressed in"),
        ("PROPER", "=PROPER([Customers.E2])", "A city typed in lower case"),
        ("UPPER", "=UPPER([Customers.E3])", "A city for a shipping label"),
        ("LOWER", "=LOWER([Customers.G2])", "A customer code for a URL"),
        ("LEN", "=LEN([Customers.D3])", "Characters in Tom's phone number"),
        ("LEFT", "=LEFT([Customers.G3];3)", "The surname part of a customer code"),
        ("RIGHT", "=RIGHT([Customers.A4];4)", "The number part of an ID"),
        ("MID", "=MID([Customers.G4];5;3)", "The city part of a customer code"),
        ("FIND", '=FIND("@";[Customers.C3])', "Where the domain of Tom's address begins"),
        ("FIND", '=LEFT([Customers.C3];FIND("@";[Customers.C3])-1)', "The user name in front of it"),
        ("SEARCH", '=SEARCH("example.org";[Customers.C2])', "SEARCH finds text too"),
        ("SEARCH", '=ISNUMBER(SEARCH("ORG";[Customers.C5]))', "…without minding case, which FIND would"),
        ("SUBSTITUTE", '=SUBSTITUTE([Customers.D3];" ";"")', "A phone number as one string of digits"),
        ("SUBSTITUTE", '=SUBSTITUTE([Customers.G2];"-";"/";2)', "Only the second dash replaced"),
        ("REPLACE", '=REPLACE([Customers.D2];1;3;"0")', "Maria's number in German national format"),
        ("REPT", '=REPT("*";LEN([Customers.D3])-4)&RIGHT([Customers.D3];4)', "A phone number masked for a receipt"),
        ("CONCATENATE", '=CONCATENATE(PROPER(TRIM([Customers.B2]));" <";[Customers.C2];">")',
         "An address line for an email"),
        ("EXACT", '=EXACT([Customers.E3];"london")', "EXACT minds case: London is not london"),
        ("T", "=T([Customers.F4])", "A postcode is text, so T keeps it"),
        ("T", "=T([Customers.H2])", "A year is a number, so T gives nothing"),
        ("HYPERLINK", '=HYPERLINK("mailto:"&[Customers.C4];[Customers.B4])', "A link that writes to Aiko"),
    ],
)

DATETIME = Category(
    slug="datetime",
    title="Date and time",
    sections="Part 4 §6.10",
    story="A website project's schedule, the public holidays it has to plan around, and one day's timesheet.",
    sheets=[
        Sheet("Schedule", [
            ["Task", "Start", "Working days", "End"],
            ["Kick-off", D("2026-09-01"), 1, F("=WORKDAY([.B2];[.C2]-1;[Holidays.A$2:.A$4])")],
            ["Requirements", D("2026-09-02"), 8, F("=WORKDAY([.B3];[.C3]-1;[Holidays.A$2:.A$4])")],
            ["Design", D("2026-09-14"), 10, F("=WORKDAY([.B4];[.C4]-1;[Holidays.A$2:.A$4])")],
            ["Build", D("2026-09-28"), 20, F("=WORKDAY([.B5];[.C5]-1;[Holidays.A$2:.A$4])")],
            ["Testing", D("2026-10-27"), 10, F("=WORKDAY([.B6];[.C6]-1;[Holidays.A$2:.A$4])")],
            ["Launch", D("2026-11-10"), 1, F("=WORKDAY([.B7];[.C7]-1;[Holidays.A$2:.A$4])")],
        ]),
        Sheet("Holidays", [
            ["Date", "Holiday"],
            [D("2026-10-03"), "Day of German Unity"],
            [D("2026-12-25"), "Christmas Day"],
            [D("2026-12-26"), "Boxing Day"],
        ]),
        Sheet("Timesheet", [
            ["Date", "In", "Out", "Break"],
            [D("2026-09-14"), Tm("08:47:30"), Tm("17:12:00"), Tm("00:45:00")],
        ]),
    ],
    samples=[
        ("DATE", "=DATE(2026;11;30)", "The hand-over date"),
        ("DATE", "=DATE(YEAR([Schedule.B2]);MONTH([Schedule.B2])+3;1)", "The first of the month a quarter after kick-off"),
        ("YEAR", "=YEAR([Schedule.B3])", "The year requirements start"),
        ("MONTH", "=MONTH([Schedule.B5])", "The month the build starts"),
        ("DAY", "=DAY([Schedule.B4])", "The day of the month design starts"),
        ("WEEKDAY", "=WEEKDAY([Schedule.B2];2)", "Kick-off's weekday, Monday being 1"),
        ("EDATE", "=EDATE([Schedule.B2];6)", "The six-month review"),
        ("EOMONTH", "=EOMONTH([Schedule.B2];0)", "Invoice date: the end of the kick-off month"),
        ("EOMONTH", "=EOMONTH([Schedule.B2];2)+1", "The first day of the third month"),
        ("NETWORKDAYS", "=NETWORKDAYS([Schedule.B2];[Schedule.D7];[Holidays.A2:.A4])",
         "Working days from kick-off to launch, holidays excluded"),
        ("WORKDAY", "=WORKDAY([Schedule.B6];10;[Holidays.A2:.A4])", "Ten working days after testing starts"),
        ("WEEKNUM", "=WEEKNUM([Schedule.B4];2)", "Design's week of the year, weeks starting on Monday"),
        ("ISOWEEKNUM", "=ISOWEEKNUM([Schedule.B4])", "The same week by ISO 8601"),
        ("TIME", "=TIME(8;30;0)", "When core hours begin"),
        ("TIME", "=[Timesheet.C2]-[Timesheet.B2]-[Timesheet.D2]>TIME(8;0;0)", "Worked more than eight hours?"),
        ("HOUR", "=HOUR([Timesheet.C2]-[Timesheet.B2]-[Timesheet.D2])", "Whole hours worked"),
        ("MINUTE", "=MINUTE([Timesheet.C2]-[Timesheet.B2]-[Timesheet.D2])", "…and the minutes beyond them"),
        ("SECOND", "=SECOND([Timesheet.B2])", "The seconds on the clock-in time"),
        ("HOUR", "=HOUR([Timesheet.B2])", "The hour of the clock-in time"),
        ("TODAY", "=YEAR(TODAY())>=2026", "TODAY is the date the sheet is recalculated"),
        ("NOW", "=INT(NOW())=TODAY()", "NOW is today plus the time of day"),
    ],
)

LOOKUP = Category(
    slug="lookup",
    title="Lookup",
    sections="Part 4 §6.14",
    story="A furniture shop's price list, sorted by SKU, and its shipping rates by weight band.",
    sheets=[
        Sheet("Prices", [
            ["SKU", "Product", "Unit", "Price", "Weight kg"],
            ["P-10", "Oak shelf 80 cm", "piece", 34.90, 4.2],
            ["P-11", "Oak shelf 120 cm", "piece", 49.90, 6.1],
            ["P-20", "Shelf bracket, black", "pair", 6.50, 0.4],
            ["P-21", "Shelf bracket, brass", "pair", 9.80, 0.5],
            ["P-30", "Wall anchors, 10", "pack", 3.20, 0.1],
        ]),
        Sheet("Shipping", [
            ["From kg", 0, 2, 5, 10, 20],
            ["Zone A", 4.90, 6.90, 9.90, 14.90, 24.90],
            ["Zone B", 7.90, 10.90, 15.90, 22.90, 34.90],
        ]),
    ],
    samples=[
        ("VLOOKUP", '=VLOOKUP("P-21";[Prices.A2:.D6];4;FALSE())', "Price of the brass brackets"),
        ("VLOOKUP", '=VLOOKUP("P-11";[Prices.A2:.E6];5;0)', "Weight of the long shelf"),
        ("HLOOKUP", "=HLOOKUP(6.1;[Shipping.B1:.F3];2;TRUE())", "Zone A rate for a 6.1 kg parcel"),
        ("HLOOKUP", "=HLOOKUP(0.5;[Shipping.B1:.F3];3)", "Zone B rate for a 0.5 kg parcel"),
        ("MATCH", '=MATCH("P-30";[Prices.A2:.A6];0)', "Where the wall anchors are on the list"),
        ("MATCH", "=MATCH(7;[Shipping.B1:.F1];1)", "Which weight band 7 kg falls in"),
        ("INDEX", "=INDEX([Prices.B2:.B6];MATCH(MAX([Prices.D2:.D6]);[Prices.D2:.D6];0))", "The most expensive product"),
        ("INDEX", "=INDEX([Shipping.B2:.F3];2;3)", "Zone B, third weight band"),
        ("CHOOSE", '=CHOOSE(WEEKDAY(DATE(2026;10;14);2);"Mon";"Tue";"Wed";"Thu";"Fri";"Sat";"Sun")',
         "The delivery day's name"),
        ("CHOOSE", '=CHOOSE(2;"Standard";"Express";"Pick-up")', "The second delivery option"),
        ("ADDRESS", '=ADDRESS(MATCH("P-20";[Prices.A2:.A6];0)+1;4)', "The cell holding the black brackets' price"),
        ("ADDRESS", "=ADDRESS(2;4;4)", "The same kind of address, relative"),
    ],
)

DATABASE = Category(
    slug="database",
    title="Database",
    sections="Part 4 §6.9",
    story="A staff roster, queried with criteria ranges on a sheet of their own.",
    sheets=[
        Sheet("Staff", [
            ["Name", "Department", "Role", "Salary", "Years", "Remote", "Retention"],
            ["Alice Brandt", "Engineering", "Developer", 68000, 4, True, 0.90],
            ["Bilal Haddad", "Engineering", "Lead", 82000, 9, False, 0.95],
            ["Carmen Ruiz", "Sales", "Account manager", 54000, 3, True, 0.80],
            ["Dmitri Volkov", "Sales", "Account manager", 57000, 6, False, 0.85],
            ["Emma Schulz", "Support", "Agent", 41000, 2, True, 0.70],
            ["Farah Khan", "Engineering", "Developer", 71000, 5, True, 0.90],
            ["Gustav Berg", "Support", "Lead", 49000, 8, False, 0.92],
            ["Hana Kim", "Sales", "Lead", 66000, 7, False, 0.88],
        ]),
        Sheet("Criteria", [
            ["Department", None, "Department", "Years", None, "Role", None, "Name", None, "Remote"],
            ["Engineering", None, "Sales", ">5", None, "Lead", None, "Emma Schulz", None, True],
        ]),
    ],
    samples=[
        ("DSUM", '=DSUM([Staff.A1:.G9];"Salary";[Criteria.A1:.A2])', "Engineering's payroll"),
        ("DAVERAGE", '=DAVERAGE([Staff.A1:.G9];"Salary";[Criteria.F1:.F2])', "A lead's average salary"),
        ("DCOUNT", '=DCOUNT([Staff.A1:.G9];"Years";[Criteria.C1:.D2])', "Sales staff with more than five years"),
        ("DCOUNTA", '=DCOUNTA([Staff.A1:.G9];"Name";[Criteria.J1:.J2])', "People working remotely"),
        ("DGET", '=DGET([Staff.A1:.G9];"Role";[Criteria.H1:.H2])', "Emma's role"),
        ("DGET", "=DGET([Staff.A1:.G9];4;[Criteria.H1:.H2])", "Her salary, by field number"),
        ("DMAX", '=DMAX([Staff.A1:.G9];"Salary";[Criteria.A1:.A2])', "The highest engineering salary"),
        ("DMIN", '=DMIN([Staff.A1:.G9];"Years";[Criteria.F1:.F2])', "The newest lead's years of service"),
        ("DPRODUCT", '=DPRODUCT([Staff.A1:.G9];"Retention";[Criteria.A1:.A2])',
         "Chance all of Engineering stays the year: the product of each retention estimate"),
        ("DSTDEV", '=DSTDEV([Staff.A1:.G9];"Salary";[Criteria.C1:.C2])', "Spread of sales salaries, as a sample"),
        ("DSTDEVP", '=DSTDEVP([Staff.A1:.G9];"Salary";[Criteria.C1:.C2])', "The same, as a population"),
        ("DVAR", '=DVAR([Staff.A1:.G9];"Salary";[Criteria.A1:.A2])', "Variance of engineering salaries, as a sample"),
        ("DVARP", '=DVARP([Staff.A1:.G9];"Salary";[Criteria.A1:.A2])', "The same, as a population"),
    ],
)

FINANCIAL = Category(
    slug="financial",
    title="Financial",
    sections="Part 4 §6.12",
    story="A mortgage, a savings plan, a delivery van being depreciated and an investment's cash flows.",
    sheets=[
        Sheet("Loan", [
            ["Item", "Value"],
            ["House price", 350000],
            ["Deposit", 70000],
            ["Annual rate", 0.0375],
            ["Years", 25],
            ["Monthly savings", 250],
            ["Savings rate", 0.03],
        ]),
        Sheet("Van", [
            ["Item", "Value"],
            ["Cost", 38500],
            ["Salvage value", 6000],
            ["Useful life, years", 6],
        ]),
        Sheet("Project", [
            ["Year", "Cash flow"],
            [0, -120000],
            [1, 28000],
            [2, 35000],
            [3, 41000],
            [4, 38000],
            [5, 30000],
            ["Discount rate", 0.08],
        ]),
    ],
    samples=[
        ("PMT", "=PMT([Loan.B4]/12;[Loan.B5]*12;-([Loan.B2]-[Loan.B3]))", "The monthly mortgage payment"),
        ("IPMT", "=IPMT([Loan.B4]/12;1;[Loan.B5]*12;-([Loan.B2]-[Loan.B3]))", "Interest in the first payment"),
        ("PPMT", "=PPMT([Loan.B4]/12;1;[Loan.B5]*12;-([Loan.B2]-[Loan.B3]))", "Principal in the first payment"),
        ("PPMT", "=PPMT([Loan.B4]/12;300;[Loan.B5]*12;-([Loan.B2]-[Loan.B3]))", "…and in the last one"),
        ("NPER", "=NPER([Loan.B4]/12;-2000;[Loan.B2]-[Loan.B3])", "Months to repay, paying 2000 a month"),
        ("FV", "=FV([Loan.B7]/12;10*12;-[Loan.B6])", "Savings after ten years of monthly deposits"),
        ("PV", "=PV(0.04/12;20*12;-1500)", "What 1500 a month for twenty years is worth today"),
        ("RATE", "=RATE(48;-450;18000)*12", "The annual rate of a car loan: 48 payments of 450 for 18000"),
        ("NPV", "=NPV([Project.B8];[Project.B3:.B7])+[Project.B2]", "The investment's net present value"),
        ("IRR", "=IRR([Project.B2:.B7])", "Its internal rate of return"),
        ("SLN", "=SLN([Van.B2];[Van.B3];[Van.B4])", "The van's yearly depreciation, straight-line"),
        ("SYD", "=SYD([Van.B2];[Van.B3];[Van.B4];1)", "First year, sum-of-years'-digits"),
        ("DDB", "=DDB([Van.B2];[Van.B3];[Van.B4];1)", "First year, double-declining balance"),
        ("DDB", "=DDB([Van.B2];[Van.B3];[Van.B4];2)", "Second year, double-declining balance"),
    ],
)

CATEGORIES = [MATH, ROUNDING, STATISTICAL, LOGICAL, INFORMATION, TEXT, DATETIME, LOOKUP, DATABASE, FINANCIAL]
