"""Generates realistic, fully fictional sample data for the Vendor Invoice Recon workbook.

Outputs (overwritten):
  Sample Inputs/PDF Converter.xlsx  -> table InvoiceDetails (vendor invoice lines)
  Sample Inputs/PBI Status.xlsx     -> sheet Data (open-items report, Power BI export layout)
Also prints the vendor-invoice header rows to paste into tblVendorInvoices.
Run:  python make_sample_data.py   (fixed seed = same data every run)
"""
import random
from datetime import date, timedelta
from pathlib import Path
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill

random.seed(20261124)
OUT = Path(__file__).resolve().parent.parent / "Sample Inputs"
FEE_RATE = 0.0125  # vendor charges 1.25 % of each customer invoice it handles

CUSTOMERS = [  # (account id, name, country, company code)
    ("C-100231", "Northwind Telecom S.A.", "ES", "ES10"),
    ("C-100457", "Contoso Mobile GmbH", "DE", "DE20"),
    ("C-100612", "Fabrikam Networks S.A. de C.V.", "MX", "MX30"),
    ("C-100788", "Adatum Wireless S.R.L.", "AR", "AR40"),
    ("C-100903", "Tailspin Broadband Ltd", "GB", "GB50"),
    ("C-101145", "Litware Communications S.p.A.", "IT", "IT60"),
    ("C-101322", "Proseware Fibra Ltda", "BR", "BR70"),
]

# vendor invoice nr, vendor invoice date, number of customer-invoice lines
VENDOR_INVOICES = [
    ("VND-2026-0912", date(2026, 9, 12), 150),
    ("VND-2026-0927", date(2026, 9, 27), 45),
    ("VND-2026-1008", date(2026, 10, 8), 120),
    ("VND-2026-1021", date(2026, 10, 21), 20),
    ("VND-2026-1103", date(2026, 11, 3), 300),
    ("VND-2026-1110", date(2026, 11, 10), 80),
    ("VND-2026-1117", date(2026, 11, 17), 12),
    ("VND-2026-1121", date(2026, 11, 21), 10),
]
CURRENT = {"VND-2026-1103", "VND-2026-1110", "VND-2026-1117", "VND-2026-1121"}
CLEAN = {"VND-2026-1117", "VND-2026-1121"}  # every line approvable -> APPROVE / CHECK AMOUNT

STATUS_OK = ["Ready to Pay (incl. SoR)"] * 9 + ["Paid Waiting Allocation or awaiting offset"]
STATUS_BLOCK = [
    "Customer Approval Required",
    "Customer Not Yet Contacted or Awaiting Response",
    "Invoice quality issue (IQI)",
    "Invoice/Credit Note not yet received by Customer",
]

_seq = iter(range(9100234000, 9100299999, 7))


def new_customer_invoice(vdate):
    cust = random.choice(CUSTOMERS)
    inv_date = vdate - timedelta(days=random.randint(35, 150))
    amount = round(random.choice([random.uniform(800, 9000)] * 3 + [random.uniform(9000, 85000)]), 2)
    return {"nr": str(next(_seq)), "cust": cust, "date": inv_date, "amount": amount}


# ---------- vendor invoice lines ----------
lines, billed = [], []  # billed = customer invoices already used in earlier vendor invoices
for vnr, vdate, n in VENDOR_INVOICES:
    picked, used = [], set()
    if billed:  # ~15 % of a vendor invoice re-bills customer invoices from earlier months
        for ci in random.sample(billed, min(len(billed), max(1, n * 15 // 100))):
            if ci["date"] < vdate - timedelta(days=30):
                picked.append(ci); used.add(ci["nr"])
    while len(picked) < n:
        picked.append(new_customer_invoice(vdate))
    for ci in picked[:n]:
        lines.append({"vnr": vnr, "vdate": vdate, "ci": ci,
                      "fee": round(ci["amount"] * FEE_RATE, 2)})
    billed.extend(ci for ci in picked[:n] if ci["nr"] not in used)

# ---------- PBI open items: status per customer invoice ----------
status_of = {}  # nr -> list of statuses (several lines possible); missing = already paid
for vnr, vdate, n in VENDOR_INVOICES:
    for ln in (l for l in lines if l["vnr"] == vnr):
        nr = ln["ci"]["nr"]
        if nr in status_of or nr in status_of.get("_paid", set()):
            continue
        r = random.random()
        if vnr in CLEAN:
            status_of[nr] = [random.choice(STATUS_OK)] if r < 0.85 else None
        elif vnr in CURRENT:
            if r < 0.10:   status_of[nr] = None                                   # paid already
            elif r < 0.80: status_of[nr] = [random.choice(STATUS_OK)]
            elif r < 0.90: status_of[nr] = [random.choice(STATUS_BLOCK)]
            else:          status_of[nr] = ["Ready to Pay (incl. SoR)", random.choice(STATUS_BLOCK)]  # split lines
        else:  # older vendor invoices: mostly paid by now
            status_of[nr] = None if r < 0.75 else [random.choice(STATUS_OK + STATUS_BLOCK)]
# a re-billed invoice inside a clean vendor invoice must still be approvable
for ln in lines:
    if ln["vnr"] in CLEAN and status_of.get(ln["ci"]["nr"]):
        if any(s in STATUS_BLOCK for s in status_of[ln["ci"]["nr"]]):
            status_of[ln["ci"]["nr"]] = ["Ready to Pay (incl. SoR)"]

hdr_fill = PatternFill("solid", fgColor="1F4E78")
hdr_font = Font(bold=True, color="FFFFFF")

# ---------- PDF Converter.xlsx ----------
wb = Workbook(); ws = wb.active; ws.title = "Details"
cols = ["Vendor Invoice Nr", "Provider invoice nr", "Vendor Invoice Date", "Customer Invoice Date",
        "Customer Name", "Country", "Customer Invoice Amount", "Currency", "Fee Amount"]
ws.append(cols)
for ln in lines:
    ci = ln["ci"]
    ws.append([ln["vnr"], ci["nr"], ln["vdate"], ci["date"], ci["cust"][1], ci["cust"][2],
               ci["amount"], "EUR", ln["fee"]])
for row in ws.iter_rows(min_row=2):
    row[2].number_format = row[3].number_format = "yyyy-mm-dd"
    row[6].number_format = row[8].number_format = "#,##0.00"
for c, w in zip("ABCDEFGHI", [17, 18, 18, 20, 30, 9, 22, 9, 12]):
    ws.column_dimensions[c].width = w
t = Table(displayName="InvoiceDetails", ref=f"A1:I{ws.max_row}")
t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(t); ws.freeze_panes = "A2"
wb.save(OUT / "PDF Converter.xlsx")

# ---------- PBI Status.xlsx (same 52 columns as the real export) ----------
PBI_COLS = ['Source System', 'Customer Co', 'SAP Co', 'Customer Name', 'CUSTOMER_ID', 'ACCOUNT_NAME', 'Market',
            'Market Unit', 'CBT', 'Customer Team', 'Country', 'Doc Number', 'Line', 'Fiscal Year', 'Ref Number',
            'Invoice Reference', 'Document Type Description', 'Document Type', 'Posting Key', 'Reporting Value',
            'Reporting Curr', 'Posting Date', 'Doc Date', 'Due Date', 'Days Overdue', 'Credit Control Lead',
            'Line Item Status Check', 'Line Item Status Check Comment history', 'Reason for late or non payment',
            'IQI Type', 'IQI Date', 'IQI Resolved date', 'GL Account', 'Special GL Type', 'Lead BG',
            'IQI Description', 'Document Invoice Amount', 'Document Currency', 'Local Invoice Amount',
            'Local Currency', 'Upcoming P2P Date', 'Invoice Status Customer', 'Manual Invoice Number',
            'Manual Invoice Value', 'PO Number', 'Assignment Id', 'Reference Key1', 'Reference Key2',
            'Reference Key3', 'Dispute open', 'Concat', 'Aging Bucket']
MARKET = {"ES": ("Europe", "South Europe"), "IT": ("Europe", "South Europe"), "DE": ("Europe", "Central Europe"),
          "GB": ("Europe", "North Europe"), "MX": ("Latin America", "Mexico & CA"),
          "AR": ("Latin America", "South Cone"), "BR": ("Latin America", "Brazil")}
LEADS = ["A. Moreno", "J. Fischer", "L. Ortega", "M. Rossi", "S. Patel"]
COMMENT = {"Customer Approval Required": "Sent to customer PM for acceptance",
           "Customer Not Yet Contacted or Awaiting Response": "Reminder sent, awaiting reply",
           "Invoice quality issue (IQI)": "PO line mismatch reported by customer",
           "Invoice/Credit Note not yet received by Customer": "Re-sent via customer portal"}
REPORT_DATE = date(2026, 11, 24)

wb = Workbook(); ws = wb.active; ws.title = "Data"
ws.append(PBI_COLS)
ci_by_nr = {l["ci"]["nr"]: l["ci"] for l in lines}
doc = 1800450000
for nr, sts in sorted(status_of.items()):
    if not sts:
        continue
    ci = ci_by_nr[nr]; acc, name, ctry, coco = ci["cust"]
    parts = [ci["amount"]] if len(sts) == 1 else [round(ci["amount"] * 0.6, 2), round(ci["amount"] - round(ci["amount"] * 0.6, 2), 2)]
    doc += 13
    for i, (st, amt) in enumerate(zip(sts, parts), start=1):
        due = ci["date"] + timedelta(days=60)
        overdue = (REPORT_DATE - due).days
        iqi = st.startswith("Invoice quality")
        row = {
            'Source System': "SAP P01", 'Customer Co': coco, 'SAP Co': coco, 'Customer Name': name,
            'CUSTOMER_ID': acc, 'ACCOUNT_NAME': name.upper(), 'Market': MARKET[ctry][0],
            'Market Unit': MARKET[ctry][1], 'CBT': f"CBT {name.split()[0]}", 'Customer Team': f"CT-{ctry}",
            'Country': ctry, 'Doc Number': str(doc), 'Line': i, 'Fiscal Year': ci["date"].year,
            'Ref Number': nr, 'Invoice Reference': nr, 'Document Type Description': "Customer Invoice",
            'Document Type': "RV", 'Posting Key': "01", 'Reporting Value': amt, 'Reporting Curr': "EUR",
            'Posting Date': ci["date"], 'Doc Date': ci["date"], 'Due Date': due, 'Days Overdue': overdue,
            'Credit Control Lead': random.choice(LEADS), 'Line Item Status Check': st,
            'Line Item Status Check Comment history': COMMENT.get(st, ""),
            'Reason for late or non payment': "" if st in STATUS_OK else st.split(" (")[0],
            'IQI Type': "PO mismatch" if iqi else None, 'IQI Date': ci["date"] + timedelta(days=20) if iqi else None,
            'IQI Resolved date': None, 'GL Account': "1400000", 'Special GL Type': None, 'Lead BG': "Network Services",
            'IQI Description': "Customer reports PO line mismatch" if iqi else None,
            'Document Invoice Amount': amt, 'Document Currency': "EUR", 'Local Invoice Amount': amt,
            'Local Currency': "EUR", 'Upcoming P2P Date': REPORT_DATE + timedelta(days=random.choice([3, 10, 17])),
            'Invoice Status Customer': "Open", 'Manual Invoice Number': None, 'Manual Invoice Value': None,
            'PO Number': f"45{random.randint(10000000, 99999999)}", 'Assignment Id': f"AS-{doc}-{i}",
            'Reference Key1': ctry, 'Reference Key2': acc, 'Reference Key3': None,
            'Dispute open': "Yes" if st in STATUS_BLOCK else "No", 'Concat': f"{coco}{doc}{i}",
            'Aging Bucket': "Not due" if overdue <= 0 else "1-30" if overdue <= 30 else "31-60" if overdue <= 60 else "61-90" if overdue <= 90 else "90+",
        }
        ws.append([row[c] for c in PBI_COLS])
for row in ws.iter_rows(min_row=2):
    for c in row:
        if isinstance(c.value, date):
            c.number_format = "yyyy-mm-dd"
for c in ws[1]:
    c.fill, c.font = hdr_fill, hdr_font
ws.freeze_panes = "A2"
wb.save(OUT / "PBI Status.xlsx")

# ---------- header rows for tblVendorInvoices (current month) ----------
print("VENDOR INVOICE HEADERS (current month):")
for vnr, vdate, n in VENDOR_INVOICES:
    if vnr in CURRENT:
        total = round(sum(l["fee"] for l in lines if l["vnr"] == vnr), 2)
        if vnr == "VND-2026-1121":
            total = round(total + 45.00, 2)  # vendor over-billed -> CHECK AMOUNT
        print(f"{vnr}\t{total}\t{vdate}\tEUR\tCollection fee {vdate:%b %Y}")
print("lines:", len(lines), "| PBI rows:", ws.max_row - 1)
