"""
gst_2b_reco.py
==============
Reconciles a purchase register with GSTR-2B and emails the vendors whose
invoices are missing in GSTR-2B.

What it does
------------
1. Reads the purchase register, the GSTR-2B Excel (B2B sheet) and the
   vendor master.
2. Cleans GSTINs, invoice numbers, dates and amounts.
3. Matches invoices in two passes (exact key, then tax + date tolerance).
4. Writes reconciliation_output.xlsx with 7 sheets.
5. Prepares ONE email per vendor for invoices not found in GSTR-2B.
   - DRY_RUN = True  -> nothing is sent; previews are saved as .html files
   - DRY_RUN = False -> emails are sent through Gmail SMTP
6. Writes email_log.xlsx and prints a short summary.

Run:  python gst_2b_reco.py
"""

import os
import re
import smtplib
import sys
import time
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent

# =============================================================================
# CONFIG - change anything here; the logic below does not need to be touched
# =============================================================================
CONFIG = {
    # ---- Input files --------------------------------------------------------
    "PURCHASE_REGISTER_FILE": BASE_DIR / "purchase_register.xlsx",
    "GSTR2B_FILE": BASE_DIR / "gstr2b.xlsx",
    "GSTR2B_SHEET": "B2B",
    "VENDOR_MASTER_FILE": BASE_DIR / "vendor_master.xlsx",

    # ---- Output files -------------------------------------------------------
    "OUTPUT_FILE": BASE_DIR / "reconciliation_output.xlsx",
    "EMAIL_LOG_FILE": BASE_DIR / "email_log.xlsx",
    "EMAIL_PREVIEW_FOLDER": BASE_DIR / "email_preview",

    # ---- Column names in the purchase register (exact headings) ------------
    "PURCHASE_COLUMNS": {
        "vendor": "Vendor Name",
        "gstin": "GSTIN",
        "inv_no": "Invoice No",
        "inv_date": "Invoice Date",
        "taxable": "Taxable Value",
        "igst": "IGST",
        "cgst": "CGST",
        "sgst": "SGST",
    },

    # ---- Column names in GSTR-2B --------------------------------------------
    # The portal file has a two-row header and the wording changes slightly
    # between versions, so each field lists keywords. A column is picked when
    # its heading CONTAINS any of the keywords (not case-sensitive).
    "GSTR2B_COLUMNS": {
        "gstin": ["gstin of supplier"],
        "vendor": ["trade/legal name", "trade name", "legal name"],
        "inv_no": ["invoice number", "invoice no"],
        "inv_date": ["invoice date"],
        "taxable": ["taxable value"],
        "igst": ["integrated tax", "igst"],
        "cgst": ["central tax", "cgst"],
        "sgst": ["state/ut tax", "state tax", "sgst"],
    },

    # ---- Column names in the vendor master (exact headings) ----------------
    "VENDOR_MASTER_COLUMNS": {
        "gstin": "GSTIN",
        "vendor": "Vendor Name",
        "email": "Email",
    },

    # ---- Matching rules -----------------------------------------------------
    "TAX_TOLERANCE": 10.0,       # Rs - allowed difference in total tax
    "DATE_TOLERANCE_DAYS": 5,    # days - used only in pass 2

    # ---- Firm details (used in the email) ----------------------------------
    "FIRM_NAME": "ABC & Co",
    "FIRM_CONTACT_PERSON": "Accounts Team",
    "FIRM_PHONE": "+91 90000 00000",
    "FIRM_EMAIL": "accounts@example.com",
    "TAX_PERIOD": "October 2026",
    "REPLY_WITHIN_DAYS": 7,

    # ---- Email settings -----------------------------------------------------
    "DRY_RUN": True,             # True = save previews only, send nothing
    "CC_EMAIL": "",              # optional, e.g. "partner@example.com"
    "SMTP_SERVER": "smtp.gmail.com",
    "SMTP_PORT": 587,
    "SECONDS_BETWEEN_EMAILS": 2,
    # Names of the environment variables holding the Gmail login.
    # The password itself is never written in this file.
    "ENV_SENDER_EMAIL": "GST_SENDER_EMAIL",
    "ENV_APP_PASSWORD": "GST_APP_PASSWORD",
}

AMOUNT_FIELDS = ["taxable", "igst", "cgst", "sgst"]
INDIAN_NUMBER_FORMAT = r"[>=10000000]##\,##\,##\,##0.00;[>=100000]##\,##\,##0.00;##,##0.00"


# =============================================================================
# Small helpers
# =============================================================================
def stop(message):
    """Print a clear error and end the script."""
    print(f"\nERROR: {message}\n")
    sys.exit(1)


def clean_text(value):
    """Turn any cell value into clean text (118.0 -> '118', NaN -> '')."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def clean_gstin(value):
    return clean_text(value).upper().replace(" ", "")


def invoice_key(value):
    """
    Normalised invoice number used only for matching.
    'inv/0045' -> 'INV45',  'GST-00123/25-26' -> 'GST1232526'
    """
    text = clean_text(value).upper()
    text = re.sub(r"(?<![0-9])0+(?=[0-9])", "", text)   # leading zeros of each number
    return re.sub(r"[^A-Z0-9]", "", text)               # spaces, / - and other symbols


def parse_date(value):
    """Parse a date as day-first (dd-mm-yyyy). Real Excel dates are kept as is."""
    if value is None or (isinstance(value, float) and pd.isna(value)) or value is pd.NaT:
        return pd.NaT
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.Timestamp(value).normalize()
    return pd.to_datetime(str(value).strip(), dayfirst=True, errors="coerce")


def to_amount(series):
    """Convert to numbers; blanks and text become 0. Commas are removed."""
    cleaned = series.astype(str).str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def inr(amount):
    """Format a number in the Indian style: 1234567.5 -> '12,34,567.50'."""
    amount = round(float(amount), 2)
    sign = "-" if amount < 0 else ""
    whole, decimals = f"{abs(amount):.2f}".split(".")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        head = re.sub(r"(?<=\d)(?=(\d\d)+$)", ",", head)
        whole = f"{head},{tail}"
    return f"{sign}{whole}.{decimals}"


def read_excel_or_stop(path, label, **kwargs):
    path = Path(path)
    if not path.exists():
        stop(f"{label} not found: {path}\nCheck the file name and path in CONFIG.")
    try:
        return pd.read_excel(path, **kwargs)
    except ValueError as err:
        stop(f"Could not read {label} ({path.name}): {err}")
    except Exception as err:  # corrupt file, file open in Excel, etc.
        stop(f"Could not open {label} ({path.name}): {err}")


def finish_cleaning(df):
    """Common cleaning for purchase register and GSTR-2B."""
    df = df.copy()
    df["vendor"] = df["vendor"].map(clean_text)
    df["gstin"] = df["gstin"].map(clean_gstin)
    df["inv_no"] = df["inv_no"].map(clean_text)
    df["inv_date"] = pd.to_datetime(df["inv_date"].map(parse_date))
    for field in AMOUNT_FIELDS:
        df[field] = to_amount(df[field])
    df["total_tax"] = (df["igst"] + df["cgst"] + df["sgst"]).round(2)
    df["key"] = df["inv_no"].map(invoice_key)
    # drop fully blank lines (no GSTIN and no invoice number)
    df = df[(df["gstin"] != "") | (df["inv_no"] != "")]
    return df.reset_index(drop=True)


# =============================================================================
# Reading the three input files
# =============================================================================
def load_purchase_register():
    columns = CONFIG["PURCHASE_COLUMNS"]
    raw = read_excel_or_stop(CONFIG["PURCHASE_REGISTER_FILE"], "Purchase register")
    raw.columns = [str(c).strip() for c in raw.columns]
    missing = [name for name in columns.values() if name not in raw.columns]
    if missing:
        stop(
            "Purchase register is missing column(s): " + ", ".join(missing)
            + "\nColumns found: " + ", ".join(raw.columns)
            + "\nFix the file or update PURCHASE_COLUMNS in CONFIG."
        )
    df = raw[list(columns.values())].copy()
    df.columns = list(columns.keys())
    return finish_cleaning(df)


def load_gstr2b():
    """
    Reads the B2B sheet of the portal file. The heading row is not on row 1
    and usually spans two rows, so the header is detected automatically.
    """
    path = Path(CONFIG["GSTR2B_FILE"])
    if not path.exists():
        stop(f"GSTR-2B file not found: {path}\nCheck the file name and path in CONFIG.")
    try:
        sheets = pd.ExcelFile(path).sheet_names
    except Exception as err:
        stop(f"Could not open GSTR-2B file ({path.name}): {err}")
    sheet = CONFIG["GSTR2B_SHEET"]
    if sheet not in sheets:
        stop(f"Sheet '{sheet}' not found in {path.name}. Sheets available: {', '.join(sheets)}")

    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    keywords = CONFIG["GSTR2B_COLUMNS"]

    # 1. Find the row that contains the supplier GSTIN heading
    header_row = None
    for i in range(min(len(raw), 30)):
        cells = [clean_text(v).lower() for v in raw.iloc[i]]
        if any(any(k in cell for k in keywords["gstin"]) for cell in cells):
            header_row = i
            break
    if header_row is None:
        stop(f"Could not find the header row in sheet '{sheet}' of {path.name} "
             f"(looked for '{keywords['gstin'][0]}' in the first 30 rows).")

    top = [clean_text(v) for v in raw.iloc[header_row]]
    gstin_col = next(i for i, cell in enumerate(top)
                     if any(k in cell.lower() for k in keywords["gstin"]))

    # 2. If the next row has no GSTIN under the GSTIN heading, it is the
    #    second header row (sub-headings such as 'Invoice number').
    headers, data_start = top, header_row + 1
    if data_start < len(raw) and clean_text(raw.iloc[data_start, gstin_col]) == "":
        sub = [clean_text(v) for v in raw.iloc[data_start]]
        headers = [s if s else t for s, t in zip(sub, top)]
        data_start += 1

    # 3. Map headings to our field names using the keywords
    lower = [h.lower() for h in headers]
    picked, missing = {}, []
    for field, words in keywords.items():
        position = next((i for i, h in enumerate(lower) if any(w in h for w in words)), None)
        if position is None:
            missing.append(f"{field} (keywords: {', '.join(words)})")
        else:
            picked[field] = position
    if missing:
        stop("GSTR-2B is missing column(s): " + "; ".join(missing)
             + "\nHeadings found: " + ", ".join(h for h in headers if h)
             + "\nUpdate GSTR2B_COLUMNS in CONFIG.")

    data = raw.iloc[data_start:]
    df = pd.DataFrame({field: data.iloc[:, pos].values for field, pos in picked.items()})
    df = finish_cleaning(df)
    # keep only real invoice lines (a valid GSTIN is 15 characters)
    return df[df["gstin"].str.len() == 15].reset_index(drop=True)


def load_vendor_master():
    columns = CONFIG["VENDOR_MASTER_COLUMNS"]
    raw = read_excel_or_stop(CONFIG["VENDOR_MASTER_FILE"], "Vendor master")
    raw.columns = [str(c).strip() for c in raw.columns]
    missing = [name for name in columns.values() if name not in raw.columns]
    if missing:
        stop("Vendor master is missing column(s): " + ", ".join(missing)
             + "\nColumns found: " + ", ".join(raw.columns))
    df = raw[list(columns.values())].copy()
    df.columns = list(columns.keys())
    df["gstin"] = df["gstin"].map(clean_gstin)
    df["email"] = df["email"].map(clean_text)
    return dict(zip(df["gstin"], df["email"]))   # GSTIN -> email


# =============================================================================
# Matching
# =============================================================================
def reconcile(books, gstr2b):
    """
    Adds two columns to the purchase register:
      status   - Matched / Value mismatch / Probable match / Not in 2B
      b2b_row  - row number of the matched GSTR-2B invoice (-1 if none)
    A GSTR-2B invoice is never used twice.
    """
    tolerance = CONFIG["TAX_TOLERANCE"]
    max_days = CONFIG["DATE_TOLERANCE_DAYS"]
    books = books.copy()
    books["status"] = "Not in 2B"
    books["b2b_row"] = -1
    used = set()

    # Index GSTR-2B by (GSTIN, normalised invoice number)
    by_key = {}
    for j, (gstin, key) in enumerate(zip(gstr2b["gstin"], gstr2b["key"])):
        by_key.setdefault((gstin, key), []).append(j)

    # ---- Pass 1: GSTIN + normalised invoice number --------------------------
    for i in books.index:
        if books.at[i, "key"] == "":
            continue
        options = [j for j in by_key.get((books.at[i, "gstin"], books.at[i, "key"]), [])
                   if j not in used]
        if not options:
            continue
        tax = books.at[i, "total_tax"]
        j = min(options, key=lambda r: abs(gstr2b.at[r, "total_tax"] - tax))
        used.add(j)
        books.at[i, "b2b_row"] = j
        difference = abs(gstr2b.at[j, "total_tax"] - tax)
        books.at[i, "status"] = "Matched" if difference <= tolerance else "Value mismatch"

    # ---- Pass 2: GSTIN + total tax within tolerance + date within N days ----
    for i in books.index[books["status"] == "Not in 2B"]:
        tax, date = books.at[i, "total_tax"], books.at[i, "inv_date"]
        if pd.isna(date):
            continue
        best, best_score = None, None
        for j in gstr2b.index[gstr2b["gstin"] == books.at[i, "gstin"]]:
            if j in used or pd.isna(gstr2b.at[j, "inv_date"]):
                continue
            tax_gap = abs(gstr2b.at[j, "total_tax"] - tax)
            day_gap = abs((gstr2b.at[j, "inv_date"] - date).days)
            if tax_gap <= tolerance and day_gap <= max_days:
                score = (tax_gap, day_gap)
                if best is None or score < best_score:
                    best, best_score = j, score
        if best is not None:
            used.add(best)
            books.at[i, "b2b_row"] = best
            books.at[i, "status"] = "Probable match"

    not_in_books = gstr2b.loc[[j for j in gstr2b.index if j not in used]]
    return books, not_in_books


# =============================================================================
# Building the output sheets
# =============================================================================
def invoice_sheet(df):
    """Standard invoice listing used for 'Not in 2B' and 'Not in books'."""
    return pd.DataFrame({
        "Vendor Name": df["vendor"].values,
        "GSTIN": df["gstin"].values,
        "Invoice No": df["inv_no"].values,
        "Invoice Date": df["inv_date"].values,
        "Taxable Value": df["taxable"].values,
        "IGST": df["igst"].values,
        "CGST": df["cgst"].values,
        "SGST": df["sgst"].values,
        "Total Tax": df["total_tax"].values,
    })


def comparison_sheet(books_rows, gstr2b, remark=None):
    """Books and GSTR-2B side by side for matched / probable / mismatch rows."""
    two_b = gstr2b.loc[books_rows["b2b_row"].values]
    out = pd.DataFrame({
        "Vendor Name": books_rows["vendor"].values,
        "GSTIN": books_rows["gstin"].values,
        "Invoice No (Books)": books_rows["inv_no"].values,
        "Invoice No (2B)": two_b["inv_no"].values,
        "Invoice Date (Books)": books_rows["inv_date"].values,
        "Invoice Date (2B)": two_b["inv_date"].values,
        "Taxable Value (Books)": books_rows["taxable"].values,
        "Taxable Value (2B)": two_b["taxable"].values,
        "Total Tax (Books)": books_rows["total_tax"].values,
        "Total Tax (2B)": two_b["total_tax"].values,
    })
    out["Tax Difference (Books - 2B)"] = (out["Total Tax (Books)"] - out["Total Tax (2B)"]).round(2)
    if remark:
        out["Remark"] = remark
    return out


def build_sheets(books, gstr2b, not_in_books):
    matched = books[books["status"] == "Matched"]
    mismatch = books[books["status"] == "Value mismatch"]
    probable = books[books["status"] == "Probable match"]
    missing = books[books["status"] == "Not in 2B"]

    not_in_2b = invoice_sheet(missing)

    vendor_wise = (
        not_in_2b.groupby(["Vendor Name", "GSTIN"], as_index=False)
        .agg(**{"No. of Invoices": ("Invoice No", "count"),
                "Taxable Value": ("Taxable Value", "sum"),
                "ITC at Risk": ("Total Tax", "sum")})
        .sort_values("ITC at Risk", ascending=False)
        .reset_index(drop=True)
    ) if len(not_in_2b) else pd.DataFrame(
        columns=["Vendor Name", "GSTIN", "No. of Invoices", "Taxable Value", "ITC at Risk"])

    summary = pd.DataFrame([
        ["Matched", len(matched), matched["total_tax"].sum(), "Claim ITC"],
        ["Value mismatch", len(mismatch), mismatch["total_tax"].sum(),
         "Check invoice; ask vendor for correction or credit/debit note"],
        ["Probable match - check invoice number", len(probable), probable["total_tax"].sum(),
         "Confirm the invoice number, then treat as matched"],
        ["Not in 2B", len(missing), missing["total_tax"].sum(),
         "ITC at risk - follow up with vendor"],
        ["Not in books (in 2B only)", len(not_in_books), not_in_books["total_tax"].sum(),
         "Check for missed entry or invoice not belonging to us"],
    ], columns=["Category", "No. of Invoices", "Total Tax (ITC)", "Suggested Action"])

    return {
        "Summary": summary,
        "Not in 2B": not_in_2b,
        "Vendor-wise ITC at risk": vendor_wise,
        "Not in books": invoice_sheet(not_in_books),
        "Probable matches": comparison_sheet(probable, gstr2b,
                                             "Probable match - check invoice number"),
        "Value mismatches": comparison_sheet(mismatch, gstr2b),
        "Matched": comparison_sheet(matched, gstr2b),
    }


def write_workbook(path, sheets):
    """Write sheets to Excel with bold headers, auto-fit widths, Indian format."""
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)
                ws = writer.sheets[name]
                for cell in ws[1]:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="1F4E78")
                    cell.alignment = Alignment(horizontal="center", vertical="center",
                                               wrap_text=True)
                ws.freeze_panes = "A2"
                for col_no, column in enumerate(df.columns, start=1):
                    width = len(str(column))
                    for row in range(2, ws.max_row + 1):
                        cell = ws.cell(row=row, column=col_no)
                        if isinstance(cell.value, datetime):
                            cell.number_format = "DD-MM-YYYY"
                            shown = "00-00-0000"
                        elif isinstance(cell.value, float):
                            cell.number_format = INDIAN_NUMBER_FORMAT
                            shown = inr(cell.value)
                        else:
                            shown = "" if cell.value is None else str(cell.value)
                        width = max(width, len(shown))
                    ws.column_dimensions[get_column_letter(col_no)].width = min(width + 3, 60)
    except PermissionError:
        stop(f"Cannot write {Path(path).name}. Close the file in Excel and run again.")


# =============================================================================
# Emails
# =============================================================================
def email_subject():
    return f"Invoices not reflected in GSTR-2B - {CONFIG['FIRM_NAME']} - {CONFIG['TAX_PERIOD']}"


def email_html(vendor, gstin, invoices):
    rows = "".join(
        "<tr>"
        f"<td style='padding:6px 10px;border:1px solid #999'>{r['Invoice No']}</td>"
        f"<td style='padding:6px 10px;border:1px solid #999'>"
        f"{'' if pd.isna(r['Invoice Date']) else pd.Timestamp(r['Invoice Date']).strftime('%d-%m-%Y')}</td>"
        f"<td style='padding:6px 10px;border:1px solid #999;text-align:right'>{inr(r['Taxable Value'])}</td>"
        f"<td style='padding:6px 10px;border:1px solid #999;text-align:right'>{inr(r['Total Tax'])}</td>"
        "</tr>"
        for _, r in invoices.iterrows()
    )
    head = "padding:6px 10px;border:1px solid #999;background:#1F4E78;color:#fff"
    c = CONFIG
    return f"""<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#222">
<p>Dear Sir/Madam,</p>
<p>Greetings from {c['FIRM_NAME']}.</p>
<p>While reconciling our purchase records with GSTR-2B for <b>{c['TAX_PERIOD']}</b>,
we noticed that the following invoice(s) issued by <b>{vendor}</b> (GSTIN: {gstin})
do not appear in our GSTR-2B for the period:</p>
<table style="border-collapse:collapse">
<tr><th style="{head}">Invoice No</th><th style="{head}">Invoice Date</th>
<th style="{head}">Taxable Value (Rs)</th><th style="{head}">Total Tax (Rs)</th></tr>
{rows}
<tr><td colspan="2" style="padding:6px 10px;border:1px solid #999"><b>Total</b></td>
<td style="padding:6px 10px;border:1px solid #999;text-align:right"><b>{inr(invoices['Taxable Value'].sum())}</b></td>
<td style="padding:6px 10px;border:1px solid #999;text-align:right"><b>{inr(invoices['Total Tax'].sum())}</b></td></tr>
</table>
<p>We request you to report these invoices in your GSTR-1 / IFF at the earliest so that
the input tax credit is available to us. If they have already been reported, kindly
confirm the date of filing and the return period.</p>
<p>We would appreciate your reply within <b>{c['REPLY_WITHIN_DAYS']} days</b>.
Please ignore this email if the invoices have since been reported.</p>
<p>Thank you for your cooperation.</p>
<p>Regards,<br>{c['FIRM_CONTACT_PERSON']}<br><b>{c['FIRM_NAME']}</b><br>
Phone: {c['FIRM_PHONE']}<br>Email: {c['FIRM_EMAIL']}</p>
</body></html>"""


def email_plain_text(vendor, gstin, invoices):
    lines = [
        "Dear Sir/Madam,", "",
        f"While reconciling our purchase records with GSTR-2B for {CONFIG['TAX_PERIOD']}, "
        f"we noticed that the following invoice(s) issued by {vendor} (GSTIN: {gstin}) "
        "do not appear in our GSTR-2B for the period:", "",
    ]
    for _, r in invoices.iterrows():
        date = "" if pd.isna(r["Invoice Date"]) else pd.Timestamp(r["Invoice Date"]).strftime("%d-%m-%Y")
        lines.append(f"  {r['Invoice No']} | {date} | Taxable Rs {inr(r['Taxable Value'])} "
                     f"| Tax Rs {inr(r['Total Tax'])}")
    lines += [
        "",
        "We request you to report these invoices in your GSTR-1 / IFF at the earliest, "
        "or confirm the date of filing if already reported.",
        f"We would appreciate your reply within {CONFIG['REPLY_WITHIN_DAYS']} days.", "",
        "Regards,", CONFIG["FIRM_CONTACT_PERSON"], CONFIG["FIRM_NAME"],
        f"Phone: {CONFIG['FIRM_PHONE']}", f"Email: {CONFIG['FIRM_EMAIL']}",
    ]
    return "\n".join(lines)


def connect_smtp():
    """Log in to Gmail. Returns (connection, sender, error_message)."""
    sender = os.environ.get(CONFIG["ENV_SENDER_EMAIL"], "").strip()
    password = os.environ.get(CONFIG["ENV_APP_PASSWORD"], "").replace(" ", "")
    if not sender or not password:
        return None, sender, (f"Environment variables {CONFIG['ENV_SENDER_EMAIL']} and "
                              f"{CONFIG['ENV_APP_PASSWORD']} are not set")
    try:
        server = smtplib.SMTP(CONFIG["SMTP_SERVER"], CONFIG["SMTP_PORT"], timeout=30)
        server.starttls()
        server.login(sender, password)
        return server, sender, None
    except Exception as err:
        return None, sender, f"Gmail login failed: {err}"


def send_vendor_emails(not_in_2b, vendor_emails):
    """One email per vendor. Returns the email log as a DataFrame."""
    log = []
    if not_in_2b.empty:
        return pd.DataFrame(columns=["Vendor", "GSTIN", "Email", "No. of Invoices",
                                     "Total Tax", "Status", "Timestamp"])

    dry_run = CONFIG["DRY_RUN"]
    preview_folder = Path(CONFIG["EMAIL_PREVIEW_FOLDER"])
    server, sender, login_error = (None, "", None)
    if dry_run:
        preview_folder.mkdir(exist_ok=True)
    else:
        server, sender, login_error = connect_smtp()
        if login_error:
            print(f"\nWARNING: {login_error}. No emails were sent.")

    first_email = True
    for gstin, invoices in not_in_2b.groupby("GSTIN", sort=False):
        vendor = invoices["Vendor Name"].iloc[0]
        to_address = vendor_emails.get(gstin, "")
        status = ""

        if not to_address or "@" not in to_address:
            status = "No email"
        elif dry_run:
            safe_name = re.sub(r"[^A-Za-z0-9]+", "_", vendor).strip("_")
            file = preview_folder / f"{safe_name}_{gstin}.html"
            header = (f"<!-- To: {to_address} | Cc: {CONFIG['CC_EMAIL']} | "
                      f"Subject: {email_subject()} -->\n")
            file.write_text(header + email_html(vendor, gstin, invoices), encoding="utf-8")
            status = "Preview saved"
        elif login_error:
            status = f"Failed - {login_error}"
        else:
            try:
                if not first_email:
                    time.sleep(CONFIG["SECONDS_BETWEEN_EMAILS"])
                first_email = False
                message = EmailMessage()
                message["Subject"] = email_subject()
                message["From"] = sender
                message["To"] = to_address
                if CONFIG["CC_EMAIL"]:
                    message["Cc"] = CONFIG["CC_EMAIL"]
                message.set_content(email_plain_text(vendor, gstin, invoices))
                message.add_alternative(email_html(vendor, gstin, invoices), subtype="html")
                server.send_message(message)
                status = "Sent"
            except Exception as err:
                status = f"Failed - {err}"

        log.append({
            "Vendor": vendor, "GSTIN": gstin, "Email": to_address,
            "No. of Invoices": len(invoices),
            "Total Tax": float(invoices["Total Tax"].sum()),
            "Status": status,
            "Timestamp": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
        })

    if server is not None:
        try:
            server.quit()
        except Exception:
            pass
    return pd.DataFrame(log)


# =============================================================================
# Main
# =============================================================================
def main():
    print("GST 2B reconciliation -", CONFIG["TAX_PERIOD"])
    books = load_purchase_register()
    gstr2b = load_gstr2b()
    vendor_emails = load_vendor_master()
    print(f"Read {len(books)} purchase invoices and {len(gstr2b)} GSTR-2B invoices.")

    books, not_in_books = reconcile(books, gstr2b)
    sheets = build_sheets(books, gstr2b, not_in_books)
    write_workbook(CONFIG["OUTPUT_FILE"], sheets)

    email_log = send_vendor_emails(sheets["Not in 2B"], vendor_emails)
    write_workbook(CONFIG["EMAIL_LOG_FILE"], {"Email log": email_log})

    counts = books["status"].value_counts()
    itc_at_risk = sheets["Not in 2B"]["Total Tax"].sum()
    status_counts = email_log["Status"].str.split(" - ").str[0].value_counts() if len(email_log) else {}
    no_email = email_log.loc[email_log["Status"] == "No email", "Vendor"].tolist() if len(email_log) else []

    print("\n---------------- SUMMARY ----------------")
    print(f"Matched                : {counts.get('Matched', 0)}")
    print(f"Value mismatch         : {counts.get('Value mismatch', 0)}")
    print(f"Probable match         : {counts.get('Probable match', 0)}")
    print(f"Not in 2B              : {counts.get('Not in 2B', 0)}")
    print(f"Not in books (2B only) : {len(not_in_books)}")
    print(f"ITC at risk            : Rs {inr(itc_at_risk)}")
    if CONFIG["DRY_RUN"]:
        print(f"Emails                 : DRY RUN - {status_counts.get('Preview saved', 0)} "
              f"preview(s) saved in '{Path(CONFIG['EMAIL_PREVIEW_FOLDER']).name}', nothing sent")
    else:
        print(f"Emails sent            : {status_counts.get('Sent', 0)}"
              f"   Failed: {status_counts.get('Failed', 0)}")
    if no_email:
        print(f"Skipped (no email)     : {', '.join(no_email)}")
    print(f"\nOutput : {Path(CONFIG['OUTPUT_FILE']).name}")
    print(f"Log    : {Path(CONFIG['EMAIL_LOG_FILE']).name}")


if __name__ == "__main__":
    main()
