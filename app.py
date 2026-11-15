"""
app.py - web portal for the GST 2B reconciliation
=================================================
Run locally :  streamlit run app.py
Needs       :  gst_2b_reco.py in the same folder (the matching logic lives there)

Settings (never written in this file) are read from Streamlit secrets
(.streamlit/secrets.toml) or, if not found there, from environment variables:

    GST_SENDER_EMAIL   the mailbox the emails are sent from
    GST_APP_PASSWORD   the app password of that mailbox (Gmail SMTP route)
    BREVO_API_KEY      optional - if set, emails go through Brevo's web API
                       instead of Gmail SMTP (for hosts that block SMTP)
    PORTAL_USERS       logins for the portal, e.g.  "admin:Pass@123,staff:Pass@456"
                       (in secrets.toml you can use a [users] table instead)
"""

import hmac
import io
import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import gst_2b_reco as reco


class RecoError(Exception):
    """Raised instead of ending the program when an input file has a problem."""


def _raise(message):
    raise RecoError(message)


reco.stop = _raise   # the script calls stop() on errors; on the portal we show them


# =============================================================================
# Settings and login
# =============================================================================
def setting(name, default=""):
    """Read a setting from Streamlit secrets, then from environment variables."""
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.environ.get(name, default)


def portal_users():
    """Returns {username: password}."""
    try:
        if "users" in st.secrets:
            return {str(k): str(v) for k, v in dict(st.secrets["users"]).items()}
    except Exception:
        pass
    users = {}
    for pair in str(os.environ.get("PORTAL_USERS", "")).split(","):
        if ":" in pair:
            name, password = pair.split(":", 1)
            users[name.strip()] = password.strip()
    return users


def login_screen():
    st.title("GST 2B Reconciliation Portal")
    users = portal_users()
    if not users:
        st.error("No portal users are set up. Add a [users] table to "
                 ".streamlit/secrets.toml or set the PORTAL_USERS environment variable.")
        st.stop()
    with st.form("login"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.form_submit_button("Log in"):
            expected = users.get(username.strip())
            if expected and hmac.compare_digest(password.encode(), expected.encode()):
                st.session_state["user"] = username.strip()
                st.rerun()
            else:
                st.error("Wrong username or password.")
    st.stop()


# =============================================================================
# Reconciliation (no Streamlit code here, so it can be tested on its own)
# =============================================================================
def run_reconciliation(purchase_bytes, gstr2b_bytes, vendor_bytes):
    """Runs the script's logic on uploaded files. Returns results as a dict."""
    with tempfile.TemporaryDirectory() as folder:
        folder = Path(folder)
        (folder / "purchase_register.xlsx").write_bytes(purchase_bytes)
        (folder / "gstr2b.xlsx").write_bytes(gstr2b_bytes)
        (folder / "vendor_master.xlsx").write_bytes(vendor_bytes)
        reco.CONFIG.update({
            "PURCHASE_REGISTER_FILE": folder / "purchase_register.xlsx",
            "GSTR2B_FILE": folder / "gstr2b.xlsx",
            "VENDOR_MASTER_FILE": folder / "vendor_master.xlsx",
        })
        books = reco.load_purchase_register()
        gstr2b = reco.load_gstr2b()
        vendor_emails = reco.load_vendor_master()
        books, not_in_books = reco.reconcile(books, gstr2b)
        sheets = reco.build_sheets(books, gstr2b, not_in_books)
        output = folder / "reconciliation_output.xlsx"
        reco.write_workbook(output, sheets)
        return {
            "sheets": sheets,
            "vendor_emails": vendor_emails,
            "workbook": output.read_bytes(),
            "books_count": len(books),
            "gstr2b_count": len(gstr2b),
        }


def log_to_excel(log):
    buffer = io.BytesIO()
    log.to_excel(buffer, index=False, sheet_name="Email log")
    return buffer.getvalue()


# =============================================================================
# Pages
# =============================================================================
def sidebar():
    with st.sidebar:
        st.write(f"Logged in as **{st.session_state['user']}**")
        if st.button("Log out"):
            st.session_state.clear()
            st.rerun()
        st.divider()
        st.subheader("Firm details (used in emails)")
        c = reco.CONFIG
        c["FIRM_NAME"] = st.text_input("Firm name", c["FIRM_NAME"])
        c["TAX_PERIOD"] = st.text_input("Tax period", c["TAX_PERIOD"])
        c["FIRM_CONTACT_PERSON"] = st.text_input("Contact person", c["FIRM_CONTACT_PERSON"])
        c["FIRM_PHONE"] = st.text_input("Phone", c["FIRM_PHONE"])
        c["FIRM_EMAIL"] = st.text_input("Contact email", c["FIRM_EMAIL"])
        c["CC_EMAIL"] = st.text_input("CC on every email (optional)", c["CC_EMAIL"])
        st.divider()
        st.subheader("Matching rules")
        c["TAX_TOLERANCE"] = st.number_input("Tax tolerance (Rs)", 0.0, 1000.0,
                                             float(c["TAX_TOLERANCE"]), step=1.0)
        c["DATE_TOLERANCE_DAYS"] = int(st.number_input("Date tolerance (days)", 0, 60,
                                                       int(c["DATE_TOLERANCE_DAYS"])))


def upload_section():
    st.header("1. Upload files")
    col1, col2, col3 = st.columns(3)
    purchase = col1.file_uploader("Purchase register", type=["xlsx"])
    gstr2b = col2.file_uploader("GSTR-2B (portal Excel)", type=["xlsx"])
    vendors = col3.file_uploader("Vendor master", type=["xlsx"])
    ready = purchase and gstr2b and vendors
    if st.button("Run reconciliation", type="primary", disabled=not ready):
        try:
            with st.spinner("Reconciling..."):
                st.session_state["result"] = run_reconciliation(
                    purchase.getvalue(), gstr2b.getvalue(), vendors.getvalue())
            st.session_state.pop("email_log", None)
        except RecoError as err:
            st.session_state.pop("result", None)
            st.error(str(err))
        except Exception as err:
            st.session_state.pop("result", None)
            st.error(f"Could not process the files: {err}")


def results_section(result):
    sheets = result["sheets"]
    summary = sheets["Summary"].set_index("Category")
    st.header("2. Results")
    st.caption(f"{result['books_count']} purchase invoices compared with "
               f"{result['gstr2b_count']} GSTR-2B invoices.")

    cols = st.columns(5)
    labels = ["Matched", "Value mismatch", "Probable match", "Not in 2B", "Not in books"]
    for col, label, (category, row) in zip(cols, labels, summary.iterrows()):
        col.metric(label, int(row["No. of Invoices"]),
                   f"Rs {reco.inr(row['Total Tax (ITC)'])}", delta_color="off")

    st.download_button("Download reconciliation workbook (Excel)", result["workbook"],
                       file_name="reconciliation_output.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    names = [n for n in sheets if n != "Summary"]
    for tab, name in zip(st.tabs(names), names):
        with tab:
            if sheets[name].empty:
                st.info("No invoices in this category.")
            else:
                st.dataframe(sheets[name], use_container_width=True, hide_index=True)


def email_section(result):
    st.header("3. Email vendors for invoices not in 2B")
    not_in_2b = result["sheets"]["Not in 2B"]
    if not_in_2b.empty:
        st.success("Every purchase invoice appears in GSTR-2B. No emails needed.")
        return

    emails = result["vendor_emails"]
    vendors = result["sheets"]["Vendor-wise ITC at risk"].copy()
    vendors["Email"] = vendors["GSTIN"].map(lambda g: emails.get(g, ""))
    has_email = vendors["Email"].str.contains("@", na=False)
    vendors.insert(0, "Send", has_email)

    if (~has_email).any():
        st.warning("No email address in the vendor master for: "
                   + ", ".join(vendors.loc[~has_email, "Vendor Name"])
                   + ". These vendors will be skipped.")

    edited = st.data_editor(
        vendors, hide_index=True, use_container_width=True,
        disabled=[c for c in vendors.columns if c != "Send"],
        column_config={"Send": st.column_config.CheckboxColumn("Send")},
    )
    selected = edited[edited["Send"] & edited["Email"].str.contains("@", na=False)]

    # ---- Preview -------------------------------------------------------------
    preview_name = st.selectbox("Preview the email for", vendors["Vendor Name"])
    row = vendors[vendors["Vendor Name"] == preview_name].iloc[0]
    invoices = not_in_2b[not_in_2b["GSTIN"] == row["GSTIN"]]
    st.write(f"**To:** {row['Email'] or '(no email)'}  \n**Subject:** {reco.email_subject()}")
    components.html(reco.email_html(row["Vendor Name"], row["GSTIN"], invoices),
                    height=480, scrolling=True)

    # ---- Send ----------------------------------------------------------------
    sender = str(setting("GST_SENDER_EMAIL")).strip()
    password = str(setting("GST_APP_PASSWORD")).strip()
    brevo_key = str(setting("BREVO_API_KEY")).strip()
    if not sender or not (password or brevo_key):
        st.error("Sending is switched off: set GST_SENDER_EMAIL and either "
                 "BREVO_API_KEY or GST_APP_PASSWORD in the secrets.")
        return

    route = "Brevo" if brevo_key else "Gmail SMTP"
    st.write(f"Emails will be sent from **{sender}** through {route}.")
    test_mode = st.checkbox(f"Test mode: send every email to {sender} instead of the vendor",
                            value=True)
    confirm = st.checkbox(f"I have checked the preview. Send {len(selected)} email(s).")
    if st.button("Send emails", type="primary", disabled=not confirm or selected.empty):
        os.environ[reco.CONFIG["ENV_SENDER_EMAIL"]] = sender
        os.environ[reco.CONFIG["ENV_APP_PASSWORD"]] = password
        os.environ[reco.CONFIG["ENV_BREVO_API_KEY"]] = brevo_key
        reco.CONFIG["DRY_RUN"] = False
        to_send = not_in_2b[not_in_2b["GSTIN"].isin(selected["GSTIN"])]
        addresses = ({g: sender for g in selected["GSTIN"]} if test_mode
                     else dict(zip(selected["GSTIN"], selected["Email"])))
        with st.spinner("Sending..."):
            st.session_state["email_log"] = reco.send_vendor_emails(to_send, addresses)

    log = st.session_state.get("email_log")
    if log is not None:
        sent = int((log["Status"] == "Sent").sum())
        failed = int(log["Status"].str.startswith("Failed").sum())
        (st.success if failed == 0 else st.warning)(f"Sent: {sent}   Failed: {failed}")
        st.dataframe(log, use_container_width=True, hide_index=True)
        st.download_button("Download email log (Excel)", log_to_excel(log),
                           file_name="email_log.xlsx")


def main():
    st.set_page_config(page_title="GST 2B Reconciliation", layout="wide")
    if "user" not in st.session_state:
        login_screen()
    sidebar()
    st.title("GST 2B Reconciliation Portal")
    upload_section()
    result = st.session_state.get("result")
    if result:
        results_section(result)
        email_section(result)


if __name__ == "__main__":
    main()
