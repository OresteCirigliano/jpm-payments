import streamlit as st
import pandas as pd
from datetime import datetime
import json
import os

USER_DB = {
    "oreste.cirigliano@juiceplus.com": "Comm2026",
    "elisa.galimberti@juiceplus.com": "Comm2026",
    "cristiana.desimoi@juiceplus.com": "Comm2026",
    "adrian.divita@juiceplus.com": "Comm2026"
}

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if st.session_state["authenticated"]:
        return True

    st.title("🔐 Commission File Generator")
    email_input = st.text_input("Email").lower().strip()
    pass_input = st.text_input("Password", type="password")
    
    if st.button("Log in"):
        if email_input in USER_DB and USER_DB[email_input] == pass_input:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Email or Password wrong.")
    return False

if not check_password():
    st.stop()

from payments.utils import read_and_rename, MONTH_ABBREV
from payments import gb, ch
from payments.euro import generate as euro_generate, EURO_COUNTRIES
from payments.nordic import generate as nordic_generate
from payments.pl import generate as pl_generate
from payments.ae import generate as ae_generate
from payments.validator import validate

st.set_page_config(
    page_title="Commission File Generator",
    page_icon="🏦",
    layout="centered"
)

st.markdown("""
<style>
    .stButton>button { width: 100%; height: 50px; font-size: 16px; }
    .green-box { background-color: #92D050; padding: 15px; border-radius: 8px;
                 color: white; font-size: 18px; font-weight: bold; text-align: center; }
    .red-box   { background-color: #FF0000; padding: 15px; border-radius: 8px;
                 color: white; font-size: 18px; font-weight: bold; text-align: center; }
</style>
""", unsafe_allow_html=True)

if st.sidebar.button("Log out"):
    st.session_state["authenticated"] = False
    st.rerun()

LOG_FILE = "payment_log.json"

def load_log():
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, 'r') as f:
            return json.load(f)
    return []

def save_log(entry):
    log = load_log()
    log.insert(0, entry)
    with open(LOG_FILE, 'w') as f:
        json.dump(log[:50], f, indent=2)

COUNTRY_OPTIONS = [
    ("🇬🇧 United Kingdom (GB)",  "GB"),
    ("🇨🇭 Switzerland (CH)",     "CH"),
    ("🇩🇰 Denmark (DK)",         "DK"),
    ("🇸🇪 Sweden (SE)",          "SE"),
    ("🇳🇴 Norway (NO)",          "NO"),
    ("🇧🇪 Belgium (BE)",         "BE"),
    ("🇮🇪 Ireland (EIR)",        "IE"),
    ("🇪🇸 Spain (ES)",           "ES"),
    ("🇫🇮 Finland (FI)",         "FI"),
    ("🇫🇷 France (FR)",          "FR"),
    ("🇩🇪 Germany (GER)",        "GER"),
    ("🇮🇹 Italy (IT)",           "IT"),
    ("🇱🇺 Luxembourg (LU)",      "LU"),
    ("🇳🇱 Netherlands (NL)",     "NL"),
    ("🇦🇹 Austria (OS)",         "OS"),
    ("🇵🇹 Portugal (PT)",        "PT"),
    ("🇵🇱 Poland (PL)",          "PL"),
    ("🇦🇪 UAE (AE)",             "AE"),
]

EMEA_FILTER = {
    'GER': 'DE',
    'OS':  'AT',
}

EURO_CODES      = set(EURO_COUNTRIES.keys())
NORDIC_CODES    = {'DK', 'SE', 'NO'}
PayQuicker_COUNTRIES = {'BE', 'NL'}

st.title("🏦 Commission File Generator")
st.markdown("---")

st.subheader("1. Upload EMEA file")
uploaded_file = st.file_uploader("Select the monthly Excel file", type=['xlsx'])
if uploaded_file:
    st.success(f"✅ File uploaded: **{uploaded_file.name}**")

st.markdown("---")
st.subheader("2. File parameters")

col1, col2 = st.columns(2)
with col1:
    country_label = st.selectbox("Country", options=[label for label, _ in COUNTRY_OPTIONS])
    country_code  = next(code for label, code in COUNTRY_OPTIONS if label == country_label)
with col2:
    month = st.selectbox("Month", options=list(MONTH_ABBREV.keys()), index=datetime.now().month - 1)

payment_date = st.text_input("Payment date (YYYYMMDD)", value=datetime.now().strftime("%Y%m%d"))

if country_code in PayQuicker_COUNTRIES:
    st.info("ℹ️ For this country, PayableTy 5 (PayQuicker) is automatically excluded.")
if country_code == 'SE':
    st.info("ℹ️ For Sweden, the payment reference contains only the CustomerID (no month).")
if country_code == 'AE':
    st.info("ℹ️ For UAE, the bank name (column Z) is populated automatically from the Swift code where available.")

st.markdown("---")

if st.button("▶ Generate payment file + Validation report", type="primary"):
    if not uploaded_file:
        st.error("⚠️ Please upload the Excel file first!")
    elif len(payment_date) != 8 or not payment_date.isdigit():
        st.error("⚠️ Date must be in YYYYMMDD format (e.g. 20260306)")
    else:
        with st.spinner("Processing..."):
            try:
                file_bytes = uploaded_file.read()
                df = read_and_rename(file_bytes)

                if country_code == 'GB':
                    buf, num_tr, total, currency, gen_ids, gen_totals = gb.generate(df, payment_date, MONTH_ABBREV[month])
                    filename = f"GB_payments_{payment_date}.xlsx"
                elif country_code == 'CH':
                    buf, num_tr, total, currency, gen_ids, gen_totals = ch.generate(df, payment_date, month)
                    filename = f"CH_payments_{payment_date}.xlsx"
                elif country_code in NORDIC_CODES:
                    buf, num_tr, total, currency, gen_ids, gen_totals = nordic_generate(df, payment_date, month, country_code)
                    filename = f"{country_code}_payments_{payment_date}.xlsx"
                elif country_code == 'PL':
                    buf, num_tr, total, currency, gen_ids, gen_totals = pl_generate(df, payment_date, month)
                    filename = f"PL_payments_{payment_date}.xlsx"
                elif country_code == 'AE':
                    buf, num_tr, total, currency, gen_ids, gen_totals = ae_generate(df, payment_date, month)
                    filename = f"AE_payments_{payment_date}.xlsx"
                else:
                    buf, num_tr, total, currency, gen_ids, gen_totals = euro_generate(df, payment_date, month, country_code)
                    filename = f"{country_code}_payments_{payment_date}.xlsx"

                emea_code   = EMEA_FILTER.get(country_code, country_code)
                PayQuicker_excl = country_code in PayQuicker_COUNTRIES
                status, summary, buf_report = validate(
                    df, country_code, emea_code, gen_ids, gen_totals, PayQuicker_excl
                )

                if status == 'green':
                    st.markdown('<div class="green-box">🟢 Validation OK — No anomalies found</div>', unsafe_allow_html=True)
                elif status == 'yellow':
                    st.markdown('<div style="background-color:#FFC000;padding:15px;border-radius:8px;color:black;font-size:18px;font-weight:bold;text-align:center;">🟡 WARNING — IBAN issues detected, please check the Payments sheet</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="red-box">🔴 WARNING — Anomalies detected! Please check the report.</div>', unsafe_allow_html=True)

                st.markdown("")

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Country",      country_code)
                c2.metric("Transactions", f"{num_tr:,}")
                c3.metric("Total",        f"{currency} {total:,.2f}")
                c4.metric("Anomalies",    summary['n_anomalies'])

                with st.expander("📊 Validation details"):
                    col_a, col_b = st.columns(2)
                    col_a.metric("EMEA Total",      f"{currency} {summary['total_emea']:,.2f}")
                    col_b.metric("Generated Total", f"{currency} {summary['total_generated']:,.2f}")
                    col_a.metric("EMEA records",    summary['n_emea'])
                    col_b.metric("In file",         summary['n_generated'])
                    col_a.metric("Excluded (normal)", summary['n_exclusions'])
                    col_b.metric("Anomalies",       summary['n_anomalies'])

                st.markdown("---")

                st.download_button(
                    label="📥 Download payment file",
                    data=buf,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                report_filename = f"Report_{country_code}_{payment_date}.xlsx"
                st.download_button(
                    label="📋 Download validation report",
                    data=buf_report,
                    file_name=report_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                save_log({
                    "date":         datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "country":      country_code,
                    "month":        month,
                    "payment_date": payment_date,
                    "transactions": num_tr,
                    "total":        f"{currency} {total:,.2f}",
                    "anomalies":    summary['n_anomalies'],
                    "status":       status,
                    "input_file":   uploaded_file.name,
                    "output_file":  filename,
                })

            except Exception as e:
                st.error(f"❌ Error: {e}")
                import traceback
                st.code(traceback.format_exc())

st.markdown("---")
st.subheader("📋 3. Reports history")
log = load_log()
if not log:
    st.info("No reports generated yet.")
else:
    df_log = pd.DataFrame(log)
    st.dataframe(df_log, use_container_width=True, hide_index=True)
