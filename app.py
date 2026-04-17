import streamlit as st
import pandas as pd
from datetime import datetime
import json
import os

from payments.utils import read_and_rename, MONTH_ABBREV
from payments import gb, ch
from payments.euro import generate as euro_generate, EURO_COUNTRIES
from payments.nordic import generate as nordic_generate
from payments.pl import generate as pl_generate
from payments.ae import generate as ae_generate
from payments.validator import validate

# 1. Configurazione Pagina
st.set_page_config(
    page_title="Commission File Generator",
    page_icon="🏦",
    layout="centered"
)

# 2. Database Utenti
USER_DB = {
    "oreste.cirigliano@juiceplus.com": "Comm2026",
    "elisa.galimberti@juiceplus.com": "Comm2026",
    "cristiana.desimoi@juiceplus.com": "Comm2026"
}

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if st.session_state["authenticated"]:
        return True

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.title("🔐 Accesso Riservato")
    st.info("Inserisci le tue credenziali per accedere al tool.")
    
    email_input = st.text_input("Email").lower().strip()
    pass_input = st.text_input("Password", type="password")
    
    if st.button("Accedi"):
        if email_input in USER_DB and USER_DB[email_input] == pass_input:
            st.session_state["authenticated"] = True
            st.success("Accesso effettuato!")
            st.rerun()
        else:
            st.error("Email o Password errati.")
    return False

# 3. Logica Applicativa
if check_password():
    st.markdown("""
    <style>
        .stButton>button { width: 100%; height: 50px; font-size: 16px; }
        .green-box { background-color: #92D050; padding: 15px; border-radius: 8px; color: white; font-weight: bold; text-align: center; }
        .red-box   { background-color: #FF0000; padding: 15px; border-radius: 8px; color: white; font-weight: bold; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

    if st.sidebar.button("Log out"):
        st.session_state["authenticated"] = False
        st.rerun()

    LOG_FILE = "payment_log.json"
    def load_log():
        if os.path.exists(LOG_FILE):
            try:
                with open(LOG_FILE, 'r') as f: return json.load(f)
            except: return []
        return []

    def save_log(entry):
        log = load_log()
        log.append(entry)
        with open(LOG_FILE, 'w') as f: json.dump(log, f, indent=4)

    st.title("🏦 Commission File Generator")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        country_code = st.selectbox("Country", ["GB", "CH", "PL", "AE"] + EURO_COUNTRIES + ["Nordic"])
    with col2:
        month = st.selectbox("Month", list(MONTH_ABBREV.keys()), index=datetime.now().month - 1)
    with col3:
        payment_date = st.text_input("Date (YYYYMMDD)", datetime.now().strftime("%Y%m%d"))

    uploaded_file = st.file_uploader("Upload EMEA source file", type=["xlsx", "xls"])
    payquicker_excl = st.checkbox("Exclude PayQuicker/Sodexo payments", value=True)

    if uploaded_file and st.button("Generate Payment File"):
        with st.spinner("Processing..."):
            try:
                df = read_and_rename(uploaded_file)
                emea_code = country_code # Default

                # --- LOGICA DI GENERAZIONE ---
                if country_code == "GB":
                    buf, filename, num_tr, total, currency, gen_ids, gen_totals = gb.generate(df, month, payment_date)
                elif country_code == "CH":
                    buf, filename, num_tr, total, currency, gen_ids, gen_totals = ch.generate(df, month, payment_date)
                elif country_code == "PL":
                    buf, filename, num_tr, total, currency, gen_ids, gen_totals = pl_generate(df, month, payment_date)
                elif country_code == "AE":
                    buf, filename, num_tr, total, currency, gen_ids, gen_totals = ae_generate(df, month, payment_date)
                elif country_code in EURO_COUNTRIES:
                    buf, filename, num_tr, total, currency, gen_ids, gen_totals = euro_generate(df, country_code, month, payment_date)
                elif country_code == "Nordic":
                    buf, filename, num_tr, total, currency, gen_ids, gen_totals = nordic_generate(df, month, payment_date)
                else:
                    raise ValueError("Country non supportato")

                # --- VALIDAZIONE ---
                status, summary, buf_report = validate(df, country_code, emea_code, gen_ids, gen_totals, payquicker_excl)

                st.markdown("---")
                if status == 'green':
                    st.markdown(f'<div class="green-box">✅ {country_code} SUCCESS: {num_tr} Trx | Total: {currency} {total:,.2f}</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="red-box">⚠️ {country_code} CHECK REPORT: Anomalies or IBAN issues detected</div>', unsafe_allow
