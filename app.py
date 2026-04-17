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

# Configurazione Pagina
st.set_page_config(
    page_title="Commission File Generator",
    page_icon="🏦",
    layout="centered"
)

# Database utenti autorizzati
USER_DB = {
    "oreste.cirigliano@juiceplus.com": "Comm2026",
    "elisa.galimberti@juiceplus.com": "Comm2026",
    "cristiana.desimoi@juiceplus.com": "Comm2026"
}

def check_password():
    """Restituisce True se l'utente è autenticato, altrimenti mostra il form di login."""
    
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    # Interfaccia di Login
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.title("🔐 Accesso Riservato")
    st.info("Inserisci le tue credenziali Juice Plus+ per accedere al tool.")
    
    with st.container():
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

# Esecuzione principale
if check_password():

    # Stili CSS
    st.markdown("""
    <style>
        .stButton>button { width: 100%; height: 50px; font-size: 16px; }
        .green-box { background-color: #92D050; padding: 15px; border-radius: 8px;
                     color: white; font-size: 18px; font-weight: bold; text-align: center; }
        .red-box   { background-color: #FF0000; padding: 15px; border-radius: 8px;
                     color: white; font-size: 18px; font-weight: bold; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

    # Logout nella sidebar (opzionale)
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
        log.append(entry)
        with open(LOG_FILE, 'w') as f:
            json.dump(log, f, indent=4)

    st.title("🏦 Commission File Generator")
    st.subheader("🚀 1. Select Country and Configuration")

    col1, col2, col3 = st.columns(3)
    with col1:
        country_code = st.selectbox("Country", ["GB", "CH", "PL", "AE"] + EURO_COUNTRIES + ["Nordic"])
    with col2:
        month = st.selectbox("Commission Month", list(MONTH_ABBREV.keys()), index=datetime.now().month - 1)
    with col3:
        payment_date = st.text_input("Payment Date (YYYYMMDD)", datetime.now().strftime("%Y%m%d"))

    # Mapping per la validazione
    EMEA_CODES = {
        "GB": "GB", "CH": "CH", "PL": "PL", "AE": "AE",
        "BE": "BE", "NL": "NL", "FR": "FR", "ES": "ES", "IT": "IT",
        "DE": "DE", "AT": "AT", "IE": "IE"
    }
    emea_code = EMEA_CODES.get(country_code, country_code)

    st.subheader("📂 2. Upload EMEA Excel")
    uploaded_file = st.file_uploader("Upload EMEA source file", type=["xlsx", "xls"])
    
    # Checkbox per Sodexo/Payquicker basato sulla tua logica
    PayQuicker_excl = st.checkbox("Exclude PayQuicker/Sodexo payments (if applicable)", value=True)

    if uploaded_file and st.button("Generate Payment File"):
        with st.spinner("Processing..."):
            try:
                df = read_and_rename(uploaded_file)
                
                # Selezione logica generazione
                if country_code == "GB":
                    buf, filename, num_tr, total, currency, gen_ids, gen_totals = gb.generate(df, month, payment_date)
                elif country_code == "CH":
                    buf, filename, num_tr, total, currency, gen_ids, gen_totals = ch.generate(df, month, payment_date)
                elif country_code == "PL":
                    buf, filename, num_tr, total, currency, gen_ids, gen_totals = pl_generate(df, month, payment_date)
                elif country_code == "AE":
                    buf, filename, num_tr, total, currency, gen_ids, gen_totals = ae_generate(df, month, payment_date)
                elif country_code in EURO_COUNTRIES:
                    buf, filename, num_tr, total, currency, gen_ids, gen_totals = euro_generate
