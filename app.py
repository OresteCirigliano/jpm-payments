import streamlit as st
import pandas as pd
from datetime import datetime
import json
import os

from payments.utils import read_and_rename, MONTH_ABBREV
from payments import gb, ch
from payments.euro import generate as euro_generate, EURO_COUNTRIES
from payments.nordic import generate as nordic_generate

# ── Configurazione pagina ──────────────────────────────────
st.set_page_config(
    page_title="JPMorgan Payment Generator",
    page_icon="💳",
    layout="centered"
)

st.markdown("""
<style>
    .stButton>button { width: 100%; height: 50px; font-size: 16px; }
</style>
""", unsafe_allow_html=True)

# ── Log ────────────────────────────────────────────────────
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

# ── Opzioni paesi ──────────────────────────────────────────
COUNTRY_OPTIONS = [
    ("🇬🇧 United Kingdom (GB)",  "GB"),
    ("🇨🇭 Switzerland (CH)",     "CH"),
    ("🇩🇰 Denmark (DK)",         "DK"),
    ("🇸🇪 Sweden (SE)",          "SE"),
    ("🇳🇴 Norway (NO)",          "NO"),
    ("🇧🇪 Belgium (BE)",         "BE"),
    ("🇮🇪 Ireland (EIR)",        "EIR"),
    ("🇪🇸 Spain (ES)",           "ES"),
    ("🇫🇮 Finland (FI)",         "FI"),
    ("🇫🇷 France (FR)",          "FR"),
    ("🇩🇪 Germany (GER)",        "GER"),
    ("🇮🇹 Italy (IT)",           "IT"),
    ("🇱🇺 Luxembourg (LU)",      "LU"),
    ("🇳🇱 Netherlands (NL)",     "NL"),
    ("🇦🇹 Austria (OS)",         "OS"),
    ("🇵🇹 Portugal (PT)",        "PT"),
]

EURO_CODES   = set(EURO_COUNTRIES.keys())
NORDIC_CODES = {'DK', 'SE', 'NO'}

# ── UI ─────────────────────────────────────────────────────
st.title("💳 JPMorgan Payment Generator")
st.markdown("---")

st.subheader("1. Carica il file EMEA")
uploaded_file = st.file_uploader("Seleziona il file Excel mensile", type=['xlsx'])
if uploaded_file:
    st.success(f"✅ File caricato: **{uploaded_file.name}**")

st.markdown("---")
st.subheader("2. Parametri di pagamento")

col1, col2 = st.columns(2)

with col1:
    country_label = st.selectbox(
        "Paese",
        options=[label for label, _ in COUNTRY_OPTIONS],
    )
    country_code = next(code for label, code in COUNTRY_OPTIONS if label == country_label)

with col2:
    month = st.selectbox(
        "Mese",
        options=list(MONTH_ABBREV.keys()),
        index=datetime.now().month - 1
    )

payment_date = st.text_input(
    "Data pagamento (YYYYMMDD)",
    value=datetime.now().strftime("%Y%m%d")
)

if country_code in ('BE', 'NL'):
    st.info("ℹ️ Per questo paese il PayableTy 5 (Sodexo) viene escluso automaticamente.")

if country_code == 'SE':
    st.info("ℹ️ Per la Svezia il payment reference contiene solo il CustomerID (senza mese).")

st.markdown("---")

if st.button("▶ Genera file di pagamento", type="primary"):
    if not uploaded_file:
        st.error("⚠️ Carica prima il file Excel!")
    elif len(payment_date) != 8 or not payment_date.isdigit():
        st.error("⚠️ La data deve essere nel formato YYYYMMDD (es. 20260306)")
    else:
        with st.spinner("Elaborazione in corso..."):
            try:
                file_bytes = uploaded_file.read()
                df = read_and_rename(file_bytes)

                if country_code == 'GB':
                    buf, num_tr, total, currency = gb.generate(df, payment_date, MONTH_ABBREV[month])
                    filename = f"GB_payments_{payment_date}.xlsx"

                elif country_code == 'CH':
                    buf, num_tr, total, currency = ch.generate(df, payment_date, month)
                    filename = f"CH_payments_{payment_date}.xlsx"

                elif country_code in NORDIC_CODES:
                    buf, num_tr, total, currency = nordic_generate(df, payment_date, month, country_code)
                    filename = f"{country_code}_payments_{payment_date}.xlsx"

                else:
                    buf, num_tr, total, currency = euro_generate(df, payment_date, month, country_code)
                    filename = f"{country_code}_payments_{payment_date}.xlsx"

                st.success("✅ File generato con successo!")
                c1, c2, c3 = st.columns(3)
                c1.metric("Paese", country_code)
                c2.metric("Transazioni", f"{num_tr:,}")
                c3.metric("Totale", f"{currency} {total:,.2f}")

                st.download_button(
                    label="📥 Scarica file Excel",
                    data=buf,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

                save_log({
                    "data":        datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "paese":       country_code,
                    "mese":        month,
                    "data_pag":    payment_date,
                    "transazioni": num_tr,
                    "totale":      f"{currency} {total:,.2f}",
                    "file_input":  uploaded_file.name,
                    "file_output": filename,
                })

            except Exception as e:
                st.error(f"❌ Errore: {e}")

# ── Storico ────────────────────────────────────────────────
st.markdown("---")
st.subheader("📋 Storico pagamenti")
log = load_log()
if not log:
    st.info("Nessun pagamento generato ancora.")
else:
    df_log = pd.DataFrame(log)
    df_log.columns = ["Data/Ora", "Paese", "Mese", "Data Pag.", "Transazioni", "Totale", "File Input", "File Output"]
    st.dataframe(df_log, use_container_width=True, hide_index=True)
