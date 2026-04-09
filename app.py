import streamlit as st
import pandas as pd
from datetime import datetime
import json
import os

from payments.utils import read_and_rename, MONTH_ABBREV
from payments import gb, ch

# ── Configurazione pagina ──────────────────────────────────
st.set_page_config(
    page_title="JPMorgan Payment Generator",
    page_icon="💳",
    layout="centered"
)

# ── Stile ──────────────────────────────────────────────────
st.markdown("""
<style>
    .main { max-width: 700px; }
    .stButton>button { width: 100%; height: 50px; font-size: 16px; }
    .log-box { background: #f8f9fa; border-radius: 8px; padding: 15px;
               font-family: monospace; font-size: 13px; }
</style>
""", unsafe_allow_html=True)

# ── Log persistence ────────────────────────────────────────
LOG_FILE = "payment_log.json"

def load_log():
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, 'r') as f:
            return json.load(f)
    return []

def save_log(entry):
    log = load_log()
    log.insert(0, entry)
    log = log[:50]  # Tieni solo gli ultimi 50
    with open(LOG_FILE, 'w') as f:
        json.dump(log, f, indent=2)

# ── UI ─────────────────────────────────────────────────────
st.title("💳 JPMorgan Payment Generator")
st.markdown("---")

# Caricamento file
st.subheader("1. Carica il file EMEA")
uploaded_file = st.file_uploader("Seleziona il file Excel mensile", type=['xlsx'])

if uploaded_file:
    st.success(f"✅ File caricato: **{uploaded_file.name}**")

st.markdown("---")

# Parametri
st.subheader("2. Parametri di pagamento")

col1, col2 = st.columns(2)

with col1:
    country = st.selectbox(
        "Paese",
        options=["GB 🇬🇧 United Kingdom", "CH 🇨🇭 Switzerland"],
    )
    country_code = country[:2]

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

st.markdown("---")

# Bottone genera
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
                    month_abbrev = MONTH_ABBREV[month]
                    buf, num_tr, total, currency = gb.generate(df, payment_date, month_abbrev)
                    filename = f"GB_payments_{payment_date}.xlsx"
                else:
                    buf, num_tr, total, currency = ch.generate(df, payment_date, month)
                    filename = f"CH_payments_{payment_date}.xlsx"

                # Mostra risultato
                st.success("✅ File generato con successo!")
                col1, col2, col3 = st.columns(3)
                col1.metric("Paese", country_code)
                col2.metric("Transazioni", f"{num_tr:,}")
                col3.metric("Totale", f"{currency} {total:,.2f}")

                # Download
                st.download_button(
                    label="📥 Scarica file Excel",
                    data=buf,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

                # Salva log
                log_entry = {
                    "data":         datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "paese":        country_code,
                    "mese":         month,
                    "data_pag":     payment_date,
                    "transazioni":  num_tr,
                    "totale":       f"{currency} {total:,.2f}",
                    "file_input":   uploaded_file.name,
                    "file_output":  filename,
                }
                save_log(log_entry)

            except Exception as e:
                st.error(f"❌ Errore: {e}")

# ── Log ────────────────────────────────────────────────────
st.markdown("---")
st.subheader("📋 Storico pagamenti")

log = load_log()
if not log:
    st.info("Nessun pagamento generato ancora.")
else:
    df_log = pd.DataFrame(log)
    df_log.columns = ["Data/Ora", "Paese", "Mese", "Data Pag.", "Transazioni", "Totale", "File Input", "File Output"]
    st.dataframe(df_log, use_container_width=True, hide_index=True)
