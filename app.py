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

st.set_page_config(
    page_title="JPMorgan Payment Generator",
    page_icon="💳",
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
    ("🇵🇱 Poland (PL)",          "PL"),
    ("🇦🇪 UAE (AE)",             "AE"),
]

# Mappa country_code → emea_filter_code (per paesi con codice diverso nel file EMEA)
EMEA_FILTER = {
    'GER': 'DE',
    'OS':  'AT',
}

EURO_CODES   = set(EURO_COUNTRIES.keys())
NORDIC_CODES = {'DK', 'SE', 'NO'}
SODEXO_COUNTRIES = {'BE', 'NL'}

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
    country_label = st.selectbox("Paese", options=[label for label, _ in COUNTRY_OPTIONS])
    country_code  = next(code for label, code in COUNTRY_OPTIONS if label == country_label)
with col2:
    month = st.selectbox("Mese", options=list(MONTH_ABBREV.keys()), index=datetime.now().month - 1)

payment_date = st.text_input("Data pagamento (YYYYMMDD)", value=datetime.now().strftime("%Y%m%d"))

if country_code in SODEXO_COUNTRIES:
    st.info("ℹ️ Per questo paese il PayableTy 5 (Sodexo) viene escluso automaticamente.")
if country_code == 'SE':
    st.info("ℹ️ Per la Svezia il payment reference contiene solo il CustomerID (senza mese).")
if country_code == 'AE':
    st.info("ℹ️ Per UAE la colonna banca (Z) viene compilata automaticamente dal Swift code se disponibile.")

st.markdown("---")

if st.button("▶ Genera file di pagamento + Report validazione", type="primary"):
    if not uploaded_file:
        st.error("⚠️ Carica prima il file Excel!")
    elif len(payment_date) != 8 or not payment_date.isdigit():
        st.error("⚠️ La data deve essere nel formato YYYYMMDD (es. 20260306)")
    else:
        with st.spinner("Elaborazione in corso..."):
            try:
                file_bytes = uploaded_file.read()
                df = read_and_rename(file_bytes)

                # ── Genera file pagamento ──────────────────
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

                # ── Validazione ────────────────────────────
                emea_code    = EMEA_FILTER.get(country_code, country_code)
                sodexo_excl  = country_code in SODEXO_COUNTRIES
                status, summary, buf_report = validate(
                    df, country_code, emea_code, gen_ids, gen_totals, sodexo_excl
                )

                # ── Semaforo ───────────────────────────────
                if status == 'green':
                    st.markdown('<div class="green-box">🟢 Validazione OK — Nessuna anomalia</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="red-box">🔴 ATTENZIONE — Anomalie rilevate! Controlla il report.</div>', unsafe_allow_html=True)

                st.markdown("")

                # ── Metriche ───────────────────────────────
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Paese",        country_code)
                c2.metric("Transazioni",  f"{num_tr:,}")
                c3.metric("Totale",       f"{currency} {total:,.2f}")
                c4.metric("Anomalie",     summary['n_anomalies'],
                          delta=None if summary['n_anomalies'] == 0 else "⚠️")

                with st.expander("📊 Dettaglio validazione"):
                    col_a, col_b = st.columns(2)
                    col_a.metric("Totale EMEA",      f"{currency} {summary['total_emea']:,.2f}")
                    col_b.metric("Totale Generato",  f"{currency} {summary['total_generated']:,.2f}")
                    col_a.metric("Incaricati EMEA",  summary['n_emea'])
                    col_b.metric("Nel file",         summary['n_generated'])
                    col_a.metric("Esclusi (normale)", summary['n_exclusions'])
                    col_b.metric("Anomalie",         summary['n_anomalies'])

                st.markdown("---")

                # ── Download file pagamento ────────────────
                st.download_button(
                    label="📥 Scarica file pagamento",
                    data=buf,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                # ── Download report validazione ────────────
                report_filename = f"Report_{country_code}_{payment_date}.xlsx"
                st.download_button(
                    label="📋 Scarica report validazione",
                    data=buf_report,
                    file_name=report_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                save_log({
                    "data":        datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "paese":       country_code,
                    "mese":        month,
                    "data_pag":    payment_date,
                    "transazioni": num_tr,
                    "totale":      f"{currency} {total:,.2f}",
                    "anomalie":    summary['n_anomalies'],
                    "status":      status,
                    "file_input":  uploaded_file.name,
                    "file_output": filename,
                })

            except Exception as e:
                st.error(f"❌ Errore: {e}")
                import traceback
                st.code(traceback.format_exc())

st.markdown("---")
st.subheader("📋 Storico pagamenti")
log = load_log()
if not log:
    st.info("Nessun pagamento generato ancora.")
else:
    df_log = pd.DataFrame(log)
    df_log.columns = ["Data/Ora", "Paese", "Mese", "Data Pag.", "Transazioni", "Totale", "Anomalie", "Status", "File Input", "File Output"]
    st.dataframe(df_log, use_container_width=True, hide_index=True)
