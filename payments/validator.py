import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import io
from payments.iban_validator import validate_iban

# Costanti per la logica di esclusione
PAYABLE_EXCLUDE = [0, 10]

# Mappatura specifica per paese
COUNTRY_SPECIAL = {
    'BE': {'third_party_label': 'PayQuicker', 'payable_type': 5, 'swift': None},
    'NL': {'third_party_label': 'PayQuicker', 'payable_type': 5, 'swift': None},
    'PL': {'third_party_label': 'Sodexo',     'payable_type': 8, 'swift': 'SODEXO'},
}

def validate(df_emea, country_code, emea_filter_code, generated_ids, generated_totals, sodexo_exclude=False):
    # 1. Preparazione Dati
    df = df_emea[df_emea['Country'].str.strip().str.upper() == emea_filter_code.upper()].copy()
    df['effective_id'] = df['CustomerID'].astype(str).str.strip()
    df['PayableTy']    = pd.to_numeric(df['PayableTy'], errors='coerce')
    df['Field11']      = pd.to_numeric(df['Field11'],   errors='coerce')
    df['Amount']       = pd.to_numeric(df['Amount'],    errors='coerce').fillna(0)

    # Raggruppamento per calcolare i totali EMEA
    emea_totals = df.groupby('effective_id')['Amount'].sum().round(2)
    names       = df.groupby('effective_id')['DepositName'].first()
    ibans       = df.groupby('effective_id')['IBAN'].first()
    bill_counts = df.groupby('effective_id').size()

    # Recupero info speciali per il paese
    special = COUNTRY_SPECIAL.get(country_code, {})
    tp_label = special.get('third_party_label', 'Third Party')
    tp_pay   = special.get('payable_type', None)
    tp_swft  = special.get('swift', None)

    exclusions_normal = []
    anomalies         = []
    all_emea_ids      = set(df['effective_id'].unique())

    # 2. Ciclo di Validazione
    for cid in all_emea_ids:
        rows       = df[df['effective_id'] == cid]
        total_emea = round(emea_totals.get(cid, 0), 2)

        # Analisi dei flag nelle righe EMEA
        payable_vals      = rows['PayableTy'].dropna().unique().tolist()
        field11_vals      = rows['Field11'].dropna().unique().tolist()
        has_payable_0_10  = any(v in PAYABLE_EXCLUDE for v in payable_vals)
        has_field11_block = any(v not in [3] for v in field11_vals if pd.notna(v))

        # Check per Sodexo / PayQuicker
        has_third_party = False
        if (sodexo_exclude or country_code in ['BE', 'NL']) and 5 in payable_vals:
            has_third_party = True
        elif tp_pay is not None and tp_pay in payable_vals:
            has_third_party = True
        elif tp_swft is not None:
            swift_vals = rows['SwiftCode'].astype(str).str.strip().str.upper()
            if swift_vals.eq(tp_swft.upper()).any():
                has_third_party = True

        # Check Coordinate Bancarie
        iban = str(rows['IBAN'].iloc[0]).strip()
        acct = str(rows['DepositAccountNumber'].iloc[0]).strip()
        has_bank  = (
            iban.upper() not in ('NULL', 'NAN', '') or
            (acct.upper() not in ('NULL', 'NAN', '') and acct.strip('0') != '')
        )
        iban_missing = iban.upper() in ('NULL', 'NAN', '')

        # --- LOGICA DI CONFRONTO ---
        if cid in generated_ids:
            total_gen = round(generated_totals.get(cid, 0), 2)
            diff      = round(total_gen - total_emea, 2)
            if abs(diff) > 0.01:
                anomalies.append({
                    'CustomerID':       cid,
                    'Type':             'Amount discrepancy',
                    'EMEA Amount':      total_emea,
                    'Generated Amount': total_gen,
                    'Difference':       diff,
                    'Detail':           f'Expected {total_emea}, generated {total_gen}',
                })
        else:
            # --- GERARCHIA DI ESCLUSIONE ---
            if total_emea < 0.01:
                exclusions_normal.append({'CustomerID': cid, 'Reason': f'Negative or zero amount ({total_emea})', 'EMEA Amount': total_emea})
            elif has_payable_0_10:
                exclusions_normal.append({'CustomerID': cid, 'Reason': 'Payable excluded (0 or 10 = hold)', 'EMEA Amount': total_emea})
            elif has_third_party:
                exclusions_normal.append({'CustomerID': cid, 'Reason': f'{tp_label} payment (not JPM)', 'EMEA Amount': total_emea})
            elif has_field11_block:
                exclusions_normal.append({'CustomerID': cid, 'Reason': f'Field11 blocked (value: {field11_vals})', 'EMEA Amount': total_emea})
            elif not has_bank or iban_missing:
                exclusions_normal.append({'CustomerID': cid, 'Reason': 'Missing bank details or empty IBAN', 'EMEA Amount': total_emea})
            else:
                anomalies.append({
                    'CustomerID':       cid,
