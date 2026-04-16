import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import io
from payments.iban_validator import validate_iban

PAYABLE_EXCLUDE = [0, 10]

COUNTRY_SPECIAL = {
    'BE': {'sodexo_payable': 5,    'sodexo_swift': None},
    'NL': {'sodexo_payable': 5,    'sodexo_swift': None},
    'PL': {'sodexo_payable': 8,    'sodexo_swift': 'SODEXO'},
}

def validate(df_emea, country_code, emea_filter_code, generated_ids, generated_totals, sodexo_exclude=False):
    df = df_emea[df_emea['Country'].str.strip().str.upper() == emea_filter_code.upper()].copy()
    df['effective_id'] = df['CustomerID'].astype(str).str.strip()
    df['PayableTy']    = pd.to_numeric(df['PayableTy'], errors='coerce')
    df['Field11']      = pd.to_numeric(df['Field11'],   errors='coerce')
    df['Amount']       = pd.to_numeric(df['Amount'],    errors='coerce').fillna(0)

    emea_totals = df.groupby('effective_id')['Amount'].sum().round(2)
    names       = df.groupby('effective_id')['DepositName'].first()
    ibans       = df.groupby('effective_id')['IBAN'].first()
    bill_counts = df.groupby('effective_id').size()

    special    = COUNTRY_SPECIAL.get(country_code, {})
    sodexo_pay = special.get('sodexo_payable', None)
    sodexo_swft= special.get('sodexo_swift', None)

    exclusions_normal = []
    anomalies          = []
    all_emea_ids      = set(df['effective_id'].unique())

    for cid in all_emea_ids:
        rows       = df[df['effective_id'] == cid]
        total_emea = round(emea_totals.get(cid, 0), 2)

        payable_vals      = rows['PayableTy'].dropna().unique().tolist()
        field11_vals      = rows['Field11'].dropna().unique().tolist()
        has_payable_0_10  = any(v in PAYABLE_EXCLUDE for v in payable_vals)
        has_field11_block = any(v not in [3] for v in field11_vals if pd.notna(v))

        has_sodexo = False
        if sodexo_exclude and 5 in payable_vals:
            has_sodexo = True
        if sodexo_pay is not None and sodexo_pay in payable_vals:
            has_sodexo = True
        if sodexo_swft is not None:
            swift_vals = rows['SwiftCode'].astype(str).str.strip().str.upper()
            if swift_vals.eq(sodexo_swft.upper()).any():
                has_sodexo = True

        iban = str(rows['IBAN'].iloc[0]).strip()
        acct = str(rows['DepositAccountNumber'].iloc[0]).strip()
        has_bank  = (
            iban.upper() not in ('NULL', 'NAN', '') or
            (acct.upper() not in ('NULL', 'NAN', '') and acct.strip('0') != '')
        )
        iban_missing = iban.upper() in ('NULL', 'NAN', '')

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
            # --- NUOVA GERARCHIA ESCLUSIONI ---
            # 1. Importo nullo o negativo
            if total_emea <= 0:
                exclusions_normal.append({'CustomerID': cid, 'Reason': f'Negative or zero amount ({total_emea})', 'EMEA Amount': total_emea})
            # 2. Hold
            elif has_payable_0_10:
                exclusions_normal.append({'CustomerID': cid, 'Reason': 'Payable excluded (0 or 10 = hold)', 'EMEA Amount': total_emea})
            # 3. Sodexo / Pagamento esterno
            elif has_sodexo:
                exclusions_normal.append({'CustomerID': cid, 'Reason': 'Sodexo/Third Party payment (not JPM)', 'EMEA Amount': total_emea})
            # 4. Field11
            elif has_field11_block:
                exclusions_normal.append({'CustomerID': cid, 'Reason': f'Field11 blocked (value: {field11_vals})', 'EMEA Amount': total_emea})
            # 5. Dati bancari mancanti
            elif not has_bank or iban_missing:
                exclusions_normal.append({'CustomerID': cid, 'Reason': 'Missing bank details or empty IBAN', 'EMEA Amount': total_emea})
            # 6. Altro
            else:
                anomalies.append({
                    'CustomerID':       cid,
                    'Type':             'Excluded without clear reason',
                    'EMEA Amount':      total_emea,
                    'Generated Amount': 0,
                    'Difference':       -total_emea,
                    'Detail':           'Present in EMEA but not in generated file without known reason',
                })

    for cid in set(generated_ids) - all_emea_ids:
        anomalies.append({
            'CustomerID':       cid,
            'Type':             'ID not found in EMEA',
            'EMEA Amount':      0,
            'Generated Amount': generated_totals.get(cid, 0),
            'Difference':       generated_totals.get(cid, 0),
            'Detail':           'CustomerID in generated file but not found in EMEA',
        })

    status = 'green' if len(anomalies) == 0 else 'red'
    total_emea_all = round(emea_totals.sum(), 2)
    total_gen_all  = round(sum(generated_totals.values()), 2)

    summary = {
        'status':          status,
        'total_emea':      total_emea_all,
        'total_generated': total_gen_all,
        'diff_total':      round(total_gen_all - total_emea_all, 2),
        'n_emea':          len(all_emea_ids),
        'n_generated':     len(generated_ids),
        'n_exclusions':    len(exclusions_normal),
        'n_
