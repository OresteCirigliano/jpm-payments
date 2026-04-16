import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import io
from payments.iban_validator import validate_iban

# Costanti per la logica di esclusione
PAYABLE_EXCLUDE = [0, 10]

# Mappatura specifica per paese: etichette e criteri per pagamenti non-JPM
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

    # Recupero info speciali per il paese (PayQuicker vs Sodexo)
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
            # --- NUOVA GERARCHIA DI ESCLUSIONE ---
            # 1. Priorità: Importo Nullo o Negativo (Risolve il caso Rose Doyle)
            if total_emea < 0.01:
                exclusions_normal.append({'CustomerID': cid, 'Reason': f'Negative or zero amount ({total_emea})', 'EMEA Amount': total_emea})
            
            # 2. Priorità: Hold (PayableTy 0 o 10)
            elif has_payable_0_10:
                exclusions_normal.append({'CustomerID': cid, 'Reason': 'Payable excluded (0 or 10 = hold)', 'EMEA Amount': total_emea})
            
            # 3. Priorità: Metodo Esterno (Sodexo o PayQuicker)
            elif has_third_party:
                exclusions_normal.append({'CustomerID': cid, 'Reason': f'{tp_label} payment (not JPM)', 'EMEA Amount': total_emea})
            
            # 4. Blocchi amministrativi (Field11)
            elif has_field11_block:
                exclusions_normal.append({'CustomerID': cid, 'Reason': f'Field11 blocked (value: {field11_vals})', 'EMEA Amount': total_emea})
            
            # 5. Dati mancanti
            elif not has_bank or iban_missing:
                exclusions_normal.append({'CustomerID': cid, 'Reason': 'Missing bank details or empty IBAN', 'EMEA Amount': total_emea})
            
            # 6. Fallback: Anomalia
            else:
                anomalies.append({
                    'CustomerID':       cid,
                    'Type':             'Excluded without clear reason',
                    'EMEA Amount':      total_emea,
                    'Generated Amount': 0,
                    'Difference':       -total_emea,
                    'Detail':           'Present in EMEA but not in generated file without known reason',
                })

    # Verifica ID extra nel file generato (che non esistono in EMEA)
    for cid in set(generated_ids) - all_emea_ids:
        anomalies.append({
            'CustomerID':       cid,
            'Type':             'ID not found in EMEA',
            'EMEA Amount':      0,
            'Generated Amount': generated_totals.get(cid, 0),
            'Difference':       generated_totals.get(cid, 0),
            'Detail':           'CustomerID in generated file but not found in EMEA',
        })

    # 3. Finalizzazione Report
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
        'n_anomalies':     len(anomalies),
    }

    # Validazione IBAN per la lista pagamenti
    payments_list = []
    iban_issues   = 0
    for cid in sorted(generated_ids):
        iban_raw             = str(ibans.get(cid, '')).strip()
        is_valid, emoji, msg = validate_iban(iban_raw, country_code)
        if not is_valid:
            iban_issues += 1
        payments_list.append({
            'CustomerID':   cid,
            'Name':         str(names.get(cid, '')),
            'Amount':       generated_totals.get(cid, 0),
            'Bills':        int(bill_counts.get(cid, 1)),
            'IBAN':         iban_raw,
            'IBAN Status':  emoji,
            'IBAN Detail':  msg,
        })

    summary['iban_issues'] = iban_issues
    if iban_issues > 0 and status == 'green':
        summary['status'] = 'yellow'

    buf = _build_report(summary, exclusions_normal, anomalies, payments_list, country_code)
    return summary['status'], summary, buf

# (La funzione _build_report rimane invariata, la includo per completezza)
def _build_report(summary, exclusions_normal, anomalies, payments_list, country_code):
    wb = Workbook()
    GREEN = 'FF92D050'; YELLOW = 'FFFFC000'; RED = 'FFFF0000'
    HEADER = 'FF4472C4'; WHITE = 'FFFFFFFF'; BLACK = 'FF000000'

    header_font = Font(bold=True, color=WHITE)
    header_fill = PatternFill('solid', fgColor=HEADER)
    center_align = Alignment(horizontal='center')

    def write_header(ws, cols):
        for col_idx, col_name in enumerate(cols, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font; cell.fill = header_fill; cell.alignment = center_align

    # 1. Summary
    ws1 = wb.active
    ws1.title = '1. Summary'
    
    st_map = {
        'green': ('🟢 OK — No anomalies', GREEN, WHITE),
        'yellow': ('🟡 WARNING — IBAN issues', YELLOW, BLACK),
        'red': ('🔴 WARNING — Anomalies detected!', RED, WHITE)
    }
    text, color, fcolor = st_map.get(summary['status'])
    ws1.merge_cells('A1:C1')
    ws1['A1'].value = text
    ws1['A1'].font = Font(bold=True, size=14, color=fcolor)
    ws1['A1'].fill = PatternFill('solid', fgColor=color)
    ws1['A1'].alignment = center_align

    data = [
        ('', '', ''),
        ('Country', country_code, ''),
        ('EMEA Total', summary['total_emea'], ''),
        ('Generated Total', summary['total_generated'], ''),
        ('Difference', summary['diff_total'], '⚠️' if abs(summary['diff_total']) > 0.01 else '✅'),
        ('Anomalies', summary['n_anomalies'], '⚠️' if summary['n_anomalies'] > 0 else '✅'),
    ]
    for r, (l, v, n) in enumerate(data, 2):
        ws1.cell(r, 1, l).font = Font(bold=True)
        ws1.cell(r, 2, v)
        ws1.cell(r, 3, n)
    
    ws1.column_dimensions['A'].width = 25

    # 2. Payments
    ws2 = wb.create_sheet('2. Payments')
    write_header(ws2, ['CustomerID', 'Name', 'Amount', '# Bills', 'IBAN', 'IBAN Status', 'IBAN Detail'])
    for r, p in enumerate(payments_list, 2):
        ws2.cell(r, 1, p['CustomerID']); ws2.cell(r, 2, p['Name'])
        ws2.cell(r, 3, p['Amount']); ws2.cell(r, 4, p['Bills'])
        ws2.cell(r, 5, p['IBAN']).number_format = '@'
        ws2.cell(r, 6, p['IBAN Status']); ws2.cell(r, 7, p['IBAN Detail'])
        if p['IBAN Status'] == '❌':
            for c in range(1, 8): ws2.cell(r, c).fill = PatternFill('solid', fgColor='FFFFE0E0')

    # 3. Normal Exclusions
    ws3 = wb.create_sheet('3. Normal exclusions')
    write_header(ws3, ['CustomerID', 'Exclusion reason', 'EMEA Amount'])
    for r, exc in enumerate(exclusions_normal, 2):
        ws3.cell(r, 1, exc['CustomerID']); ws3.cell(r, 2, exc['Reason']); ws3.cell(r, 3, exc['EMEA Amount'])

    # 4. Anomalies
    ws4 = wb.create_sheet('4. Anomalies')
    write_header(ws4, ['CustomerID', 'Type', 'EMEA Amount', 'Generated Amount', 'Difference', 'Detail'])
    for r, an in enumerate(anomalies, 2):
        ws4.cell(r, 1, an['CustomerID']); ws4.cell(r, 2, an['Type'])
        ws4.cell(r, 3, an['EMEA Amount']); ws4.cell(r, 4, an['Generated Amount'])
        ws4.cell(r, 5, an['Difference']); ws4.cell(r, 6, an['Detail'])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
