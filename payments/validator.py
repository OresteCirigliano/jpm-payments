import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import io

PAYABLE_EXCLUDE = [0, 10]

# Configurazione esclusioni speciali per paese
COUNTRY_SPECIAL = {
    'BE': {'sodexo_payable': 5, 'sodexo_swift': None},
    'NL': {'sodexo_payable': 5, 'sodexo_swift': None},
    'PL': {'sodexo_payable': 8, 'sodexo_swift': 'SODEXO'},
}

def validate(df_emea, country_code, emea_filter_code, generated_ids, generated_totals, sodexo_exclude=False):
    df = df_emea[df_emea['Country'].str.strip().str.upper() == emea_filter_code.upper()].copy()
    df['effective_id'] = df['CustomerID'].astype(str).str.strip()
    df['PayableTy']    = pd.to_numeric(df['PayableTy'], errors='coerce')
    df['Field11']      = pd.to_numeric(df['Field11'],   errors='coerce')
    df['Amount']       = pd.to_numeric(df['Amount'],    errors='coerce').fillna(0)

    emea_totals = df.groupby('effective_id')['Amount'].sum().round(2)
    names       = df.groupby('effective_id')['DepositName'].first()

    special     = COUNTRY_SPECIAL.get(country_code, {})
    sodexo_pay  = special.get('sodexo_payable', None)
    sodexo_swft = special.get('sodexo_swift', None)

    exclusions_normal = []
    anomalies         = []
    all_emea_ids      = set(df['effective_id'].unique())

    for cid in all_emea_ids:
        rows       = df[df['effective_id'] == cid]
        total_emea = round(emea_totals.get(cid, 0), 2)

        payable_vals      = rows['PayableTy'].dropna().unique().tolist()
        field11_vals      = rows['Field11'].dropna().unique().tolist()
        has_payable_0_10  = any(v in PAYABLE_EXCLUDE for v in payable_vals)
        has_field11_block = any(v not in [3] for v in field11_vals if pd.notna(v))

        # Sodexo check
        has_sodexo = False
        if sodexo_exclude and 5 in payable_vals:
            has_sodexo = True
        if sodexo_pay is not None and sodexo_pay in payable_vals:
            has_sodexo = True
        if sodexo_swft is not None:
            swift_vals = rows['SwiftCode'].astype(str).str.strip().str.upper()
            if swift_vals.eq(sodexo_swft.upper()).any():
                has_sodexo = True

        # Bank details
        iban = str(rows['IBAN'].iloc[0]).strip()
        acct = str(rows['DepositAccountNumber'].iloc[0]).strip()
        has_bank = (
            iban.upper() not in ('NULL', 'NAN', '') or
            (acct.upper() not in ('NULL', 'NAN', '') and acct.strip('0') != '')
        )

        # IBAN vuoto = esclusione normale per PL
        iban_missing = iban.upper() in ('NULL', 'NAN', '')

        if cid in generated_ids:
            total_gen = round(generated_totals.get(cid, 0), 2)
            diff      = round(total_gen - total_emea, 2)
            if abs(diff) > 0.01:
                anomalies.append({
                    'CustomerID':       cid,
                    'Tipo':             'Discrepanza importo',
                    'Importo EMEA':     total_emea,
                    'Importo Generato': total_gen,
                    'Differenza':       diff,
                    'Dettaglio':        f'Atteso {total_emea}, generato {total_gen}',
                })
        else:
            if has_payable_0_10:
                exclusions_normal.append({'CustomerID': cid, 'Motivo': 'Payable escluso (0 o 10 = hold)', 'Importo EMEA': total_emea})
            elif has_sodexo:
                exclusions_normal.append({'CustomerID': cid, 'Motivo': 'Pagamento Sodexo (non JPM)', 'Importo EMEA': total_emea})
            elif has_field11_block:
                exclusions_normal.append({'CustomerID': cid, 'Motivo': f'Field11 bloccato (valore: {field11_vals})', 'Importo EMEA': total_emea})
            elif not has_bank or iban_missing:
                exclusions_normal.append({'CustomerID': cid, 'Motivo': 'Bank details mancanti o IBAN vuoto', 'Importo EMEA': total_emea})
            elif total_emea <= 0:
                exclusions_normal.append({'CustomerID': cid, 'Motivo': f'Importo negativo o zero ({total_emea})', 'Importo EMEA': total_emea})
            else:
                anomalies.append({
                    'CustomerID':       cid,
                    'Tipo':             'Escluso senza motivo chiaro',
                    'Importo EMEA':     total_emea,
                    'Importo Generato': 0,
                    'Differenza':       -total_emea,
                    'Dettaglio':        'Presente in EMEA ma non nel file generato senza motivo noto',
                })

    for cid in set(generated_ids) - all_emea_ids:
        anomalies.append({
            'CustomerID':       cid,
            'Tipo':             'ID non presente in EMEA',
            'Importo EMEA':     0,
            'Importo Generato': generated_totals.get(cid, 0),
            'Differenza':       generated_totals.get(cid, 0),
            'Dettaglio':        'CustomerID nel file generato ma assente nel file EMEA',
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
        'n_anomalies':     len(anomalies),
    }

    payments_list = []
    for cid in sorted(generated_ids):
        payments_list.append({
            'CustomerID': cid,
            'Nome':       str(names.get(cid, '')),
            'Importo':    generated_totals.get(cid, 0),
        })

    buf = _build_report(summary, exclusions_normal, anomalies, payments_list, country_code)
    return status, summary, buf


def _build_report(summary, exclusions_normal, anomalies, payments_list, country_code):
    wb = Workbook()

    GREEN  = 'FF92D050'
    RED    = 'FFFF0000'
    HEADER = 'FF4472C4'
    WHITE  = 'FFFFFFFF'

    header_font  = Font(bold=True, color=WHITE)
    header_fill  = PatternFill('solid', fgColor=HEADER)
    center_align = Alignment(horizontal='center')

    def write_header(ws, cols):
        for col_idx, col_name in enumerate(cols, 1):
            cell           = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font      = header_font
            cell.fill      = header_fill
            cell.alignment = center_align
        ws.row_dimensions[1].height = 20

    # ── Foglio 1: Riepilogo ────────────────────────────────
    ws1 = wb.active
    ws1.title = '1. Riepilogo'

    status_text = '🟢 OK — Nessuna anomalia' if summary['status'] == 'green' else '🔴 ATTENZIONE — Anomalie rilevate'
    status_fill = PatternFill('solid', fgColor=GREEN if summary['status'] == 'green' else RED)

    ws1.merge_cells('A1:C1')
    cell           = ws1['A1']
    cell.value     = status_text
    cell.font      = Font(bold=True, size=14, color=WHITE)
    cell.fill      = status_fill
    cell.alignment = center_align
    ws1.row_dimensions[1].height = 30

    data = [
        ('', '', ''),
        ('Paese', country_code, ''),
        ('', '', ''),
        ('Totale EMEA (tutti)',   summary['total_emea'],     ''),
        ('Totale file generato',  summary['total_generated'], ''),
        ('Differenza',            summary['diff_total'],      '⚠️' if abs(summary['diff_total']) > 0.01 else '✅'),
        ('', '', ''),
        ('Incaricati in EMEA',    summary['n_emea'],         ''),
        ('Incaricati nel file',   summary['n_generated'],    ''),
        ('Esclusi (motivo noto)', summary['n_exclusions'],   ''),
        ('Anomalie',              summary['n_anomalies'],    '⚠️' if summary['n_anomalies'] > 0 else '✅'),
    ]

    for row_idx, (label, value, note) in enumerate(data, 2):
        ws1.cell(row=row_idx, column=1, value=label).font = Font(bold=True)
        ws1.cell(row=row_idx, column=2, value=value)
        ws1.cell(row=row_idx, column=3, value=note)

    ws1.column_dimensions['A'].width = 30
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 10

    # ── Foglio 2: Pagamenti ────────────────────────────────
    ws2 = wb.create_sheet('2. Pagamenti')
    write_header(ws2, ['CustomerID', 'Nome', 'Importo'])

    for row_idx, p in enumerate(payments_list, 2):
        ws2.cell(row=row_idx, column=1, value=p['CustomerID'])
        ws2.cell(row=row_idx, column=2, value=p['Nome'])
        ws2.cell(row=row_idx, column=3, value=p['Importo'])

    last_row = len(payments_list) + 2
    ws2.cell(row=last_row, column=1, value='TOTALE').font = Font(bold=True)
    ws2.cell(row=last_row, column=3, value=round(sum(p['Importo'] for p in payments_list), 2)).font = Font(bold=True)

    ws2.column_dimensions['A'].width = 15
    ws2.column_dimensions['B'].width = 35
    ws2.column_dimensions['C'].width = 15

    # ── Foglio 3: Esclusioni normali ───────────────────────
    ws3 = wb.create_sheet('3. Esclusioni normali')
    write_header(ws3, ['CustomerID', 'Motivo esclusione', 'Importo EMEA'])

    for row_idx, exc in enumerate(exclusions_normal, 2):
        ws3.cell(row=row_idx, column=1, value=exc['CustomerID'])
        ws3.cell(row=row_idx, column=2, value=exc['Motivo'])
        ws3.cell(row=row_idx, column=3, value=exc['Importo EMEA'])

    ws3.column_dimensions['A'].width = 15
    ws3.column_dimensions['B'].width = 40
    ws3.column_dimensions['C'].width = 15

    # ── Foglio 4: Anomalie ─────────────────────────────────
    ws4 = wb.create_sheet('4. Anomalie')
    write_header(ws4, ['CustomerID', 'Tipo', 'Importo EMEA', 'Importo Generato', 'Differenza', 'Dettaglio'])

    for row_idx, an in enumerate(anomalies, 2):
        ws4.cell(row=row_idx, column=1, value=an['CustomerID'])
        ws4.cell(row=row_idx, column=2, value=an['Tipo'])
        ws4.cell(row=row_idx, column=3, value=an['Importo EMEA'])
        ws4.cell(row=row_idx, column=4, value=an['Importo Generato'])
        ws4.cell(row=row_idx, column=5, value=an['Differenza'])
        ws4.cell(row=row_idx, column=6, value=an['Dettaglio'])

    ws4.column_dimensions['A'].width = 15
    ws4.column_dimensions['B'].width = 30
    ws4.column_dimensions['C'].width = 15
    ws4.column_dimensions['D'].width = 18
    ws4.column_dimensions['E'].width = 15
    ws4.column_dimensions['F'].width = 50

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
