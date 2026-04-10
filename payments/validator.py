import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import io

PAYABLE_EXCLUDE = [0, 10]

def validate(df_emea, country_code, emea_filter_code, generated_ids, generated_totals, sodexo_exclude=False):
    """
    Confronta il file EMEA con il file generato e produce un report di validazione.
    
    Args:
        df_emea         : DataFrame del file EMEA (già rinominato)
        country_code    : codice paese usato nell'app (es. 'GB', 'GER')
        emea_filter_code: codice usato per filtrare il file EMEA (es. 'DE' per GER)
        generated_ids   : set di CustomerID presenti nel file generato
        generated_totals: dict {CustomerID: importo} dal file generato
        sodexo_exclude  : True se PayableTy 5 va escluso (BE, NL)
    
    Returns:
        (status, summary, buf_report)
    """

    # ── Filtro paese ───────────────────────────────────────
    df = df_emea[df_emea['Country'].str.strip().str.upper() == emea_filter_code.upper()].copy()
    df['effective_id'] = df['CustomerID'].astype(str).str.strip()
    df['PayableTy']    = pd.to_numeric(df['PayableTy'], errors='coerce')
    df['Field11']      = pd.to_numeric(df['Field11'],   errors='coerce')
    df['Amount']       = pd.to_numeric(df['Amount'],    errors='coerce').fillna(0)

    # ── Calcola importo atteso per CustomerID ──────────────
    emea_totals = df.groupby('effective_id')['Amount'].sum().round(2)

    # ── Classifica ogni CustomerID ─────────────────────────
    exclusions_normal = []   # esclusi per motivi normali
    anomalies         = []   # casi gravi

    all_emea_ids = set(df['effective_id'].unique())

    for cid in all_emea_ids:
        rows = df[df['effective_id'] == cid]
        total_emea = round(emea_totals.get(cid, 0), 2)

        # Controlla motivi esclusione
        payable_vals  = rows['PayableTy'].dropna().unique().tolist()
        field11_vals  = rows['Field11'].dropna().unique().tolist()
        has_payable_0_10 = any(v in PAYABLE_EXCLUDE for v in payable_vals)
        has_payable_5    = 5 in payable_vals
        has_field11_block = any(v not in [3] for v in field11_vals if pd.notna(v))

        # Bank details
        iban    = str(rows['IBAN'].iloc[0]).strip()
        acct    = str(rows['DepositAccountNumber'].iloc[0]).strip()
        routing = rows['DepositRoutingNumber'].astype(str).str.strip().iloc[0]
        has_bank = (
            (iban.upper() not in ('NULL', 'NAN', '')) or
            (acct.upper() not in ('NULL', 'NAN', '') and acct.strip('0') != '')
        )

        if cid in generated_ids:
            # CustomerID nel file generato — controlla importo
            total_gen = round(generated_totals.get(cid, 0), 2)
            diff      = round(total_gen - total_emea, 2)
            if abs(diff) > 0.01:
                anomalies.append({
                    'CustomerID':     cid,
                    'Tipo':           'Discrepanza importo',
                    'Importo EMEA':   total_emea,
                    'Importo Generato': total_gen,
                    'Differenza':     diff,
                    'Dettaglio':      f'Atteso {total_emea}, generato {total_gen}',
                })
        else:
            # CustomerID NON nel file generato — perché?
            if has_payable_0_10:
                exclusions_normal.append({
                    'CustomerID': cid,
                    'Motivo':     'Payable escluso (0 o 10 = hold)',
                    'Importo EMEA': total_emea,
                })
            elif sodexo_exclude and has_payable_5:
                exclusions_normal.append({
                    'CustomerID': cid,
                    'Motivo':     'Payable 5 escluso (Sodexo)',
                    'Importo EMEA': total_emea,
                })
            elif has_field11_block:
                exclusions_normal.append({
                    'CustomerID': cid,
                    'Motivo':     f'Field11 bloccato (valore: {field11_vals})',
                    'Importo EMEA': total_emea,
                })
            elif not has_bank:
                exclusions_normal.append({
                    'CustomerID': cid,
                    'Motivo':     'Bank details mancanti',
                    'Importo EMEA': total_emea,
                })
            elif total_emea <= 0:
                exclusions_normal.append({
                    'CustomerID': cid,
                    'Motivo':     f'Importo negativo o zero ({total_emea})',
                    'Importo EMEA': total_emea,
                })
            else:
                anomalies.append({
                    'CustomerID':       cid,
                    'Tipo':             'Escluso senza motivo chiaro',
                    'Importo EMEA':     total_emea,
                    'Importo Generato': 0,
                    'Differenza':       -total_emea,
                    'Dettaglio':        'Presente in EMEA ma non nel file generato senza motivo noto',
                })

    # CustomerID nel generato ma NON in EMEA
    phantom_ids = set(generated_ids) - all_emea_ids
    for cid in phantom_ids:
        anomalies.append({
            'CustomerID':       cid,
            'Tipo':             'ID non presente in EMEA',
            'Importo EMEA':     0,
            'Importo Generato': generated_totals.get(cid, 0),
            'Differenza':       generated_totals.get(cid, 0),
            'Dettaglio':        'CustomerID nel file generato ma assente nel file EMEA',
        })

    # ── Semaforo ───────────────────────────────────────────
    status = 'green' if len(anomalies) == 0 else 'red'

    # ── Totali riepilogo ───────────────────────────────────
    total_emea_all = round(emea_totals.sum(), 2)
    total_gen_all  = round(sum(generated_totals.values()), 2)

    summary = {
        'status':            status,
        'total_emea':        total_emea_all,
        'total_generated':   total_gen_all,
        'diff_total':        round(total_gen_all - total_emea_all, 2),
        'n_emea':            len(all_emea_ids),
        'n_generated':       len(generated_ids),
        'n_exclusions':      len(exclusions_normal),
        'n_anomalies':       len(anomalies),
    }

    # ── Crea report Excel ──────────────────────────────────
    buf = _build_report(summary, exclusions_normal, anomalies, country_code)

    return status, summary, buf


def _build_report(summary, exclusions_normal, anomalies, country_code):
    wb = Workbook()

    # Colori
    GREEN  = 'FF92D050'
    RED    = 'FFFF0000'
    GREY   = 'FFD9D9D9'
    HEADER = 'FF4472C4'
    WHITE  = 'FFFFFFFF'

    header_font  = Font(bold=True, color=WHITE)
    header_fill  = PatternFill('solid', fgColor=HEADER)
    center_align = Alignment(horizontal='center')

    def write_header(ws, cols):
        for col_idx, col_name in enumerate(cols, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font  = header_font
            cell.fill  = header_fill
            cell.alignment = center_align
        ws.row_dimensions[1].height = 20

    # ── Foglio 1: Riepilogo ────────────────────────────────
    ws1 = wb.active
    ws1.title = '1. Riepilogo'

    # Semaforo
    status_text = '🟢 OK — Nessuna anomalia' if summary['status'] == 'green' else '🔴 ATTENZIONE — Anomalie rilevate'
    status_fill = PatternFill('solid', fgColor=GREEN if summary['status'] == 'green' else RED)

    ws1.merge_cells('A1:C1')
    cell = ws1['A1']
    cell.value     = status_text
    cell.font      = Font(bold=True, size=14, color=WHITE)
    cell.fill      = status_fill
    cell.alignment = center_align
    ws1.row_dimensions[1].height = 30

    data = [
        ('', '', ''),
        ('Paese', country_code, ''),
        ('', '', ''),
        ('Totale EMEA (tutti)',     summary['total_emea'],      ''),
        ('Totale file generato',    summary['total_generated'],  ''),
        ('Differenza',              summary['diff_total'],       '⚠️' if abs(summary['diff_total']) > 0.01 else '✅'),
        ('', '', ''),
        ('Incaricati in EMEA',      summary['n_emea'],          ''),
        ('Incaricati nel file',     summary['n_generated'],     ''),
        ('Esclusi (motivo noto)',   summary['n_exclusions'],    ''),
        ('Anomalie',                summary['n_anomalies'],     '⚠️' if summary['n_anomalies'] > 0 else '✅'),
    ]

    for row_idx, (label, value, note) in enumerate(data, 2):
        ws1.cell(row=row_idx, column=1, value=label).font = Font(bold=True)
        ws1.cell(row=row_idx, column=2, value=value)
        ws1.cell(row=row_idx, column=3, value=note)

    ws1.column_dimensions['A'].width = 30
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 10

    # ── Foglio 2: Esclusioni normali ───────────────────────
    ws2 = wb.create_sheet('2. Esclusioni normali')
    cols2 = ['CustomerID', 'Motivo esclusione', 'Importo EMEA']
    write_header(ws2, cols2)

    for row_idx, exc in enumerate(exclusions_normal, 2):
        ws2.cell(row=row_idx, column=1, value=exc['CustomerID'])
        ws2.cell(row=row_idx, column=2, value=exc['Motivo'])
        ws2.cell(row=row_idx, column=3, value=exc['Importo EMEA'])

    ws2.column_dimensions['A'].width = 15
    ws2.column_dimensions['B'].width = 40
    ws2.column_dimensions['C'].width = 15

    # ── Foglio 3: Anomalie ─────────────────────────────────
    ws3 = wb.create_sheet('3. Anomalie')
    cols3 = ['CustomerID', 'Tipo', 'Importo EMEA', 'Importo Generato', 'Differenza', 'Dettaglio']
    write_header(ws3, cols3)

    for row_idx, an in enumerate(anomalies, 2):
        ws3.cell(row=row_idx, column=1, value=an['CustomerID'])
        ws3.cell(row=row_idx, column=2, value=an['Tipo'])
        ws3.cell(row=row_idx, column=3, value=an['Importo EMEA'])
        ws3.cell(row=row_idx, column=4, value=an['Importo Generato'])
        ws3.cell(row=row_idx, column=5, value=an['Differenza'])
        ws3.cell(row=row_idx, column=6, value=an['Dettaglio'])

    ws3.column_dimensions['A'].width = 15
    ws3.column_dimensions['B'].width = 30
    ws3.column_dimensions['C'].width = 15
    ws3.column_dimensions['D'].width = 18
    ws3.column_dimensions['E'].width = 15
    ws3.column_dimensions['F'].width = 50

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
