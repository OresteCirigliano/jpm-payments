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

    special    = COUNTRY_SPECIAL.get(country_code, {})
    sodexo_pay = special.get('sodexo_payable', None)
    sodexo_swft= special.get('sodexo_swift', None)

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
            if has_payable_0_10:
                exclusions_normal.append({'CustomerID': cid, 'Reason': 'Payable excluded (0 or 10 = hold)', 'EMEA Amount': total_emea})
            elif has_sodexo:
                exclusions_normal.append({'CustomerID': cid, 'Reason': 'Sodexo payment (not JPM)', 'EMEA Amount': total_emea})
            elif has_field11_block:
                exclusions_normal.append({'CustomerID': cid, 'Reason': f'Field11 blocked (value: {field11_vals})', 'EMEA Amount': total_emea})
            elif not has_bank or iban_missing:
                exclusions_normal.append({'CustomerID': cid, 'Reason': 'Missing bank details or empty IBAN', 'EMEA Amount': total_emea})
            elif total_emea <= 0:
                exclusions_normal.append({'CustomerID': cid, 'Reason': f'Negative or zero amount ({total_emea})', 'EMEA Amount': total_emea})
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
        'n_anomalies':     len(anomalies),
    }

    # Payments list with IBAN validation
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
            'IBAN':         iban_raw,
            'IBAN Status':  emoji,
            'IBAN Detail':  msg,
        })

    summary['iban_issues'] = iban_issues
    if iban_issues > 0 and status == 'green':
        status = 'yellow'
        summary['status'] = 'yellow'

    buf = _build_report(summary, exclusions_normal, anomalies, payments_list, country_code)
    return status, summary, buf


def _build_report(summary, exclusions_normal, anomalies, payments_list, country_code):
    wb = Workbook()

    GREEN  = 'FF92D050'
    YELLOW = 'FFFFC000'
    RED    = 'FFFF0000'
    HEADER = 'FF4472C4'
    WHITE  = 'FFFFFFFF'
    BLACK  = 'FF000000'

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

    # ── Foglio 1: Summary ─────────────────────────────────
    ws1 = wb.active
    ws1.title = '1. Summary'

    if summary['status'] == 'green':
        status_text = '🟢 OK — No anomalies found'
        status_color = GREEN
        font_color   = WHITE
    elif summary['status'] == 'yellow':
        status_text  = '🟡 WARNING — IBAN issues detected, please check Payments sheet'
        status_color = YELLOW
        font_color   = BLACK
    else:
        status_text  = '🔴 WARNING — Anomalies detected! Please check the report.'
        status_color = RED
        font_color   = WHITE

    ws1.merge_cells('A1:C1')
    cell           = ws1['A1']
    cell.value     = status_text
    cell.font      = Font(bold=True, size=14, color=font_color)
    cell.fill      = PatternFill('solid', fgColor=status_color)
    cell.alignment = center_align
    ws1.row_dimensions[1].height = 30

    data = [
        ('', '', ''),
        ('Country', country_code, ''),
        ('', '', ''),
        ('EMEA Total (all)',        summary['total_emea'],      ''),
        ('Generated file total',    summary['total_generated'],  ''),
        ('Difference',              summary['diff_total'],       '⚠️' if abs(summary['diff_total']) > 0.01 else '✅'),
        ('', '', ''),
        ('Records in EMEA',         summary['n_emea'],          ''),
        ('Records in file',         summary['n_generated'],     ''),
        ('Excluded (known reason)', summary['n_exclusions'],    ''),
        ('Anomalies',               summary['n_anomalies'],     '⚠️' if summary['n_anomalies'] > 0 else '✅'),
        ('IBAN issues',             summary.get('iban_issues', 0), '⚠️' if summary.get('iban_issues', 0) > 0 else '✅'),
    ]

    for row_idx, (label, value, note) in enumerate(data, 2):
        ws1.cell(row=row_idx, column=1, value=label).font = Font(bold=True)
        ws1.cell(row=row_idx, column=2, value=value)
        ws1.cell(row=row_idx, column=3, value=note)

    ws1.column_dimensions['A'].width = 30
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 10

    # ── Foglio 2: Payments ────────────────────────────────
    ws2 = wb.create_sheet('2. Payments')
    write_header(ws2, ['CustomerID', 'Name', 'Amount', 'IBAN', 'IBAN Status', 'IBAN Detail'])

    RED_FILL    = PatternFill('solid', fgColor='FFFFE0E0')
    YELLOW_FILL = PatternFill('solid', fgColor='FFFFFFF0')

    for row_idx, p in enumerate(payments_list, 2):
        ws2.cell(row=row_idx, column=1, value=p['CustomerID'])
        ws2.cell(row=row_idx, column=2, value=p['Name'])
        ws2.cell(row=row_idx, column=3, value=p['Amount'])
        iban_cell        = ws2.cell(row=row_idx, column=4, value=p['IBAN'])
        iban_cell.number_format = '@'
        status_cell      = ws2.cell(row=row_idx, column=5, value=p['IBAN Status'])
        detail_cell      = ws2.cell(row=row_idx, column=6, value=p['IBAN Detail'])

        # Highlight row if IBAN issue
        if p['IBAN Status'] == '❌':
            for col in range(1, 7):
                ws2.cell(row=row_idx, column=col).fill = RED_FILL
        elif p['IBAN Status'] == '⚠️':
            for col in range(1, 7):
                ws2.cell(row=row_idx, column=col).fill = YELLOW_FILL

    last_row = len(payments_list) + 2
    ws2.cell(row=last_row, column=1, value='TOTAL').font = Font(bold=True)
    ws2.cell(row=last_row, column=3, value=round(sum(p['Amount'] for p in payments_list), 2)).font = Font(bold=True)

    ws2.column_dimensions['A'].width = 15
    ws2.column_dimensions['B'].width = 35
    ws2.column_dimensions['C'].width = 15
    ws2.column_dimensions['D'].width = 35
    ws2.column_dimensions['E'].width = 12
    ws2.column_dimensions['F'].width = 45

    # ── Foglio 3: Normal exclusions ───────────────────────
    ws3 = wb.create_sheet('3. Normal exclusions')
    write_header(ws3, ['CustomerID', 'Exclusion reason', 'EMEA Amount'])

    for row_idx, exc in enumerate(exclusions_normal, 2):
        ws3.cell(row=row_idx, column=1, value=exc['CustomerID'])
        ws3.cell(row=row_idx, column=2, value=exc['Reason'])
        ws3.cell(row=row_idx, column=3, value=exc['EMEA Amount'])

    ws3.column_dimensions['A'].width = 15
    ws3.column_dimensions['B'].width = 40
    ws3.column_dimensions['C'].width = 15

    # ── Foglio 4: Anomalies ───────────────────────────────
    ws4 = wb.create_sheet('4. Anomalies')
    write_header(ws4, ['CustomerID', 'Type', 'EMEA Amount', 'Generated Amount', 'Difference', 'Detail'])

    for row_idx, an in enumerate(anomalies, 2):
        ws4.cell(row=row_idx, column=1, value=an['CustomerID'])
        ws4.cell(row=row_idx, column=2, value=an['Type'])
        ws4.cell(row=row_idx, column=3, value=an['EMEA Amount'])
        ws4.cell(row=row_idx, column=4, value=an['Generated Amount'])
        ws4.cell(row=row_idx, column=5, value=an['Difference'])
        ws4.cell(row=row_idx, column=6, value=an['Detail'])

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
