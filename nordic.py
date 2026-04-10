from payments.utils import apply_common_filters
from openpyxl import Workbook
import io

SWIFT_TO_BANK = {
    'WIOBAEAD':    'WIO BANK P.J.S.C.',
    'WIOBAEADXXX': 'WIO BANK P.J.S.C.',
    'EBILAEAD':    'Emirates NBD Bank',
    'EBILAEADXXX': 'Emirates NBD Bank',
    'BOMLAEAD':    'Mashreq Bank',
    'BBMEAEAD':    'HSBC',
    'BBMEAEADABU': 'HSBC BANK MIDDLE',
    'NRAKAEAK':    'National Bank of Ras Al-Khaimah',
    'ADCBAEAA':    'ABCD Bank',
}

CFG = {
    'header_ref': 20200000000000,
    'col_b': 'WIRES', 'col_c': 'CHASDEFXXXX',
    'col_d': '6161536617', 'col_e': 'N', 'col_f': 'AED',
    'col_n': 'IBAN', 'col_u': 'AE', 'col_ad': 'AE',
}

def generate(df, payment_date, month_full, country_code='AE'):
    df_c = apply_common_filters(df, 'AE')
    df_c['IBAN']      = df_c['IBAN'].astype(str).str.strip().str.upper()
    df_c['SwiftCode'] = df_c['SwiftCode'].astype(str).str.strip().str.upper()
    df_c = df_c[df_c['IBAN'].str.startswith('AE')]

    df_g = df_c.groupby('effective_id').agg(
        total_amount = ('Amount',      'sum'),
        deposit_name = ('DepositName', 'first'),
        iban         = ('IBAN',        'first'),
        swift        = ('SwiftCode',   'first'),
    ).reset_index()
    df_g.columns = ['partner_id', 'total_amount', 'deposit_name', 'iban', 'swift']
    df_g = df_g[df_g['total_amount'] > 0]
    df_g['total_amount'] = df_g['total_amount'].round(2)

    rows = [['HEADER', CFG['header_ref'], 1]]
    for _, rec in df_g.iterrows():
        pid       = str(rec['partner_id']).strip()
        swift     = str(rec['swift']).strip().upper()
        bank_name = SWIFT_TO_BANK.get(swift, '')
        pmt_ref   = f"{pid}{month_full} Comm"
        row = (
            ['P', CFG['col_b'], CFG['col_c'], CFG['col_d'], CFG['col_e'], CFG['col_f']] +
            [rec['total_amount']] +
            [''] * 7 +
            [CFG['col_n']] +
            [str(rec['iban']).strip()] +
            [str(rec['deposit_name']).strip()] +
            [''] * 4 +
            [CFG['col_u']] +
            [''] * 3 +
            [swift, bank_name] +
            [''] * 3 +
            [CFG['col_ad']] +
            [''] * 48 +
            [pmt_ref]
        )
        rows.append(row)

    num_tr    = len(df_g)
    total_aed = round(df_g['total_amount'].sum(), 2)
    rows.append(['TRAILER', num_tr, total_aed])

    wb = Workbook()
    ws = wb.active
    for row_idx, row in enumerate(rows, 1):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if col_idx == 15 and value not in (None, ''):
                cell.value = str(value)
                cell.number_format = '@'
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    gen_ids    = set(df_g['partner_id'].astype(str).str.strip())
    gen_totals = dict(zip(df_g['partner_id'].astype(str).str.strip(), df_g['total_amount'].round(2)))

    return buf, num_tr, total_aed, 'AED', gen_ids, gen_totals
