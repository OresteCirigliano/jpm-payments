from payments.utils import apply_common_filters
from openpyxl import Workbook
import io

CFG = {
    'company':               'NSAINTINC',
    'country':               'PL',
    'currency':              'PLN',
    'jpm_ref':               '550003388',
    'empty_cols_after_name': 10,
}

def generate(df, payment_date, month_full, country_code='PL'):
    df_c = apply_common_filters(df, 'PL')
    df_c['IBAN'] = df_c['IBAN'].astype(str).str.strip().str.upper()
    df_c = df_c[df_c['IBAN'].str.startswith('PL')]

    df_g = df_c.groupby('effective_id').agg(
        total_amount = ('Amount',      'sum'),
        deposit_name = ('DepositName', 'first'),
        iban         = ('IBAN',        'first'),
    ).reset_index()
    df_g.columns = ['partner_id', 'total_amount', 'deposit_name', 'iban']
    df_g = df_g[df_g['total_amount'] > 0]
    df_g['amount_cents'] = (df_g['total_amount'] * 100).round().astype(int)

    rows = [['FH', CFG['company'], payment_date, '101112', '01100']]
    for _, rec in df_g.iterrows():
        pid = str(rec['partner_id']).strip()
        tr = ['TR', pid, payment_date, CFG['country'], '', '', str(rec['iban']).strip(), '',
              rec['amount_cents'], CFG['currency'], 'GIR', '01', '', CFG['jpm_ref'], '', '',
              str(rec['deposit_name']).strip()]
        tr += [''] * CFG['empty_cols_after_name']
        tr += [f"{pid}{month_full}Comm"]
        rows.append(tr)

    num_tr      = len(df_g)
    total_cents = int(df_g['amount_cents'].sum())
    rows.append(['FT', num_tr, num_tr + 2, total_cents])

    wb = Workbook()
    ws = wb.active
    for row_idx, row in enumerate(rows, 1):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if col_idx == 7 and value not in (None, ''):
                cell.value = str(value)
                cell.number_format = '@'
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    gen_ids    = set(df_g['partner_id'].astype(str).str.strip())
    gen_totals = dict(zip(df_g['partner_id'].astype(str).str.strip(), df_g['total_amount'].round(2)))

    return buf, num_tr, total_cents / 100, CFG['currency'], gen_ids, gen_totals
