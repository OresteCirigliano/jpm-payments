from payments.utils import apply_common_filters
from openpyxl import Workbook
import io

NORDIC_CONFIG = {
    'DK': {
        'company':     'TJPCAPSDKK',
        'country':     'DK',
        'currency':    'DKK',
        'jpm_ref':     '550002687',
        'col_m':       '',
        'pmt_ref':     'full_month',
    },
    'SE': {
        'company':     'TJPCAPSSEK',
        'country':     'SE',
        'currency':    'SEK',
        'jpm_ref':     '550002570',
        'col_m':       'NPG',
        'pmt_ref':     'id_only',
    },
    'NO': {
        'company':     'TJPCAPSNOK',
        'country':     'NO',
        'currency':    'NOK',
        'jpm_ref':     '550002679',
        'col_m':       '',
        'pmt_ref':     'full_month',
    },
}

def generate(df, payment_date, month_full, country_code):
    country_code = country_code.upper()
    cfg = NORDIC_CONFIG[country_code]

    df_c = apply_common_filters(df, country_code)

    # Pulisci colonne — G=SwiftCode, H=IBAN (nel file EMEA)
    df_c['SwiftCode'] = df_c['SwiftCode'].astype(str).str.strip()
    df_c['IBAN']      = df_c['IBAN'].astype(str).str.strip()

    # Escludi righe dove colonna H (IBAN) è vuota/null/nan
    df_c = df_c[
        df_c['IBAN'].notna() &
        (df_c['IBAN'].str.upper() != 'NULL') &
        (df_c['IBAN'].str.upper() != 'NAN') &
        (df_c['IBAN'].str.strip('0') != '')
    ]

    # Escludi righe dove colonna G (SwiftCode) è vuota/null/nan
    df_c = df_c[
        df_c['SwiftCode'].notna() &
        (df_c['SwiftCode'].str.upper() != 'NULL') &
        (df_c['SwiftCode'].str.upper() != 'NAN') &
        (df_c['SwiftCode'] != '')
    ]

    # Raggruppa per CustomerID
    df_g = df_c.groupby('effective_id').agg(
        total_amount = ('Amount',      'sum'),
        deposit_name = ('DepositName', 'first'),
        swift        = ('SwiftCode',   'first'),  # G: swift o sort code
        iban_or_acct = ('IBAN',        'first'),  # H: IBAN o bank account
    ).reset_index()
    df_g.columns = ['partner_id', 'total_amount', 'deposit_name', 'swift', 'iban_or_acct']

    df_g = df_g[df_g['total_amount'] > 0]

    # Pulizia post-groupby
    df_g['swift']        = df_g['swift'].astype(str).str.strip()
    df_g['iban_or_acct'] = df_g['iban_or_acct'].astype(str).str.strip()

    # Rimuovi righe con bank details non validi dopo groupby
    df_g = df_g[
        df_g['iban_or_acct'].notna() &
        (df_g['iban_or_acct'].str.upper() != 'NULL') &
        (df_g['iban_or_acct'].str.upper() != 'NAN') &
        (df_g['iban_or_acct'].str.strip('0') != '') &
        df_g['swift'].notna() &
        (df_g['swift'].str.upper() != 'NULL') &
        (df_g['swift'].str.upper() != 'NAN') &
        (df_g['swift'] != '')
    ]

    df_g['amount_cents'] = (df_g['total_amount'] * 100).round().astype(int)

    rows = [['FH', cfg['company'], payment_date, '130000', '01100']]

    for _, rec in df_g.iterrows():
        pid        = str(rec['partner_id']).strip()
        swift      = str(rec['swift']).strip()
        iban_acct  = str(rec['iban_or_acct']).strip()

        # Se SwiftCode inizia con lettera → è uno swift → F vuota, G = IBAN
        # Se SwiftCode inizia con numero → è un sort code → F = sort code, G = bank account
        if swift[0].isalpha():
            col_f = ''
            col_g = iban_acct.upper()
        else:
            col_f = swift
            col_g = iban_acct

        if cfg['pmt_ref'] == 'id_only':
            pmt_ref = pid
        else:
            pmt_ref = f"{pid}{month_full}Comm"

        tr = [
            'TR', pid, payment_date, cfg['country'],
            '',
            col_f,
            col_g,
            '0',
            rec['amount_cents'],
            cfg['currency'],
            'GIR',
            '01',
            cfg['col_m'],
            cfg['jpm_ref'],
            '',
            '',
            str(rec['deposit_name']).strip(),
        ]
        tr += [''] * 10
        tr += [pmt_ref]
        rows.append(tr)

    num_tr      = len(df_g)
    total_cents = int(df_g['amount_cents'].sum())
    rows.append(['FT', num_tr, num_tr + 2, total_cents])

    wb = Workbook()
    ws = wb.active
    for row_idx, row in enumerate(rows, 1):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if col_idx in (6, 7) and value not in (None, ''):
                cell.value = str(value)
                cell.number_format = '@'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return buf, num_tr, total_cents / 100, cfg['currency']
