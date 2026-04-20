from payments.utils import apply_common_filters, clean_name
from openpyxl import Workbook
import io

NORDIC_CONFIG = {
    'DK': {'company': 'TJPCAPSDKK', 'country': 'DK', 'currency': 'DKK', 'jpm_ref': '550002687', 'col_m': '',    'pmt_ref': 'full_month'},
    'SE': {'company': 'TJPCAPSSEK', 'country': 'SE', 'currency': 'SEK', 'jpm_ref': '550002570', 'col_m': 'NPG', 'pmt_ref': 'id_only'},
    'NO': {'company': 'TJPCAPSNOK', 'country': 'NO', 'currency': 'NOK', 'jpm_ref': '550002679', 'col_m': '',    'pmt_ref': 'full_month'},
}

def is_empty(val):
    return str(val).strip().upper() in ('', 'NULL', 'NAN', 'NONE')

def generate(df, payment_date, month_full, country_code):
    country_code = country_code.upper()
    cfg = NORDIC_CONFIG[country_code]

    df_c = apply_common_filters(df, country_code)
    df_c['SwiftCode'] = df_c['SwiftCode'].astype(str).str.strip()
    df_c['IBAN']      = df_c['IBAN'].astype(str).str.strip()

    df_c = df_c[~df_c['IBAN'].apply(is_empty)]
    has_valid_swift = ~df_c['SwiftCode'].apply(is_empty)
    has_direct_iban = df_c['IBAN'].str.upper().str.startswith(country_code)
    df_c = df_c[has_valid_swift | has_direct_iban]

    df_g = df_c.groupby('effective_id').agg(
        total_amount = ('Amount',      'sum'),
        deposit_name = ('DepositName', 'first'),
        swift        = ('SwiftCode',   'first'),
        iban_or_acct = ('IBAN',        'first'),
    ).reset_index()
    df_g.columns = ['partner_id', 'total_amount', 'deposit_name', 'swift', 'iban_or_acct']

    df_g = df_g[df_g['total_amount'] > 0.01]
    df_g['swift']        = df_g['swift'].astype(str).str.strip()
    df_g['iban_or_acct'] = df_g['iban_or_acct'].astype(str).str.strip()
    df_g = df_g[~df_g['iban_or_acct'].apply(is_empty)]
    has_valid_swift_g = ~df_g['swift'].apply(is_empty)
    has_direct_iban_g = df_g['iban_or_acct'].str.upper().str.startswith(country_code)
    df_g = df_g[has_valid_swift_g | has_direct_iban_g]

    df_g['amount_cents'] = (df_g['total_amount'] * 100).round().astype(int)

    rows = [['FH', cfg['company'], payment_date, '130000', '01100']]

    for _, rec in df_g.iterrows():
        pid       = str(rec['partner_id']).strip()
        swift     = str(rec['swift']).strip()
        iban_acct = str(rec['iban_or_acct']).strip()

        # If value starts with country prefix → it's an IBAN → F empty, G = IBAN
        # If value starts with a digit → it's a bank account → F = sort code, G = account
        # Otherwise → F empty, G = value uppercase
        if iban_acct.upper().startswith(country_code):
            col_f = ''
            col_g = iban_acct.upper()
        elif iban_acct[0].isdigit() and not is_empty(swift):
            col_f = swift
            col_g = iban_acct
        else:
            col_f = ''
            col_g = iban_acct.upper()

        pmt_ref = pid if cfg['pmt_ref'] == 'id_only' else f"{pid}{month_full}Comm"

        # Clean name to remove special characters (e.g. ü → ue)
        name = clean_name(str(rec['deposit_name']))

        tr = [
            'TR', pid, payment_date, cfg['country'], '',
            col_f, col_g, '0',
            rec['amount_cents'], cfg['currency'],
            'GIR', '01', cfg['col_m'], cfg['jpm_ref'], '', '',
            name,
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

    gen_ids    = set(df_g['partner_id'].astype(str).str.strip())
    gen_totals = dict(zip(df_g['partner_id'].astype(str).str.strip(), df_g['total_amount'].round(2)))

    return buf, num_tr, total_cents / 100, cfg['currency'], gen_ids, gen_totals
