from payments.utils import apply_common_filters
from openpyxl import Workbook
import io

NORDIC_CONFIG = {
    'DK': {
        'company':     'TJPCAPSDKK',
        'country':     'DK',
        'currency':    'DKK',
        'jpm_ref':     '550002687',
        'iban_prefix': 'DK',
        'col_m':       '',
        'pmt_ref':     'full_month',
    },
    'SE': {
        'company':     'TJPCAPSSEK',
        'country':     'SE',
        'currency':    'SEK',
        'jpm_ref':     '550002570',
        'iban_prefix': 'SE',
        'col_m':       'NPG',
        'pmt_ref':     'id_only',
    },
    'NO': {
        'company':     'TJPCAPSNOK',
        'country':     'NO',
        'currency':    'NOK',
        'jpm_ref':     '550002679',
        'iban_prefix': 'NO',
        'col_m':       '',
        'pmt_ref':     'full_month',
    },
}

def generate(df, payment_date, month_full, country_code):
    country_code = country_code.upper()
    cfg = NORDIC_CONFIG[country_code]

    df_c = apply_common_filters(df, country_code)

    # Pulisci colonne
    df_c['IBAN'] = df_c['IBAN'].astype(str).str.strip()
    df_c['DepositRoutingNumber'] = df_c['DepositRoutingNumber'].astype(str).str.strip()
    df_c['DepositAccountNumber'] = df_c['DepositAccountNumber'].astype(str).str.strip()

    # Escludi righe senza bank details validi
    has_iban = df_c['IBAN'].str.upper().str.startswith(cfg['iban_prefix'])
    has_account = (
        (df_c['DepositAccountNumber'].str.upper() != 'NULL') &
        (df_c['DepositAccountNumber'].str.upper() != 'NAN') &
        (df_c['DepositAccountNumber'].str.strip('0') != '')
    )
    df_c = df_c[has_iban | has_account]

    # Raggruppa per CustomerID
    df_g = df_c.groupby('effective_id').agg(
        total_amount   = ('Amount',               'sum'),
        deposit_name   = ('DepositName',          'first'),
        iban           = ('IBAN',                 'first'),
        routing_number = ('DepositRoutingNumber', 'first'),
        account_number = ('DepositAccountNumber', 'first'),
    ).reset_index()
    df_g.columns = ['partner_id', 'total_amount', 'deposit_name',
                    'iban', 'routing_number', 'account_number']

    # Escludi totali negativi o zero
    df_g = df_g[df_g['total_amount'] > 0]

    # Pulisci e rimuovi righe senza bank details validi dopo il groupby
    df_g['iban'] = df_g['iban'].astype(str).str.strip()
    df_g['account_number'] = df_g['account_number'].astype(str).str.strip()
    df_g['routing_number'] = df_g['routing_number'].astype(str).str.strip()

    has_iban_g = df_g['iban'].str.upper().str.startswith(cfg['iban_prefix'])
    has_account_g = (
        (df_g['account_number'].str.upper() != 'NULL') &
        (df_g['account_number'].str.upper() != 'NAN') &
        (df_g['account_number'].str.strip('0') != '')
    )
    df_g = df_g[has_iban_g | has_account_g]

    df_g['amount_cents'] = (df_g['total_amount'] * 100).round().astype(int)

    rows = [['FH', cfg['company'], payment_date, '130000', '01100']]

    for _, rec in df_g.iterrows():
        pid  = str(rec['partner_id']).strip()
        iban = str(rec['iban']).strip()

        if iban.upper().startswith(cfg['iban_prefix']):
            col_f = ''
            col_g = iban.upper()
        else:
            col_f = str(rec['routing_number']).strip()
            col_g = str(rec['account_number']).strip()

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
