from payments.utils import apply_common_filters, save_excel

CFG = {
    'company':              'NSAINTINC',
    'country_code':         'GB',
    'currency':             'GBP',
    'fixed_1':              'GIR',
    'fixed_2':              '01',
    'fixed_3':              '99',
    'jpm_ref':              '67101103',
    'zero_flag':            '0',
    'empty_cols_after_name': 10,
}

def generate(df, payment_date, month_abbrev):
    df_c = apply_common_filters(df, 'GB')

    # Pulisci routing number (rimuovi trattini)
    df_c['DepositRoutingNumber'] = (
        df_c['DepositRoutingNumber'].astype(str)
        .str.replace('-', '', regex=False).str.strip()
    )
    df_c['DepositAccountNumber'] = df_c['DepositAccountNumber'].astype(str).str.strip()

    # Filtro bank details validi
    df_c = df_c[
        df_c['DepositAccountNumber'].notna() &
        df_c['DepositRoutingNumber'].notna() &
        (df_c['DepositAccountNumber'].str.upper() != 'NULL') &
        (df_c['DepositRoutingNumber'].str.upper() != 'NULL') &
        (df_c['DepositAccountNumber'].str.upper() != 'NAN') &
        (df_c['DepositRoutingNumber'].str.upper() != 'NAN') &
        (df_c['DepositAccountNumber'].str.strip('0') != '') &
        (df_c['DepositRoutingNumber'].str.strip('0') != '')
    ]

    # Raggruppa per CustomerID
    df_g = df_c.groupby('effective_id').agg(
        total_amount   = ('Amount',               'sum'),
        deposit_name   = ('DepositName',          'first'),
        account_number = ('DepositAccountNumber', 'first'),
        routing_number = ('DepositRoutingNumber', 'first'),
    ).reset_index()
    df_g.columns = ['partner_id', 'total_amount', 'deposit_name', 'account_number', 'routing_number']
    df_g = df_g[df_g['total_amount'] > 0]
    df_g['amount_pence'] = (df_g['total_amount'] * 100).round().astype(int)

    # Costruisci righe
    rows = [['FH', CFG['company'], payment_date, '101112', '01100']]
    for _, rec in df_g.iterrows():
        pid = str(rec['partner_id']).strip()
        tr = [
            'TR', pid, payment_date, CFG['country_code'], '',
            str(rec['routing_number']).strip(),
            str(rec['account_number']).strip(),
            CFG['zero_flag'], rec['amount_pence'], CFG['currency'],
            CFG['fixed_1'], CFG['fixed_2'], CFG['fixed_3'], CFG['jpm_ref'],
            '', '', str(rec['deposit_name']).strip(),
        ]
        tr += [''] * CFG['empty_cols_after_name']
        tr += [f"{pid}{month_abbrev}Comm"]
        rows.append(tr)

    num_tr      = len(df_g)
    total_pence = int(df_g['amount_pence'].sum())
    rows.append(['FT', num_tr, num_tr + 2, total_pence])

    buf = save_excel(rows, text_cols={6, 7})
    return buf, num_tr, total_pence / 100, 'GBP'
