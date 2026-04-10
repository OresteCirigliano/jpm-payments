from payments.utils import apply_common_filters, clean_name, save_excel

CFG = {
    'company':               'NSAINTINC',
    'country_code':          'CH',
    'currency':              'CHF',
    'fixed_1':               'GIR',
    'fixed_2':               '01',
    'jpm_ref':               '8770000966',
    'empty_cols_after_name': 9,
}

def generate(df, payment_date, month_full):
    df_c = apply_common_filters(df, 'CH')
    df_c['IBAN'] = df_c['IBAN'].astype(str).str.strip().str.upper()
    df_c = df_c[
        df_c['IBAN'].notna() &
        (df_c['IBAN'].str.upper() != 'NULL') &
        (df_c['IBAN'].str.upper() != 'NAN') &
        (df_c['IBAN'].str.strip('0') != '') &
        (df_c['IBAN'].str.upper().str.startswith(('CH', 'LI')))
    ]
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
        pid  = str(rec['partner_id']).strip()
        name = clean_name(rec['deposit_name'])
        tr = ['TR', pid, payment_date, CFG['country_code'], '', '',
              str(rec['iban']).strip(), '', rec['amount_cents'], CFG['currency'],
              CFG['fixed_1'], CFG['fixed_2'], '', CFG['jpm_ref'], '', '', name]
        tr += [''] * CFG['empty_cols_after_name']
        tr += [f"{pid}{month_full}Comm"]
        rows.append(tr)

    num_tr      = len(df_g)
    total_cents = int(df_g['amount_cents'].sum())
    rows.append(['FT', num_tr, num_tr + 2, total_cents])

    buf = save_excel(rows, text_cols={7})

    gen_ids    = set(df_g['partner_id'].astype(str).str.strip())
    gen_totals = dict(zip(df_g['partner_id'].astype(str).str.strip(), df_g['total_amount'].round(2)))

    return buf, num_tr, total_cents / 100, 'CHF', gen_ids, gen_totals
