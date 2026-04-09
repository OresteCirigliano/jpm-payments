from payments.utils import apply_common_filters
from openpyxl import Workbook
import io

# Mappatura code app → code nel file EMEA
COUNTRY_FILTER = {
    'BE':  'BE',
    'EIR': 'EIR',
    'ES':  'ES',
    'FI':  'FI',
    'FR':  'FR',
    'GER': 'DE',   # Nel file EMEA è DE
    'IT':  'IT',
    'LU':  'LU',
    'NL':  'NL',
    'OS':  'AT',   # Nel file EMEA è AT
    'PT':  'PT',
}

EURO_COUNTRIES = {
    'BE':  'Belgium',
    'EIR': 'Ireland',
    'ES':  'Spain',
    'FI':  'Finland',
    'FR':  'France',
    'GER': 'Germany',
    'IT':  'Italy',
    'LU':  'Luxembourg',
    'NL':  'Netherlands',
    'OS':  'Austria',
    'PT':  'Portugal',
}

# Paesi dove PayableTy 5 (Sodexo) va escluso
SODEXO_EXCLUDE = {'BE', 'NL'}

def generate(df, payment_date, month_full, country_code):
    country_code = country_code.upper()

    # Usa il codice corretto per filtrare il file EMEA
    emea_code = COUNTRY_FILTER.get(country_code, country_code)
    df_c = apply_common_filters(df, emea_code)

    # Escludi PayableTy 5 (Sodexo) per BE e NL
    if country_code in SODEXO_EXCLUDE:
        df_c = df_c[df_c['PayableTy'] != 5]

    # Pulisci IBAN
    df_c['IBAN'] = df_c['IBAN'].astype(str).str.strip().str.upper()

    # Escludi righe senza IBAN valido
    df_c = df_c[
        df_c['IBAN'].notna() &
        (df_c['IBAN'].str.upper() != 'NULL') &
        (df_c['IBAN'].str.upper() != 'NAN') &
        (df_c['IBAN'].str.strip('0') != '')
    ]

    # Raggruppa per CustomerID e somma importi
    df_g = df_c.groupby('effective_id').agg(
        total_amount = ('Amount',      'sum'),
        deposit_name = ('DepositName', 'first'),
        iban         = ('IBAN',        'first'),
    ).reset_index()
    df_g.columns = ['partner_id', 'total_amount', 'deposit_name', 'iban']

    # Escludi totali negativi o zero
    df_g = df_g[df_g['total_amount'] > 0]

    # Arrotonda importo a 2 decimali
    df_g['total_amount'] = df_g['total_amount'].round(2)

    # Mese in maiuscolo per colonna B (es. FEB COMM)
    month_upper = month_full[:3].upper()

    # Costruisci righe con colonne nelle posizioni corrette:
    # A=CustomerID, B=ID+mese+COMM, C=vuota, D=Importo,
    # E=vuota, F=vuota, G=IBAN, H=vuota, I=Nome
    rows = []
    for _, rec in df_g.iterrows():
        pid = str(rec['partner_id']).strip()
        rows.append([
            pid,                              # A - CustomerID
            f"{pid} {month_upper} COMM",      # B - riferimento
            '',                               # C - vuota
            rec['total_amount'],              # D - importo
            '',                               # E - vuota
            '',                               # F - vuota
            str(rec['iban']).strip(),         # G - IBAN
            '',                               # H - vuota
            str(rec['deposit_name']).strip(), # I - Nome
        ])

    num_tr    = len(df_g)
    total_eur = round(df_g['total_amount'].sum(), 2)

    # Salva Excel — colonna G (IBAN=7) come testo
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

    return buf, num_tr, total_eur, 'EUR'
