from payments.utils import apply_common_filters, save_excel
from openpyxl import Workbook
import io

# Mappatura country code nel file EMEA → nome file output
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

    # Filtri comuni (PayableTy 0/10 + Field11)
    df_c = apply_common_filters(df, country_code)

    # Escludi PayableTy 5 (Sodexo) per BE e NL
    if country_code in SODEXO_EXCLUDE:
        df_c['PayableTy'] = df_c['PayableTy'].where(
            df_c['PayableTy'].notna(), other=None
        )
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

    # Mese in maiuscolo per colonna B (es. FEB COMM, MAR COMM)
    month_upper = month_full[:3].upper()

    # Costruisci righe — 5 colonne
    rows = []
    for _, rec in df_g.iterrows():
        pid = str(rec['partner_id']).strip()
        rows.append([
            pid,
            f"{pid} {month_upper} COMM",
            rec['total_amount'],
            str(rec['iban']).strip(),
            str(rec['deposit_name']).strip(),
        ])

    num_tr    = len(df_g)
    total_eur = round(df_g['total_amount'].sum(), 2)

    # Salva Excel — colonna D (IBAN=4) come testo
    wb = Workbook()
    ws = wb.active
    for row_idx, row in enumerate(rows, 1):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if col_idx == 4 and value not in (None, ''):
                cell.value = str(value)
                cell.number_format = '@'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return buf, num_tr, total_eur, 'EUR'
