import pandas as pd
from openpyxl import Workbook
import io

PAYABLE_EXCLUDE = [0, 10]

COLUMN_NAMES = {
    0:  "CustomerID",
    1:  "Field2",
    2:  "NAME",
    3:  "DepositName",
    4:  "Email",
    5:  "Country",
    6:  "SwiftCode",
    7:  "IBAN",
    8:  "DepositAccountNumber",
    9:  "DepositRoutingNumber",
    10: "Amount",
    11: "Field11",
    12: "PayableTy",
    13: "VendorBillID",
}

SPECIAL_CHARS = {
    'ä': 'ae', 'ö': 'oe', 'ü': 'ue',
    'Ä': 'Ae', 'Ö': 'Oe', 'Ü': 'Ue',
    'ß': 'ss',
    'à': 'a', 'á': 'a', 'â': 'a', 'ã': 'a',
    'è': 'e', 'é': 'e', 'ê': 'e', 'ë': 'e',
    'ì': 'i', 'í': 'i', 'î': 'i', 'ï': 'i',
    'ò': 'o', 'ó': 'o', 'ô': 'o', 'õ': 'o',
    'ù': 'u', 'ú': 'u', 'û': 'u',
    'ý': 'y', 'ÿ': 'y',
    'ñ': 'n', 'ç': 'c',
    'À': 'A', 'Á': 'A', 'Â': 'A', 'Ã': 'A',
    'È': 'E', 'É': 'E', 'Ê': 'E', 'Ë': 'E',
    'Ì': 'I', 'Í': 'I', 'Î': 'I', 'Ï': 'I',
    'Ò': 'O', 'Ó': 'O', 'Ô': 'O', 'Õ': 'O',
    'Ù': 'U', 'Ú': 'U', 'Û': 'U',
    'Ý': 'Y', 'Ñ': 'N', 'Ç': 'C',
}

MONTH_ABBREV = {
    'January': 'Jan', 'February': 'Feb', 'March': 'Mar',
    'April': 'Apr', 'May': 'May', 'June': 'Jun',
    'July': 'Jul', 'August': 'Aug', 'September': 'Sep',
    'October': 'Oct', 'November': 'Nov', 'December': 'Dec',
}

def clean_name(name):
    if not isinstance(name, str):
        return name
    for char, replacement in SPECIAL_CHARS.items():
        name = name.replace(char, replacement)
    return name.encode('ascii', 'ignore').decode('ascii').strip()

def read_and_rename(file_bytes):
    df = pd.read_excel(
        io.BytesIO(file_bytes),
        dtype={0: str, 1: str, 7: str, 8: str, 9: str}
    )
    for pos, name in COLUMN_NAMES.items():
        if pos < len(df.columns):
            df.rename(columns={df.columns[pos]: name}, inplace=True)
    df['effective_id'] = df['CustomerID'].astype(str).str.strip()
    return df

def apply_common_filters(df, country_code):
    df_c = df[df['Country'].str.strip().str.upper() == country_code.upper()].copy()
    df_c['PayableTy'] = pd.to_numeric(df_c['PayableTy'], errors='coerce')
    df_c = df_c[~df_c['PayableTy'].isin(PAYABLE_EXCLUDE)]
    df_c['Field11'] = pd.to_numeric(df_c['Field11'], errors='coerce')
    df_c = df_c[df_c['Field11'].isna() | (df_c['Field11'] == 3)]
    return df_c

def save_excel(rows, text_cols):
    """Salva le righe in un file Excel e ritorna un buffer BytesIO."""
    wb = Workbook()
    ws = wb.active
    for row_idx, row in enumerate(rows, 1):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if col_idx in text_cols and value not in (None, ''):
                cell.value = str(value)
                cell.number_format = '@'
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
