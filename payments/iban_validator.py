# IBAN lengths by country code
IBAN_LENGTHS = {
    'AL': 28, 'AD': 24, 'AT': 20, 'AZ': 28, 'BH': 22, 'BE': 16, 'BA': 20,
    'BR': 29, 'BG': 22, 'CR': 22, 'HR': 21, 'CY': 28, 'CZ': 24, 'DK': 18,
    'DO': 28, 'EE': 20, 'FO': 18, 'FI': 18, 'FR': 27, 'GE': 22, 'DE': 22,
    'GI': 23, 'GR': 27, 'GL': 18, 'GT': 28, 'HU': 28, 'IS': 26, 'IE': 22,
    'IL': 23, 'IT': 27, 'JO': 30, 'KZ': 20, 'KW': 30, 'LV': 21, 'LB': 28,
    'LI': 21, 'LT': 20, 'LU': 20, 'MK': 19, 'MT': 31, 'MR': 27, 'MU': 30,
    'MC': 27, 'MD': 24, 'ME': 22, 'NL': 18, 'NO': 15, 'PK': 24, 'PS': 29,
    'PL': 28, 'PT': 25, 'QA': 29, 'RO': 24, 'SM': 27, 'SA': 24, 'RS': 22,
    'SK': 24, 'SI': 19, 'ES': 24, 'SE': 24, 'CH': 21, 'TN': 24, 'TR': 26,
    'AE': 23, 'GB': 22, 'VG': 24,
}

# Countries where IBAN prefix must match payment country
STRICT_PREFIX = {
    'GB':  ['GB'],
    'CH':  ['CH', 'LI'],
    'DK':  ['DK'],
    'SE':  ['SE'],
    'NO':  ['NO'],
    'PL':  ['PL'],
    'AE':  ['AE'],
}

# EURO countries — any valid IBAN accepted
EURO_COUNTRIES = {'BE', 'EIR', 'ES', 'FI', 'FR', 'GER', 'IT', 'LU', 'NL', 'OS', 'PT'}


def _mod97(iban_clean):
    """Run MOD97 check on a cleaned IBAN (no spaces, uppercase)."""
    rearranged = iban_clean[4:] + iban_clean[:4]
    numeric = ''.join(str(ord(c) - 55) if c.isalpha() else c for c in rearranged)
    return int(numeric) % 97 == 1


def validate_iban(iban, country_code):
    """
    Validate an IBAN for a given payment country.
    Returns (is_valid, status_emoji, message)
    """
    if not iban or str(iban).strip().upper() in ('', 'NULL', 'NAN', 'NONE'):
        return False, '⚠️', 'Missing IBAN'

    iban_clean = str(iban).strip().upper().replace(' ', '').replace('-', '')
    prefix     = iban_clean[:2]

    # Check prefix is letters
    if not prefix.isalpha():
        return False, '❌', f'Invalid IBAN format (does not start with country code)'

    # Check known length
    expected_len = IBAN_LENGTHS.get(prefix)
    if expected_len is None:
        return False, '⚠️', f'Unknown country prefix: {prefix}'

    if len(iban_clean) != expected_len:
        return False, '❌', f'Invalid length: {len(iban_clean)} chars (expected {expected_len} for {prefix})'

    # MOD97 check
    if not _mod97(iban_clean):
        return False, '❌', 'Invalid check digit (MOD97 failed)'

    # Prefix check for strict countries
    allowed_prefixes = STRICT_PREFIX.get(country_code)
    if allowed_prefixes and prefix not in allowed_prefixes:
        return False, '⚠️', f'IBAN prefix {prefix} unexpected for {country_code} (expected: {", ".join(allowed_prefixes)})'

    return True, '✅', 'Valid'
