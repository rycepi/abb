import pandas as pd


def _normalize_value(value):
    if pd.isna(value):
        return ''
    return str(value).strip().lower()


def _find_column(df: pd.DataFrame, target_name: str):
    target_norm = target_name.strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == target_norm:
            return col
    return None


august = pd.read_excel('Summary.xlsx', sheet_name=0)

calls = pd.read_excel('newStuffs/New Calls.xlsx', sheet_name=0)
emails = pd.read_excel('newStuffs/New Emails.xlsx', sheet_name=0)
giftcards = pd.read_excel('newStuffs/New Gift Cards.xlsx', sheet_name=0)
shipments = pd.read_excel('newStuffs/New Shipments.xlsx', sheet_name=0)
storage = pd.read_excel('newStuffs/New Storage.xlsx', sheet_name=0)

# Build normalized vendor key on every sheet (strip + case-insensitive)
for df in [august, calls, emails, giftcards, shipments, storage]:
    vcol = _find_column(df, 'Vendor Name')
    if vcol is not None:
        df['_vendor_key'] = df[vcol].apply(_normalize_value)
    else:
        df['_vendor_key'] = ''

# Case Count: prefer the "number" column if it exists (case/space-insensitive)
number_col = _find_column(calls, 'number')
if number_col is not None:
    calls_case_counts = calls.groupby('_vendor_key')[number_col].count()
else:
    calls_case_counts = calls.groupby('_vendor_key').size()

# Email counts per vendor
email_counts = emails.groupby('_vendor_key').size()

# Gift cards counts and amounts (robust to column name case/spacing)
amount_col = _find_column(giftcards, 'Amount')
if amount_col is not None:
    giftcards['_amount'] = pd.to_numeric(giftcards[amount_col], errors='coerce').fillna(0)
else:
    giftcards['_amount'] = 0
gc_counts = giftcards.groupby('_vendor_key').size()
gc_amounts = giftcards.groupby('_vendor_key')['_amount'].sum()


def tier_sum_quantity(tier_label: str) -> pd.Series:
    # Sum quantities for rows matching the tier label per vendor
    qty_col = _find_column(shipments, 'Quantity Fulfilled/Received')
    if qty_col is not None:
        quantities = pd.to_numeric(shipments[qty_col], errors='coerce').fillna(0)
    else:
        quantities = pd.Series(0, index=shipments.index, dtype='float64')

    formula_col = _find_column(shipments, 'Formula (Text)')
    if formula_col is not None:
        mask = shipments[formula_col].astype(str).str.contains(tier_label, case=False, na=False)
    else:
        mask = pd.Series(False, index=shipments.index)

    df = shipments.loc[mask, ['_vendor_key']].copy()
    df['_qty'] = quantities[mask].values
    return df.groupby('_vendor_key')['_qty'].sum()


tier1_counts = tier_sum_quantity('Tier 1')
tier2_counts = tier_sum_quantity('Tier 2')
tier3_counts = tier_sum_quantity('Tier 3')

# Fulfillment count: count unique transactions per vendor (no duplicates)
trans_col = _find_column(shipments, 'Fulfilling/Receiving Transaction')
if trans_col is not None:
    fulfillment_counts = (
        shipments.dropna(subset=[trans_col])
        .groupby('_vendor_key')[trans_col]
        .nunique()
    )
else:
    fulfillment_counts = pd.Series(dtype=int)

# Bin %: each vendor's share of all rows in Storage.
total_bins = len(storage)
bins_by_vendor = storage.loc[storage['_vendor_key'].ne('')].groupby('_vendor_key').size()
bin_percentages = (
    august['_vendor_key'].map(bins_by_vendor).fillna(0) * 100 / total_bins
    if total_bins else 0.0
)

# Replace the old template columns in place, or add Bin % if missing.
bin_columns = [
    col for col in august.columns
    if ''.join(str(col).lower().split()) in {'bin%', 'bigbin%', 'smallbin%'}
]
bin_position = august.columns.get_loc(bin_columns[0]) if bin_columns else len(august.columns)
august = august.drop(columns=bin_columns)
august.insert(bin_position, 'Bin %', bin_percentages)

# Lookups use normalized vendor key from august
august_vendor_col = _find_column(august, 'Vendor Name')

for idx, row in august.iterrows():
    vendor_key = _normalize_value(row[august_vendor_col]) if august_vendor_col is not None else ''

    case_count = int(calls_case_counts.get(vendor_key, 0))
    email_count = int(email_counts.get(vendor_key, 0))

    gc_count = int(gc_counts.get(vendor_key, 0))
    gc_amount = float(gc_amounts.get(vendor_key, 0.0))

    t1 = float(tier1_counts.get(vendor_key, 0))
    t2 = float(tier2_counts.get(vendor_key, 0))
    t3 = float(tier3_counts.get(vendor_key, 0))

    fulfill_count = int(fulfillment_counts.get(vendor_key, 0))

    august.at[idx, 'Case Count'] = case_count
    august.at[idx, 'Email Count'] = email_count
    august.at[idx, 'Tier 1 Picks Count'] = t1
    august.at[idx, 'Tier 2 Picks Count'] = t2
    august.at[idx, 'Tier 3 Picks Count'] = t3
    august.at[idx, 'Fulfillment Count'] = fulfill_count
    august.at[idx, 'Gift Cards Count'] = gc_count
    august.at[idx, 'Gift Cards Amount'] = round(gc_amount, 2)

with pd.ExcelWriter('Summary.xlsx', engine='openpyxl', mode='w') as writer:
    export_august = august.drop(columns=['_vendor_key'], errors='ignore')
    export_august.to_excel(writer, sheet_name='Sheet1', index=False)
