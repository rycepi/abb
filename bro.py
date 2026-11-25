import pandas as pd


def _load_sheet(name: str) -> pd.DataFrame:
    df = pd.read_excel(f'newStuffs/New {name}.xlsx', sheet_name='Sheet1')
    df.columns = [str(col).strip() for col in df.columns]
    return df


df = pd.read_csv('vendorlist.txt', sep='\t', header=None)
sheet_names = ['Calls', 'Emails', 'Gift Cards', 'Shipments', 'Storage']

# dataframes keyed by sheet name (not a list)
dataframes = {name: _load_sheet(name) for name in sheet_names}

# initialize per-vendor targets with correct columns
companies = {
    vendor: {name: pd.DataFrame(columns=dataframes[name].columns) for name in sheet_names}
    for vendor in df[0]
}



for name, sdf in dataframes.items():
    for _, row in sdf.iterrows():
        vendor = row['Vendor Name']
        if vendor not in companies:
            continue
        target = companies[vendor][name]
        companies[vendor][name] = pd.concat([target, row.to_frame().T], ignore_index=True)

for company, sheets in companies.items():
    with pd.ExcelWriter(f'finish/{company}.xlsx') as writer:
        for name, sdf in sheets.items():
            sdf.to_excel(writer, sheet_name=name, index=False)