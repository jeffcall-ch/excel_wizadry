import pandas as pd

df = pd.read_csv('C4_BOM_v10.0_aggregated_2025-09-10_09-39-50.csv')

print('Results after extra-long aggregation:')
print('\nNo more duplicates should exist:')
print('Duplicate items remaining:', df['Duplicate Item Number'].sum())

print('\nExample items that were aggregated (look for higher quantities):')
examples = [110211, 110200, 110202]

for item in examples:
    item_rows = df[df['Item Number'] == item]
    if len(item_rows) > 0:
        print(f'\nItem {item}:')
        for _, row in item_rows.iterrows():
            print(f'  Qty: {row["Qty"]}, Remarks: {row["Remarks"]}, Category: {row["Corrosion Category"]}')

print('\nItems with extra-long remarks:')
extra_long = df[df['Remarks'].str.contains('extra-long', na=False)]
print(f'Count: {len(extra_long)}')
print(extra_long[['Item Number', 'Qty', 'Remarks', 'Corrosion Category']].head(5))
