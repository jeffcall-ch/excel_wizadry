import pandas as pd

df = pd.read_csv('C4_BOM_v10.0_aggregated_2025-09-10_09-27-02.csv')

print('Sample of items in different corrosion categories:')
example_items = [110211, 193860]

for item in example_items:
    item_rows = df[df['Item Number'] == item]
    if len(item_rows) > 1:
        print(f'\nItem Number {item}:')
        for _, row in item_rows.iterrows():
            print(f'  {row["Description"]} - {row["Corrosion Category"]} - Total Weight: {row["Total Weight [kg]"]}')
            if pd.notna(row['Total Length [mm]']):
                print(f'    Total Length: {row["Total Length [mm]"]}mm')

print(f'\nCorrosion categories in output:')
print(df['Corrosion Category'].value_counts())

print(f'\nSample of output columns:')
print(df.columns.tolist())

print(f'\nFirst 5 rows:')
print(df[['Item Number', 'Description', 'Corrosion Category', 'Duplicate Item Number']].head())
