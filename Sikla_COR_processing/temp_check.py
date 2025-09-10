import pandas as pd

df = pd.read_excel('C4_BOM_v10.0.xlsx')

# Find items that appear in both categories
c4_items = set(df[df['Corrosion Category'] == 'C4']['Item Number'])
c5l_items = set(df[df['Corrosion Category'] == 'C5L RAL7035']['Item Number'])
common_items = list(c4_items.intersection(c5l_items))[:3]

print('Examples of items in both categories:')
for item in common_items:
    subset = df[df['Item Number'] == item]
    print(f'Item {item}:')
    for _, row in subset.iterrows():
        print(f'  {row["Description"]} - {row["Corrosion Category"]}')
    print()
