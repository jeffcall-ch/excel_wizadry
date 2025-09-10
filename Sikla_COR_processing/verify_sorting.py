import pandas as pd

df = pd.read_excel('POvsBOM_10.09_processed_2025-09-10_11-38-01.xlsx')

print('Verification of sorting and new item placement:')
print('\nFirst 10 rows showing sorting by Corrosion category, Item Category:')
print(df[['Corrosion category', 'Item Category', 'Item Number', 'New', 'Deleted']].head(10))

print('\nNew items grouped by category:')
new_items = df[df['New'] == True]
print('New items by category:')
for cat in new_items['Item Category'].unique():
    cat_items = new_items[new_items['Item Category'] == cat]
    print(f'{cat}: {len(cat_items)} items')

print('\nSample new items:')
print(new_items[['Corrosion category', 'Item Category', 'Item Number', 'Description']].head(5))

print('\nVerifying items are grouped properly:')
print('\nC4 Primary items (should be together):')
c4_primary = df[(df['Corrosion category'] == 'C4') & (df['Item Category'] == 'Primary')]
print(f'Count: {len(c4_primary)}')

print('\nC5L items (should be together):')
c5l_items = df[df['Corrosion category'] == 'C5L RAL7035']
print(f'Count: {len(c5l_items)}')
print(c5l_items[['Item Category', 'Item Number', 'New']].head(5))
