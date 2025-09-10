import pandas as pd

# Read baseline data
baseline = pd.read_excel('POvsBOM_10.09.xlsx', sheet_name='baseline')

print('Checking for duplicates in baseline tab:')
print('\nDuplicate item numbers within same corrosion category:')

# Check for duplicates
duplicates = baseline.groupby(['Item Number', 'Corrosion category']).size()
duplicate_pairs = duplicates[duplicates > 1]

print(f'Found {len(duplicate_pairs)} duplicate combinations:')

if len(duplicate_pairs) > 0:
    for (item_num, corr_cat), count in duplicate_pairs.items():
        print(f'  Item {item_num} in {corr_cat}: appears {count} times')
    
    print('\nDetailed view of duplicates:')
    for (item_num, corr_cat), count in duplicate_pairs.head(3).items():
        subset = baseline[(baseline['Item Number'] == item_num) & (baseline['Corrosion category'] == corr_cat)]
        print(f'\nItem {item_num} - {corr_cat}:')
        print(subset[['Item Number', 'Description', 'Qty PO', 'Remarks', 'Corrosion category']])
else:
    print('No duplicates found in baseline.')

print(f'\nTotal baseline records: {len(baseline)}')
print(f'Unique item-category combinations: {len(baseline.groupby(["Item Number", "Corrosion category"]))}')
