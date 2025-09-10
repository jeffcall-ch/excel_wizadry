import pandas as pd

df = pd.read_excel('C4_BOM_v10.0.xlsx')

# Find items with cut length that appear in multiple corrosion categories
cut_length_items = df[df['Cut Length [mm]'].notna()]
print("Examples of items WITH cut length to explain grouping:")

# Example 1: Items with same item number but different corrosion categories
item_114842 = cut_length_items[cut_length_items['Item Number'] == 114842]
if len(item_114842) > 0:
    print(f"\nItem Number 114842:")
    for _, row in item_114842.head(3).iterrows():
        print(f"  Description: {row['Description']}")
        print(f"  Cut Length: {row['Cut Length [mm]']}")
        print(f"  Corrosion Category: {row['Corrosion Category']}")
        print()

# Check if any cut length items appear in both categories
c4_cut_items = set(cut_length_items[cut_length_items['Corrosion Category'] == 'C4']['Item Number'])
c5l_cut_items = set(cut_length_items[cut_length_items['Corrosion Category'] == 'C5L RAL7035']['Item Number'])
common_cut_items = c4_cut_items.intersection(c5l_cut_items)

print(f"Cut length items appearing in both categories: {len(common_cut_items)}")
if common_cut_items:
    example_item = list(common_cut_items)[0]
    print(f"\nExample - Item Number {example_item}:")
    example_data = cut_length_items[cut_length_items['Item Number'] == example_item]
    for _, row in example_data.iterrows():
        print(f"  Description: {row['Description']}")
        print(f"  Cut Length: {row['Cut Length [mm]']}")
        print(f"  Corrosion Category: {row['Corrosion Category']}")
        print()
