import pandas as pd

# Read the files
baseline = pd.read_excel('POvsBOM_10.09.xlsx', sheet_name='baseline')
result = pd.read_excel('POvsBOM_10.09_processed_2025-09-11_09-59-29.xlsx')

print("=== BASELINE ORDER PRESERVATION TEST ===")

# Create a mapping of baseline positions
baseline['original_order'] = range(len(baseline))
baseline_order = dict(zip(baseline['Item Number'], baseline['original_order']))

# Filter result to exclude new items and special duplicates
# Keep only the first occurrence of each item number
result_filtered = result[result['New'] == False].copy()

# For items that appear multiple times (special items), we need to be careful
# Let's group by item number and keep track
result_grouped = result_filtered.groupby('Item Number').first().reset_index()

print(f"Baseline items: {len(baseline)}")
print(f"Result items (excluding new): {len(result_filtered)}")
print(f"Result unique items: {len(result_grouped)}")

# Check order preservation
order_preserved = 0
total_items = 0
previous_baseline_order = -1

print("\nOrder preservation check (first 30 items):")
for i, row in result_grouped.head(30).iterrows():
    item_num = row['Item Number']
    if item_num in baseline_order:
        baseline_pos = baseline_order[item_num]
        is_in_order = baseline_pos > previous_baseline_order
        order_preserved += 1 if is_in_order else 0
        total_items += 1
        
        status = "✓" if is_in_order else "✗"
        print(f"{total_items:2d}: Item {item_num} - baseline pos {baseline_pos:3d} {status}")
        
        if is_in_order:
            previous_baseline_order = baseline_pos
        elif total_items <= 10:  # Show details for first few out-of-order items
            print(f"    Expected pos > {previous_baseline_order}, got {baseline_pos}")

print(f"\nOrder preservation rate: {order_preserved}/{total_items} = {order_preserved/total_items*100:.1f}%")

# Check the special item 100251
print(f"\n=== SPECIAL ITEMS CHECK ===")
special_items = result[result['Item Number'] == 100251]
print(f"Item 100251 appears {len(special_items)} times:")
for i, (_, row) in enumerate(special_items.iterrows()):
    print(f"  {i+1}: Qty={row['Qty BOM update']}, Remarks='{row['Remarks BOM update']}'")
