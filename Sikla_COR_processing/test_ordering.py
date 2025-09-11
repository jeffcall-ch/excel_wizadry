import pandas as pd

# Read the files
baseline = pd.read_excel('POvsBOM_10.09.xlsx', sheet_name='baseline')
result = pd.read_excel('POvsBOM_10.09_processed_2025-09-11_09-56-46.xlsx')

print("=== ORDER VERIFICATION ===")
print(f"Baseline rows: {len(baseline)}")
print(f"Result rows: {len(result)}")

# Filter out new items to compare baseline order
existing_items = result[result['New'] == False]
print(f"Existing items in result: {len(existing_items)}")

# Compare the first 20 item numbers
print("\nFirst 20 baseline vs result item numbers:")
baseline_items = baseline['Item Number'].head(20).tolist()
result_items = existing_items['Item Number'].head(20).tolist()

for i, (b_item, r_item) in enumerate(zip(baseline_items, result_items)):
    match = "✓" if b_item == r_item else "✗"
    print(f"{i+1:2d}: {b_item} -> {r_item} {match}")

# Check if overall sequence is preserved (excluding new items)
baseline_sequence = baseline['Item Number'].tolist()
result_sequence = existing_items['Item Number'].tolist()

# Find where differences occur
differences = 0
for i, (b_item, r_item) in enumerate(zip(baseline_sequence, result_sequence)):
    if b_item != r_item:
        differences += 1
        if differences <= 5:  # Show first 5 differences
            print(f"Difference at position {i+1}: baseline={b_item}, result={r_item}")

print(f"\nTotal differences in sequence: {differences}")
print(f"Order preservation: {((len(baseline_sequence) - differences) / len(baseline_sequence) * 100):.1f}%")

# Check new items positioning
print("\n=== NEW ITEMS ANALYSIS ===")
new_items = result[result['New'] == True]
print(f"New items found: {len(new_items)}")

print("\nNew items and their context:")
for idx in range(min(5, len(new_items))):
    new_item_idx = new_items.index[idx]
    new_item = new_items.iloc[idx]
    
    print(f"\nNew item: {new_item['Item Number']} ({new_item['Item Category']}) at row {new_item_idx}")
    
    # Show context (2 before, 2 after)
    start_idx = max(0, new_item_idx - 2)
    end_idx = min(len(result), new_item_idx + 3)
    context = result.iloc[start_idx:end_idx]
    
    for i, (_, row) in enumerate(context.iterrows()):
        marker = ">>> NEW <<<" if row['New'] else "         "
        print(f"  {row['Item Number']} ({row['Item Category']}) {marker}")
