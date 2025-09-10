import pandas as pd
import numpy as np
import re
from datetime import datetime
import os

def clean_numeric_value(value, decimal_places=2):
    """Extract float value from text and return rounded to specified decimal places"""
    if pd.isna(value):
        return np.nan
    
    if isinstance(value, (int, float)):
        return round(float(value), decimal_places)
    
    # If it's a string, extract numbers
    if isinstance(value, str):
        # Remove any non-numeric characters except decimal point and minus
        cleaned = re.sub(r'[^\d.-]', '', str(value))
        try:
            return round(float(cleaned), decimal_places)
        except ValueError:
            return np.nan
    
    return np.nan

def process_description_with_cut_length(description):
    """
    Process description for items with cut length.
    Find ' x ' (space before and after) followed by a number and remove it along with the number.
    Keep any text that follows after the number.
    Exclude 'Fabric Tape' items from this processing.
    """
    if pd.isna(description) or not isinstance(description, str):
        return description
    
    # Exclude Fabric Tape items
    if 'Fabric Tape' in description:
        return description
    
    # Find pattern: ' x ' followed by a number
    pattern = r' x (\d+)'
    match = re.search(pattern, description)
    
    if match:
        # Remove ' x [number]' but keep any text after the number
        start_pos = match.start()
        end_pos = match.end()
        
        # Get the part before ' x ' and the part after the number
        before_x = description[:start_pos]
        after_number = description[end_pos:]
        
        # Combine them
        processed_description = before_x + after_number
        return processed_description.strip()
    
    return description

def aggregate_data(df):
    """
    Main aggregation function that processes the dataframe according to the rules
    """
    # Separate items with and without remarks (items with remarks are kept separate)
    items_with_remarks = df[df['Remarks'].notna()].copy()
    items_without_remarks = df[df['Remarks'].isna()].copy()
    
    # Separate items without remarks into those with and without cut length
    items_no_cut_length = items_without_remarks[items_without_remarks['Cut Length [mm]'].isna()].copy()
    items_with_cut_length = items_without_remarks[items_without_remarks['Cut Length [mm]'].notna()].copy()
    
    aggregated_results = []
    
    # Process items WITHOUT cut length (aggregate by item number)
    if not items_no_cut_length.empty:
        grouped_no_cut = items_no_cut_length.groupby('Item Number').agg({
            'Qty': 'sum',
            'Description': 'first',  # Should be the same for same item number
            'Total Weight [kg]': 'sum'
        }).reset_index()
        
        # Add empty columns for consistency (but no Cut Length for aggregated items)
        grouped_no_cut['Total Length [mm]'] = np.nan
        grouped_no_cut['Item Type'] = 'Aggregated'
        
        aggregated_results.append(grouped_no_cut)
    
    # Process items WITH cut length
    if not items_with_cut_length.empty:
        # Calculate Total Length [mm] = Qty * Cut Length [mm]
        items_with_cut_length['Total Length [mm]'] = items_with_cut_length['Qty'] * items_with_cut_length['Cut Length [mm]']
        
        # Process descriptions
        items_with_cut_length['Description'] = items_with_cut_length['Description'].apply(process_description_with_cut_length)
        
        # Group by Item Number and processed Description
        grouped_cut_length = items_with_cut_length.groupby(['Item Number', 'Description']).agg({
            'Total Length [mm]': 'sum',
            'Total Weight [kg]': 'sum'
        }).reset_index()
        
        # Calculate back the effective quantity (this is for reference, might not be meaningful)
        grouped_cut_length['Qty'] = np.nan  # Since we aggregated by length, qty becomes meaningless
        grouped_cut_length['Item Type'] = 'Aggregated'
        
        aggregated_results.append(grouped_cut_length)
    
    # Add items with remarks (kept separate)
    if not items_with_remarks.empty:
        items_with_remarks_processed = items_with_remarks.copy()
        # For items with cut length and remarks, still process the description
        mask_cut_length = items_with_remarks_processed['Cut Length [mm]'].notna()
        items_with_remarks_processed.loc[mask_cut_length, 'Description'] = items_with_remarks_processed.loc[mask_cut_length, 'Description'].apply(process_description_with_cut_length)
        
        # Calculate Total Length for items with cut length and remarks
        items_with_remarks_processed['Total Length [mm]'] = np.nan
        items_with_remarks_processed.loc[mask_cut_length, 'Total Length [mm]'] = (
            items_with_remarks_processed.loc[mask_cut_length, 'Qty'] * 
            items_with_remarks_processed.loc[mask_cut_length, 'Cut Length [mm]']
        )
        
        # Mark as standalone items
        items_with_remarks_processed['Item Type'] = 'Standalone'
        
        aggregated_results.append(items_with_remarks_processed)
    
    # Combine all results
    if aggregated_results:
        final_df = pd.concat(aggregated_results, ignore_index=True)
    else:
        final_df = pd.DataFrame()
    
    return final_df

def main():
    # Read the input Excel file
    input_file = 'C4_BOM_v10.0.xlsx'
    
    try:
        df = pd.read_excel(input_file)
        print(f"Successfully loaded {input_file}")
        print(f"Original data shape: {df.shape}")
        
    except FileNotFoundError:
        print(f"Error: File {input_file} not found in current directory")
        return
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    # Clean up the data
    print("Cleaning up data...")
    
    # Strip whitespace and convert empty strings to NaN
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace('', np.nan)
            df[col] = df[col].replace('nan', np.nan)
    
    # Convert data types
    df['Item Number'] = pd.to_numeric(df['Item Number'], errors='coerce').astype('Int64')
    df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').astype('Int64')
    
    # Clean numeric columns
    df['Cut Length [mm]'] = df['Cut Length [mm]'].apply(lambda x: clean_numeric_value(x, 2))
    df['Weight [kg]'] = df['Weight [kg]'].apply(lambda x: clean_numeric_value(x, 2))
    df['Total Weight [kg]'] = df['Total Weight [kg]'].apply(lambda x: clean_numeric_value(x, 2))
    
    # Remove rows where Item Number is NaN
    df = df[df['Item Number'].notna()]
    
    print(f"Data after cleaning: {df.shape}")
    
    # Process and aggregate the data
    print("Processing and aggregating data...")
    result_df = aggregate_data(df)
    
    # Prepare output columns based on item type
    # For aggregated items, exclude Cut Length; for standalone items, include it
    aggregated_items = result_df[result_df['Item Type'] == 'Aggregated'].copy()
    standalone_items = result_df[result_df['Item Type'] == 'Standalone'].copy()
    
    # Define columns for each type
    aggregated_columns = ['Item Number', 'Qty', 'Description', 'Total Length [mm]', 'Total Weight [kg]']
    standalone_columns = ['Item Number', 'Qty', 'Description', 'Cut Length [mm]', 'Total Length [mm]', 'Total Weight [kg]', 'Remarks']
    
    # Prepare aggregated items output (no Cut Length column)
    if not aggregated_items.empty:
        for col in aggregated_columns:
            if col not in aggregated_items.columns:
                aggregated_items[col] = np.nan
        aggregated_items = aggregated_items[aggregated_columns]
        aggregated_items['Remarks'] = np.nan  # Add empty Remarks column for consistency
    
    # Prepare standalone items output (include Cut Length column)
    if not standalone_items.empty:
        for col in standalone_columns:
            if col not in standalone_items.columns:
                standalone_items[col] = np.nan
        standalone_items = standalone_items[standalone_columns]
    
    # Combine results
    final_output_columns = ['Item Number', 'Qty', 'Description', 'Cut Length [mm]', 'Total Length [mm]', 'Total Weight [kg]', 'Remarks']
    
    if not aggregated_items.empty and not standalone_items.empty:
        # Add Cut Length column to aggregated items as NaN for consistency
        aggregated_items['Cut Length [mm]'] = np.nan
        aggregated_items = aggregated_items[final_output_columns]
        result_df = pd.concat([aggregated_items, standalone_items], ignore_index=True)
    elif not aggregated_items.empty:
        aggregated_items['Cut Length [mm]'] = np.nan
        aggregated_items = aggregated_items[final_output_columns]
        result_df = aggregated_items
    elif not standalone_items.empty:
        result_df = standalone_items[final_output_columns]
    else:
        result_df = pd.DataFrame(columns=final_output_columns)
    
    # Sort alphabetically by Description
    result_df = result_df.sort_values('Description').reset_index(drop=True)
    
    # Post-processing: Check for duplicate item numbers and mark them
    print("Checking for duplicate item numbers...")
    item_counts = result_df['Item Number'].value_counts()
    duplicates = item_counts[item_counts > 1].index.tolist()
    
    # Add a new column to mark duplicates
    result_df['Duplicate Item Number'] = result_df['Item Number'].isin(duplicates)
    
    if duplicates:
        print(f"Found {len(duplicates)} item numbers with duplicates:")
        for item_num in duplicates:
            count = item_counts[item_num]
            print(f"  Item Number {item_num}: appears {count} times")
    else:
        print("No duplicate item numbers found.")
    
    print(f"Final aggregated data shape: {result_df.shape}")
    
    # Generate output filename with timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    base_name = os.path.splitext(input_file)[0]
    output_file = f"{base_name}_aggregated_{timestamp}.csv"
    
    # Save to CSV
    try:
        result_df.to_csv(output_file, index=False)
        print(f"Successfully saved aggregated data to: {output_file}")
        
        # Display summary
        print("\n--- Summary ---")
        print(f"Total items processed: {len(result_df)}")
        print(f"Items with cut length: {len(result_df[result_df['Cut Length [mm]'].notna()])}")
        print(f"Items without cut length: {len(result_df[result_df['Cut Length [mm]'].isna()])}")
        print(f"Items with remarks: {len(result_df[result_df['Remarks'].notna()])}")
        
    except Exception as e:
        print(f"Error saving CSV file: {e}")

if __name__ == "__main__":
    main()
