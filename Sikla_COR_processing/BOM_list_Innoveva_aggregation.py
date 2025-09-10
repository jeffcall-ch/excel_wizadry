import pandas as pd
import numpy as np
import re
from datetime import datetime
import os

def extract_extra_long_number(remark):
    """Extract the number from extra-long remarks like '5 off extra-long'"""
    if pd.isna(remark) or 'extra-long' not in str(remark):
        return 0
    
    import re
    # Find number before 'off extra-long'
    match = re.search(r'(\d+)\s+off\s+extra-long', str(remark))
    if match:
        return int(match.group(1))
    return 0

def combine_extra_long_remarks(remarks_list):
    """Combine extra-long remarks by adding the numbers"""
    total_extra_long = 0
    non_extra_long_remarks = []
    
    for remark in remarks_list:
        if pd.isna(remark):
            continue
        elif 'extra-long' in str(remark):
            total_extra_long += extract_extra_long_number(remark)
        else:
            non_extra_long_remarks.append(str(remark))
    
    # Build final remark
    final_remarks = []
    if total_extra_long > 0:
        final_remarks.append(f"{total_extra_long} off extra-long")
    if non_extra_long_remarks:
        final_remarks.extend(non_extra_long_remarks)
    
    if final_remarks:
        return "; ".join(final_remarks)
    else:
        return np.nan

def aggregate_duplicates_with_extra_long(df):
    """
    Aggregate duplicate items within same corrosion category, 
    handling extra-long remarks specially
    """
    result_rows = []
    
    # Group by Item Number and Corrosion Category
    grouped = df.groupby(['Item Number', 'Corrosion Category'])
    
    for (item_num, corr_cat), group in grouped:
        if len(group) == 1:
            # Not a duplicate, keep as-is
            result_rows.append(group.iloc[0])
        else:
            # Duplicate - aggregate with extra-long logic
            aggregated_row = group.iloc[0].copy()  # Start with first row
            
            # Aggregate numeric values
            aggregated_row['Qty'] = group['Qty'].sum()
            aggregated_row['Total Weight [kg]'] = group['Total Weight [kg]'].sum()
            if 'Total Length [mm]' in group.columns:
                aggregated_row['Total Length [mm]'] = group['Total Length [mm]'].sum()
            
            # Handle remarks with extra-long logic
            remarks_list = group['Remarks'].tolist()
            aggregated_row['Remarks'] = combine_extra_long_remarks(remarks_list)
            
            # Mark as no longer duplicate since we're aggregating
            aggregated_row['Duplicate Item Number'] = False
            
            result_rows.append(aggregated_row)
    
    return pd.DataFrame(result_rows).reset_index(drop=True)

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
    
    # Process items WITHOUT cut length (aggregate by item number and corrosion category)
    if not items_no_cut_length.empty:
        grouped_no_cut = items_no_cut_length.groupby(['Item Number', 'Corrosion Category']).agg({
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
        
        # Group by Item Number, processed Description, and Corrosion Category
        grouped_cut_length = items_with_cut_length.groupby(['Item Number', 'Description', 'Corrosion Category']).agg({
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
    
    # Clean Corrosion Category column
    df['Corrosion Category'] = df['Corrosion Category'].astype(str).str.strip()
    df['Corrosion Category'] = df['Corrosion Category'].replace('nan', np.nan)
    
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
    aggregated_columns = ['Item Number', 'Qty', 'Description', 'Total Length [mm]', 'Total Weight [kg]', 'Corrosion Category']
    standalone_columns = ['Item Number', 'Qty', 'Description', 'Cut Length [mm]', 'Total Length [mm]', 'Total Weight [kg]', 'Remarks', 'Corrosion Category']
    
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
    final_output_columns = ['Item Number', 'Qty', 'Description', 'Cut Length [mm]', 'Total Length [mm]', 'Total Weight [kg]', 'Remarks', 'Corrosion Category']
    
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
    
    # Post-processing: Check for duplicate item numbers within the same corrosion category
    print("Checking for duplicate item numbers within same corrosion category...")
    
    # Group by corrosion category and find duplicates within each category
    duplicate_mask = pd.Series([False] * len(result_df), index=result_df.index)
    
    for category in result_df['Corrosion Category'].unique():
        category_df = result_df[result_df['Corrosion Category'] == category]
        item_counts = category_df['Item Number'].value_counts()
        duplicates_in_category = item_counts[item_counts > 1].index.tolist()
        
        if duplicates_in_category:
            print(f"Corrosion Category '{category}':")
            for item_num in duplicates_in_category:
                count = item_counts[item_num]
                print(f"  Item Number {item_num}: appears {count} times")
            
            # Mark duplicates for this category
            category_mask = (result_df['Corrosion Category'] == category) & (result_df['Item Number'].isin(duplicates_in_category))
            duplicate_mask = duplicate_mask | category_mask
    
    # Add the duplicate flag column
    result_df['Duplicate Item Number'] = duplicate_mask
    
    total_duplicates = duplicate_mask.sum()
    if total_duplicates == 0:
        print("No duplicate item numbers found within same corrosion categories.")
    else:
        print(f"Total items marked as duplicates: {total_duplicates}")
    
    # Aggregate duplicates with extra-long logic
    print("Aggregating duplicate items with extra-long remark handling...")
    result_df = aggregate_duplicates_with_extra_long(result_df)
    
    # Clean up Total Length column - replace 0 values with NaN
    if 'Total Length [mm]' in result_df.columns:
        result_df['Total Length [mm]'] = result_df['Total Length [mm]'].replace(0, np.nan)
    
    # Reorder columns to put Corrosion Category as the very last column
    column_order = ['Item Number', 'Qty', 'Description', 'Cut Length [mm]', 'Total Length [mm]', 'Total Weight [kg]', 'Remarks', 'Duplicate Item Number', 'Corrosion Category']
    result_df = result_df[column_order]
    
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
