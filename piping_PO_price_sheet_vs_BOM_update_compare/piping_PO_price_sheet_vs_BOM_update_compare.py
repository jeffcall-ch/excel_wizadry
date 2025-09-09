import pandas as pd
import numpy as np
from datetime import datetime
import os

def compare_po_vs_bom(po_file_path, bom_file_path, output_dir=None):
    """
    Compare PO and BOM files and generate a comprehensive comparison report.
    
    Args:
        po_file_path (str): Path to the PO CSV file
        bom_file_path (str): Path to the BOM CSV file
        output_dir (str): Directory to save the output file (optional)
    
    Returns:
        str: Path to the generated output file
    """
    
    # Read the CSV files
    print("Reading PO file...")
    po_df = pd.read_csv(po_file_path)
    
    print("Reading BOM file...")
    bom_df = pd.read_csv(bom_file_path)
    
    # Display basic info about the files
    print(f"PO file contains {len(po_df)} rows")
    print(f"BOM file contains {len(bom_df)} rows")
    
    # Create a copy of PO dataframe to work with
    result_df = po_df.copy()
    
    # Step 1: Add new columns to the PO layout
    print("Step 1: Adding new BOM columns to PO layout...")
    result_df['BOM QTY [pcs/m]'] = np.nan
    result_df['BOM Total weight [kg]'] = np.nan
    
    # Step 3: Add difference columns
    result_df['QTY Difference [pcs/m]'] = np.nan
    result_df['Weight Difference [kg]'] = pd.Series(dtype='object')  # Allow mixed data types (numbers and strings)
    result_df['Status'] = ''
    
    # Step 2: Match PO lines with BOM lines and populate BOM data
    print("Step 2: Matching PO lines with BOM data...")
    
    # Create a mapping dictionary for faster lookup
    bom_mapping = {}
    duplicate_count = 0
    for idx, row in bom_df.iterrows():
        pipe_component = str(row['Pipe Component']).strip()
        if pipe_component in bom_mapping:
            duplicate_count += 1
            # For duplicates, keep the first occurrence but update the tracking
            continue
        bom_mapping[pipe_component] = {
            'Total weight [kg]': row['Total weight [kg]'],
            'QTY [pcs/m]': row['QTY [pcs/m]'],
            'matched': False  # Track which BOM items have been matched
        }
    
    if duplicate_count > 0:
        print(f"Note: Found {duplicate_count} duplicate pipe components in BOM, keeping first occurrence of each")
    
    matched_bom_items = set()
    
    for idx, row in result_df.iterrows():
        po_description = str(row['DESCRIPTION/ Pipe Component']).strip()
        
        if po_description in bom_mapping:
            # Found a match
            bom_data = bom_mapping[po_description]
            
            # Populate BOM data
            result_df.at[idx, 'BOM QTY [pcs/m]'] = bom_data['QTY [pcs/m]']
            result_df.at[idx, 'BOM Total weight [kg]'] = bom_data['Total weight [kg]']
            
            # Mark this BOM item as matched
            matched_bom_items.add(po_description)
            bom_mapping[po_description]['matched'] = True
            
            # Step 3: Calculate differences
            # Handle PO quantities - they might be in different columns
            po_qty = pd.to_numeric(row.get('PO QUANTITY\n (pcs./m)', 0), errors='coerce')
            if pd.isna(po_qty):
                po_qty = 0
                
            po_weight = pd.to_numeric(row.get('PO WEIGHT\n [kg]', 0), errors='coerce')
            if pd.isna(po_weight):
                po_weight = 0
            
            bom_qty = pd.to_numeric(bom_data['QTY [pcs/m]'], errors='coerce')
            bom_weight_raw = bom_data['Total weight [kg]']
            bom_weight = pd.to_numeric(bom_weight_raw, errors='coerce')
            
            if pd.isna(bom_qty):
                bom_qty = 0
            
            # Calculate quantity difference (BOM - PO, positive means BOM has more)
            qty_diff = bom_qty - po_qty
            
            # Only show quantity differences if they are not zero
            if qty_diff != 0:
                result_df.at[idx, 'QTY Difference [pcs/m]'] = qty_diff
            
            # Handle weight difference - only calculate if BOM has weight data
            if pd.isna(bom_weight) or bom_weight_raw == '' or str(bom_weight_raw).strip() == '':
                # BOM weight is missing - show N/A instead of calculating difference
                result_df.at[idx, 'Weight Difference [kg]'] = 'N/A (BOM weight missing)'
            else:
                # BOM has weight data - calculate difference
                weight_diff = bom_weight - po_weight
                if weight_diff != 0:
                    result_df.at[idx, 'Weight Difference [kg]'] = weight_diff
                
            result_df.at[idx, 'Status'] = 'Matched'
            
        else:
            # Step 4: Item is in PO but not in BOM (removed item)
            po_qty = pd.to_numeric(row.get('PO QUANTITY\n (pcs./m)', 0), errors='coerce')
            po_weight = pd.to_numeric(row.get('PO WEIGHT\n [kg]', 0), errors='coerce')
            
            if pd.isna(po_qty):
                po_qty = 0
            if pd.isna(po_weight):
                po_weight = 0
            
            # Show negative values to indicate removal
            if po_qty > 0:
                result_df.at[idx, 'QTY Difference [pcs/m]'] = -po_qty
            if po_weight > 0:
                result_df.at[idx, 'Weight Difference [kg]'] = -po_weight
                
            result_df.at[idx, 'Status'] = 'Removed from BOM'
    
    # Step 5: Add new BOM items that are not in PO
    print("Step 5: Adding new BOM items not present in PO...")
    
    unmatched_bom_items = []
    for pipe_component, bom_data in bom_mapping.items():
        if not bom_data['matched']:
            # Find the original BOM row for this item
            bom_rows = bom_df[bom_df['Pipe Component'] == pipe_component]
            if bom_rows.empty:
                print(f"Warning: Could not find BOM row for component: {pipe_component}")
                continue
            
            bom_row = bom_rows.iloc[0]  # Take the first match if there are duplicates
            
            # Create a new row based on PO structure but with BOM data
            new_row = pd.Series(index=result_df.columns, dtype=object)
            
            # Fill in available data from BOM
            new_row['Category'] = bom_row['Category']
            new_row['MATERIAL'] = bom_row['Material']
            new_row['TYPE'] = bom_row['Type']
            new_row['DESCRIPTION/ Pipe Component'] = pipe_component
            
            # BOM quantities
            new_row['BOM QTY [pcs/m]'] = bom_data['QTY [pcs/m]']
            new_row['BOM Total weight [kg]'] = bom_data['Total weight [kg]']
            
            # Since this is new, differences equal BOM values
            bom_qty = pd.to_numeric(bom_data['QTY [pcs/m]'], errors='coerce')
            bom_weight_raw = bom_data['Total weight [kg]']
            bom_weight = pd.to_numeric(bom_weight_raw, errors='coerce')
            
            if not pd.isna(bom_qty) and bom_qty > 0:
                new_row['QTY Difference [pcs/m]'] = bom_qty
                
            # Handle weight difference - only show if BOM has weight data
            if pd.isna(bom_weight) or bom_weight_raw == '' or str(bom_weight_raw).strip() == '':
                new_row['Weight Difference [kg]'] = 'N/A (BOM weight missing)'
            elif bom_weight > 0:
                new_row['Weight Difference [kg]'] = bom_weight
            
            new_row['Status'] = 'New in BOM'
            
            unmatched_bom_items.append(new_row)
    
    # Add new BOM items to the result dataframe
    if unmatched_bom_items:
        new_rows_df = pd.DataFrame(unmatched_bom_items)
        result_df = pd.concat([result_df, new_rows_df], ignore_index=True)
    
    # Step 6: Generate output file with smart naming
    print("Step 6: Generating output file...")
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"PO_vs_BOM_Comparison_Report_{timestamp}.csv"
    
    if output_dir:
        output_path = os.path.join(output_dir, output_filename)
    else:
        # Save in the same directory as the script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_path = os.path.join(script_dir, output_filename)
    
    # Save the result
    result_df.to_csv(output_path, index=False)
    
    # Generate summary statistics
    print("\n" + "="*60)
    print("COMPARISON SUMMARY")
    print("="*60)
    
    matched_items = len(result_df[result_df['Status'] == 'Matched'])
    removed_items = len(result_df[result_df['Status'] == 'Removed from BOM'])
    new_items = len(result_df[result_df['Status'] == 'New in BOM'])
    
    print(f"Total items in PO: {len(po_df)}")
    print(f"Total items in BOM: {len(bom_df)}")
    print(f"Matched items: {matched_items}")
    print(f"Items removed from BOM: {removed_items}")
    print(f"New items in BOM: {new_items}")
    print(f"Total items in comparison report: {len(result_df)}")
    
    print(f"\nOutput file saved as: {output_path}")
    print("="*60)
    
    return output_path

def main():
    """
    Main function to run the PO vs BOM comparison.
    """
    # Get the script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Define file paths
    po_file = os.path.join(script_dir, "PO.csv")
    bom_file = os.path.join(script_dir, "BOM.csv")
    
    # Check if files exist
    if not os.path.exists(po_file):
        print(f"Error: PO.csv not found at {po_file}")
        return
    
    if not os.path.exists(bom_file):
        print(f"Error: BOM.csv not found at {bom_file}")
        return
    
    print("Starting PO vs BOM Comparison...")
    print(f"PO file: {po_file}")
    print(f"BOM file: {bom_file}")
    
    try:
        output_path = compare_po_vs_bom(po_file, bom_file, script_dir)
        print(f"\nComparison completed successfully!")
        print(f"Report saved to: {output_path}")
        
    except Exception as e:
        print(f"Error during comparison: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
