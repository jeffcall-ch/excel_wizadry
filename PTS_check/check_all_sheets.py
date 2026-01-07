import pandas as pd

xl = pd.ExcelFile('PTS_project_specific.xlsx')
print("="*80)
print("CHECKING ALL SHEETS FOR MOBILISATION ACTIVITY")
print("="*80)

mobilisation_name = "Lot 5 - Erection General Piping - Contractor On Site Mobilisation"

for sheet_name in xl.sheet_names:
    print(f"\n{'='*80}")
    print(f"SHEET: '{sheet_name}'")
    print('='*80)
    
    df = pd.read_excel('PTS_project_specific.xlsx', sheet_name=sheet_name)
    print(f"Rows: {len(df)}")
    print(f"Columns: {list(df.columns)[:5]}...")
    
    if 'Activity Name' in df.columns:
        non_null = df['Activity Name'].notna().sum()
        print(f"Non-null Activity Names: {non_null}")
        
        # Search for mobilisation
        mask_mob = df['Activity Name'].astype(str).str.contains('Mobilisation', case=False, na=False)
        print(f"Contains 'Mobilisation': {mask_mob.sum()} rows")
        
        # Search for Lot 5
        mask_lot5 = df['Activity Name'].astype(str).str.contains('Lot 5', case=False, na=False)
        print(f"Contains 'Lot 5': {mask_lot5.sum()} rows")
        
        # Search for General Piping
        mask_gp = df['Activity Name'].astype(str).str.contains('General Piping', case=False, na=False)
        print(f"Contains 'General Piping': {mask_gp.sum()} rows")
        
        # If found, show the activities
        if mask_mob.sum() > 0:
            print("\n*** FOUND MOBILISATION ACTIVITIES ***")
            mob_activities = df[mask_mob][['Activity ID', 'Activity Name']]
            for idx, row in mob_activities.iterrows():
                excel_row = idx + 2  # +2 because Excel is 1-indexed and has header
                print(f"  Excel Row {excel_row}: {row['Activity ID']} - {row['Activity Name']}")
        
        if mask_lot5.sum() > 0 and mask_mob.sum() == 0:
            print("\nSample 'Lot 5' activities:")
            lot5_activities = df[mask_lot5].head(5)
            for idx, row in lot5_activities.iterrows():
                print(f"  {row['Activity Name']}")
    else:
        print("No 'Activity Name' column in this sheet")

print("\n" + "="*80)
print("ANALYSIS COMPLETE")
print("="*80)
