import pandas as pd

# Load both files
print("="*80)
print("ANALYZING PTS MATCHING ISSUE")
print("="*80)

df_plan = pd.read_excel('PTS_project_specific.xlsx')
df_general = pd.read_excel('PTS_General_input_data.xlsx', sheet_name=0)

print(f"\n1. PROJECT FILE STATS:")
print(f"   - Total rows: {len(df_plan)}")
print(f"   - Non-null Activity Names: {df_plan['Activity Name'].notna().sum()}")
print(f"   - Columns: {list(df_plan.columns)}")

print(f"\n2. GENERAL FILE STATS:")
print(f"   - Important activities to match: {len(df_general)}")
print(f"   - Columns: {list(df_general.columns)}")

print(f"\n3. LOOKING FOR MOBILISATION ACTIVITY:")
mobilisation_name = "Lot 5 - Erection General Piping - Contractor On Site Mobilisation"
print(f"   - Searching for: '{mobilisation_name}'")
print(f"   - In general file: {'YES' if mobilisation_name in df_general['Activity'].values else 'NO'}")

# Search in project file
mask = df_plan['Activity Name'].astype(str).str.contains('Mobilisation', case=False, na=False)
print(f"   - In project file (contains 'Mobilisation'): {mask.sum()} matches")

# Check for similar patterns
mask_lot5 = df_plan['Activity Name'].astype(str).str.contains('Lot 5', case=False, na=False)
mask_gp = df_plan['Activity Name'].astype(str).str.contains('General Piping', case=False, na=False)
mask_erection = df_plan['Activity Name'].astype(str).str.contains('Erection', case=False, na=False)
mask_contractor = df_plan['Activity Name'].astype(str).str.contains('Contractor', case=False, na=False)

print(f"\n4. PATTERN SEARCH IN PROJECT FILE:")
print(f"   - Contains 'Lot 5': {mask_lot5.sum()} rows")
print(f"   - Contains 'General Piping': {mask_gp.sum()} rows")
print(f"   - Contains 'Erection': {mask_erection.sum()} rows")
print(f"   - Contains 'Contractor': {mask_contractor.sum()} rows")

print(f"\n5. SAMPLE OF PROJECT FILE ACTIVITIES (first 20 non-null):")
activities = df_plan['Activity Name'].dropna().head(20)
for idx, act in enumerate(activities, 1):
    print(f"   {idx}. {act}")

print(f"\n6. ALL ACTIVITIES FROM GENERAL FILE:")
for idx, row in df_general.iterrows():
    print(f"   {idx+1}. {row['Activity']}")

print(f"\n7. CHECKING IF PROJECT FILE STRUCTURE IS DIFFERENT:")
print(f"   - First 10 Activity IDs: {df_plan['Activity ID'].head(10).tolist()}")
print(f"   - First few rows of data:")
print(df_plan.head(10)[['Activity ID', 'Activity Name', 'Start', 'Finish']])

print("\n" + "="*80)
print("HYPOTHESIS: The project file may not contain the activities from general file")
print("="*80)
