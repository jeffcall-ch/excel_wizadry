import pandas as pd

# Load branch data
df = pd.read_csv(r'C:\Users\szil\Repos\excel_wizadry\E3D_DB_Listing_weld_check_w_python\TBY\branch_connections.csv')

print("=" * 80)
print("DIAGNOSTIC: Why are there 0 connected branches after BWD filtering?")
print("=" * 80)

# We need to simulate the find_connected_branches logic
# Looking for branches where:
# - Branch1 TPOS matches Branch2 HPOS (within tolerance)
# - Branch1 TCON = 'BWD' 
# - Branch2 HCON = 'BWD'

branches_with_tcon_bwd = df[df['TCON'] == 'BWD']
branches_with_hcon_bwd = df[df['HCON'] == 'BWD']

print(f"\nBranches with TCON='BWD': {len(branches_with_tcon_bwd)}")
print(f"Branches with HCON='BWD': {len(branches_with_hcon_bwd)}")

print("\nSample branches with TCON='BWD':")
print(branches_with_tcon_bwd[['Pipe', 'Branch', 'HCON', 'TCON']].head(10))

print("\nSample branches with HCON='BWD':")
print(branches_with_hcon_bwd[['Pipe', 'Branch', 'HCON', 'TCON']].head(10))

# Check if any branch has both positions defined
print("\n" + "=" * 80)
print("Checking position data completeness:")
print("=" * 80)

# Count branches with TPOS defined
tpos_defined = df[(df['Tail_Pipe_Length_mm'].notna()) & (df['Tail_Pipe_Length_mm'] > 0)]
print(f"Branches with TPOS (Tail_Pipe_Length > 0): {len(tpos_defined)}")

# Count branches with HPOS defined  
hpos_defined = df[(df['Head_Pipe_Length_mm'].notna()) & (df['Head_Pipe_Length_mm'] > 0)]
print(f"Branches with HPOS (Head_Pipe_Length > 0): {len(hpos_defined)}")

# The issue might be that branch_connections.csv doesn't have position data!
print("\nColumns in branch_connections.csv:")
print(list(df.columns))
