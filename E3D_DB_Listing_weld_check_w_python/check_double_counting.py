import pandas as pd

# Check connected branches
df = pd.read_csv(r'C:\Users\szil\Repos\excel_wizadry\E3D_DB_Listing_weld_check_w_python\TBY\connected_branches.csv')
print('Connected branches sample:')
print(df[['Branch_A','Branch_A_TCON','Branch_B','Branch_B_HCON']].head())

print(f'\nTotal connected branches: {len(df)}')
print(f'Branches with TCON=BWD: {(df["Branch_A_TCON"]=="BWD").sum()}')
print(f'Branches with HCON=BWD: {(df["Branch_B_HCON"]=="BWD").sum()}')

# These connected branches involve both a TCON and HCON that are BWD
# So they represent welds that are ALREADY in the 239 total BWD count
print(f'\nDouble counting analysis:')
print(f'Connected branches: 26 welds')
print(f'These involve: 26 TCON ends + 26 HCON ends = 52 BWD branch ends')
print(f'Total BWD ends in project: 148 HCON + 91 TCON = 239')
print(f'So {len(df)} of the 239 BWD ends are already matched as connected branches')

print(f'\nCorrect calculation should be:')
print(f'Component welds: 2234')
print(f'BWD branch ends: 239 (already includes the 26 connected branches)')
print(f'Touching pairs: -52')
print(f'Components at BWD ends: -61')
print(f'Total: 2234 + 239 - 52 - 61 = {2234 + 239 - 52 - 61}')
