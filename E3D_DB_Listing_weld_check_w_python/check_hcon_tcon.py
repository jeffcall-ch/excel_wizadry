import pandas as pd

df = pd.read_csv(r'C:\Users\szil\Repos\excel_wizadry\E3D_DB_Listing_weld_check_w_python\TBY\branch_connections.csv')

print('HCON values:')
print(df['HCON'].value_counts())

print('\nTCON values:')
print(df['TCON'].value_counts())

print(f'\nTotal branches: {len(df)}')
print(f'Branches with HCON=BWD: {len(df[df["HCON"]=="BWD"])}')
print(f'Branches with TCON=BWD: {len(df[df["TCON"]=="BWD"])}')
print(f'Branches with both HCON=BWD AND TCON=BWD: {len(df[(df["HCON"]=="BWD") & (df["TCON"]=="BWD")])}')

print('\nSample branches with empty HCON/TCON:')
empty = df[(df['HCON']=='') | (df['TCON']=='')]
print(empty[['Pipe', 'Branch', 'HCON', 'TCON']].head(10))
