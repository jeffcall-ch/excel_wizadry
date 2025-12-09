import pandas as pd

df = pd.read_csv(r'C:\Users\szil\Repos\excel_wizadry\E3D_DB_Listing_weld_check_w_python\TBY\component_adjacency.csv')

touching = df[df['Relationship'] == 'Touching']
print('Touching pairs by component type:')
pair_counts = touching.groupby(['Component_1_Type', 'Component_2_Type']).size().reset_index(name='Count')
for _, row in pair_counts.iterrows():
    print(f"{row['Component_1_Type']:6s} - {row['Component_2_Type']:6s}: {row['Count']:3d}")

print(f'\nTotal touching pairs: {len(touching)}')

print('\n\nAll pairs (all relationships):')
all_pair_counts = df.groupby(['Component_1_Type', 'Component_2_Type']).size().reset_index(name='Count')
for _, row in all_pair_counts.iterrows():
    print(f"{row['Component_1_Type']:6s} - {row['Component_2_Type']:6s}: {row['Count']:3d}")
print(f'\nTotal pairs: {len(df)}')
