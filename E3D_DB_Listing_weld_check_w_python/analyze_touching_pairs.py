import pandas as pd

df = pd.read_csv(r'C:\Users\szil\Repos\excel_wizadry\E3D_DB_Listing_weld_check_w_python\TBY\component_adjacency.csv')

print("=" * 80)
print("TOUCHING PAIRS ANALYSIS (Component-Specific Length Calculations)")
print("=" * 80)

touching = df[df['Relationship'] == 'Touching']
print(f"\nTotal touching pairs: {len(touching)}")

print("\n" + "=" * 80)
print("Touching pairs by component type:")
print("=" * 80)
pair_counts = touching.groupby(['Component_1_Type', 'Component_2_Type']).size().reset_index(name='Count')
for _, row in pair_counts.iterrows():
    print(f"{row['Component_1_Type']:6s} - {row['Component_2_Type']:6s}: {row['Count']:3d}")

print("\n" + "=" * 80)
print("Sample touching pairs (first 10):")
print("=" * 80)
cols = ['Component_1_Type', 'Component_2_Type', 'Distance_mm', 'Expected_Distance_mm', 
        'Threshold_Touching', 'Component_1_Length', 'Component_2_Length']
print(touching[cols].head(10).to_string(index=False))

print("\n" + "=" * 80)
print("ELBO-ELBO pairs (should use sum of both lengths):")
print("=" * 80)
elbo_elbo = touching[(touching['Component_1_Type'] == 'ELBO') & (touching['Component_2_Type'] == 'ELBO')]
if len(elbo_elbo) > 0:
    print(elbo_elbo[cols].head(5).to_string(index=False))
else:
    print("No ELBO-ELBO touching pairs found")

print("\n" + "=" * 80)
print("ELBO to other component pairs (should use only ELBO length):")
print("=" * 80)
elbo_other = touching[((touching['Component_1_Type'] == 'ELBO') & (touching['Component_2_Type'] != 'ELBO')) |
                      ((touching['Component_2_Type'] == 'ELBO') & (touching['Component_1_Type'] != 'ELBO'))]
if len(elbo_other) > 0:
    print(elbo_other[cols].head(5).to_string(index=False))
else:
    print("No ELBO-other touching pairs found")

print("\n" + "=" * 80)
print("Statistics:")
print("=" * 80)
print(f"Components with valid lengths: {df[df['Component_1_Length'].notna()]['Component_1_Type'].nunique()} unique types")
print(f"Total pairs analyzed: {len(df)}")
print(f"Touching: {len(touching)} ({len(touching)/len(df)*100:.1f}%)")
print(f"Near: {len(df[df['Relationship'] == 'Near'])} ({len(df[df['Relationship'] == 'Near'])/len(df)*100:.1f}%)")
print(f"Separated: {len(df[df['Relationship'] == 'Separated'])} ({len(df[df['Relationship'] == 'Separated'])/len(df)*100:.1f}%)")
