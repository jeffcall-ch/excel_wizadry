import pandas as pd

df = pd.read_csv(r'C:\Users\szil\Repos\excel_wizadry\E3D_DB_Listing_weld_check_w_python\TBY\component_adjacency.csv')
touching = df[df['Relationship'] == 'Touching'].head(5)

print("Verification of calculation logic:")
print("=" * 80)
for idx, row in touching.iterrows():
    print(f"\nPair: {row['Component_1_Type']} - {row['Component_2_Type']}")
    print(f"  C1 Length: {row['Component_1_Length']}mm, C2 Length: {row['Component_2_Length']}mm")
    print(f"  Expected: {row['Expected_Distance_mm']}mm")
    print(f"  Actual Distance: {row['Distance_mm']}mm")
    print(f"  Threshold (+10%): {row['Threshold_Touching']}mm")
    print(f"  Match: {row['Distance_mm']} <= {row['Threshold_Touching']} = {row['Distance_mm'] <= row['Threshold_Touching']}")
    
    # Verify calculation
    expected_calc = row['Component_1_Length'] + row['Component_2_Length']
    threshold_calc = expected_calc * 1.10
    print(f"  Verify: {row['Component_1_Length']} + {row['Component_2_Length']} = {expected_calc} (expected: {row['Expected_Distance_mm']})")
    print(f"  Verify: {expected_calc} * 1.10 = {threshold_calc:.2f} (threshold: {row['Threshold_Touching']})")
