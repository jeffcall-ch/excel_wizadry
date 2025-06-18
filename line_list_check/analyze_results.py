import pandas as pd

# Read the output Excel file
df = pd.read_excel('output/Line_List_with_Matches.xlsx')

print('Summary of check results:')
check_columns = [c for c in df.columns if c.endswith('_check')]
print(f"Check columns: {check_columns}")

# Count occurrences of each value in each check column
for col in check_columns:
    value_counts = df[col].value_counts().to_dict()
    print(f'{col}: {value_counts}')

# Sample random rows to verify the DN check
print("\nSample rows with different DN values:")
sample_df = df[['DN', 'DN_check']].drop_duplicates().head(10)
print(sample_df)

# Check if any rows have pipe classes not found in the summary
print("\nRows with pipe class not found in summary:")
missing_pipe_class = df[df['Pipe Class_check'] == 'Pipe class not found in summary']
print(f"Count: {len(missing_pipe_class)}")
if len(missing_pipe_class) > 0:
    print(missing_pipe_class[['Row Number', 'KKS', 'Pipe Class']].head())
