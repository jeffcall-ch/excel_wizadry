"""Fix painting column references in create_price_sheet_template.py

The actual painting input file has 17 columns (no E3D System column).
Template was written assuming E3D System existed at col 3.
This shifts all column indices by -1 for columns >= 3.

Correct mapping:
  Type         : col 3  (was 4)
  Description  : col 4  (was 5)
  Quantity [m] : col 13 (was 14)
  Surface [m2] : col 14 (was 15)
  Colour       : col 16 (was 17)
  Range end    : Q (col 17, was R=col 18)
"""
import re

with open('create_price_sheet_template.py', 'r') as f:
    content = f.read()

# We need to do a precise regex replacement that ONLY affects
# IN_Paint references. We match the full pattern including $R$5000
# and replace the column number AND the range letter in one go.

def fix_paint_ref(m):
    """Replace column index in IN_Paint references."""
    prefix = m.group(1)   # everything before the column number
    col = int(m.group(2)) # old column number
    suffix = m.group(3)   # everything after the column number

    # Column mapping: shift -1 for columns >= 3 (E3D System removed)
    col_map = {4: 3, 5: 4, 14: 13, 15: 14, 17: 16}
    new_col = col_map.get(col, col)

    # Also fix range letter R -> Q
    prefix = prefix.replace('$R$', '$Q$')

    return f"{prefix}{new_col}{suffix}"

# Match patterns like: IN_Paint!$A$3:$R$5000,,N)
# where N is 1-2 digit column number
pattern = r'(IN_Paint!\$A\$3:\$[RQ]\$5000,,)(\d+)(\))'
content_fixed = re.sub(pattern, fix_paint_ref, content)

# Also fix any remaining $R$5000 references for IN_Paint (e.g., in FILTER range)
# that don't have a column specifier
content_fixed = re.sub(
    r'(IN_Paint!\$A\$3:)\$R\$5000',
    r'\g<1>$Q$5000',
    content_fixed
)

# Count changes
old_r_count = content.count('IN_Paint!$A$3:$R$5000')
new_q_count = content_fixed.count('IN_Paint!$A$3:$Q$5000')
old_q_count = content.count('IN_Paint!$A$3:$Q$5000')
print(f"Changed {old_r_count} '$R$' references to '$Q$' (now {new_q_count} total)")

# Verify column changes
for old_col, new_col in [(4, 3), (5, 4), (14, 13), (15, 14), (17, 16)]:
    old_pattern = f'$Q$5000,,{old_col})'
    new_pattern = f'$Q$5000,,{new_col})'
    old_count = content_fixed.count(f'IN_Paint!$A$3:{old_pattern}')
    new_count = content_fixed.count(f'IN_Paint!$A$3:{new_pattern}')
    print(f"  col {old_col} -> {new_col}: {new_count} refs of ,,{new_col}) | {old_count} remaining of ,,{old_col})")

with open('create_price_sheet_template.py', 'w') as f:
    f.write(content_fixed)

print("\nFix applied successfully.")
