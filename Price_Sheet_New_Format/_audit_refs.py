import re
with open('create_price_sheet_template.py', 'r') as f:
    content = f.read()

# Find all INDEX patterns for IN_Paint
patterns = re.findall(r'IN_Paint!\$A\$3:\$[A-Z]\$5000,,(\d+)', content)
print('Column refs found:', sorted(set(patterns)))
for p in sorted(set(patterns)):
    count = patterns.count(p)
    print(f'  col {p}: {count} occurrences')

ranges = re.findall(r'IN_Paint!\$A\$3:\$([A-Z])\$5000', content)
print(f'\nRange end letters: {set(ranges)}')
for letter in sorted(set(ranges)):
    print(f'  ${letter}$: {ranges.count(letter)} occurrences')
