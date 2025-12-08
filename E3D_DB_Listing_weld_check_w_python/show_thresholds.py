print('PBOR-Based Threshold Table:')
print('=' * 55)
print('PBOR (mm) | Touching Threshold | Near Threshold')
print('-' * 55)
for pbor in [10, 15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300, 350, 400, 450, 500, 600]:
    touching = pbor + 50
    near = pbor * 2 + 100
    print(f'{pbor:>8} | {touching:>17.0f}mm | {near:>13.0f}mm')

print('\nFormulas:')
print('  Touching: PBOR + 50mm')
print('  Near: PBOR × 2 + 100mm')
print('\nFor component pairs, the average PBOR is used.')
