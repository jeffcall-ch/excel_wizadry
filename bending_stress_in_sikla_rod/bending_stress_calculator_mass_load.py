import math

# Given data
L = 200  # Length of rod [mm]
d = 8.3    # Diameter of rod [mm]
mass = 8  # Load mass [kg]
g = 9.81   # Gravitational acceleration [m/s²]

# Material properties for C22 (Copper-Nickel-Zinc alloy)
# Note: These are typical values - verify with material specification sheet
print("=== Material Properties: C22 Copper Alloy ===")
yield_strength = 220  # Typical yield strength [MPa] for annealed condition
tensile_strength = 400  # Typical tensile strength [MPa]
FOS = 2.0  # Factor of Safety for static loading
allowable_stress = yield_strength / FOS  # Allowable stress [MPa]

print(f"Yield Strength: {yield_strength} MPa")
print(f"Tensile Strength: {tensile_strength} MPa")
print(f"Factor of Safety: {FOS}")
print(f"Allowable Stress: σ_allowable = {allowable_stress:.2f} MPa")
print()

# Convert load to force
P = mass * g  # Force [N]
print(f"Applied load: P = {P:.2f} N")
print()

# Geometric properties of circular cross-section
print("=== Cross-section properties ===")
radius = d / 2
A = math.pi * d**2 / 4  # Cross-sectional area [mm²]
I = math.pi * d**4 / 64  # Second moment of area [mm⁴]
print(f"Cross-sectional area: A = π × d²/4 = {A:.2f} mm²")
print(f"Second moment of area: I = π × d⁴/64 = {I:.2f} mm⁴")
print()

# Maximum bending moment (at the fixed end)
print("=== Bending Moment ===")
M = P * L  # Bending moment [N·mm]
print(f"Maximum bending moment: M = P × L = {P:.2f} × {L} = {M:.2f} N·mm")
print()

# Maximum bending stress (at the fixed end, outer fiber)
print("=== Bending Stress ===")
c = d / 2  # Distance from neutral axis to outer fiber [mm]
sigma_bending = (M * c) / I  # Bending stress [N/mm² = MPa]
# Alternative formula: sigma = 32 × P × L / (π × d³)
sigma_bending_alt = (32 * P * L) / (math.pi * d**3)
print(f"Maximum bending stress: σ = (M × c)/I = ({M:.2f} × {c})/{I:.2f}")
print(f"σ_bending = {sigma_bending:.2f} MPa")
print(f"Verification: σ = 32PL/(πd³) = {sigma_bending_alt:.2f} MPa")
print()

# Shear stress (at the fixed end)
print("=== Shear Stress ===")
V = P  # Shear force at fixed end [N]
tau_avg = V / A  # Average shear stress [MPa]
tau_max = (4/3) * tau_avg  # Maximum shear stress for circular section [MPa]
print(f"Shear force: V = {V:.2f} N")
print(f"Average shear stress: τ_avg = V/A = {tau_avg:.3f} MPa")
print(f"Maximum shear stress: τ_max = (4/3) × τ_avg = {tau_max:.3f} MPa")
print()

# Safety assessment
print("=== SAFETY ASSESSMENT ===")
print(f"Maximum bending stress (CRITICAL): {sigma_bending:.2f} MPa")
print(f"Allowable stress: {allowable_stress:.2f} MPa")
print()

# Calculate safety factor
actual_safety_factor = yield_strength / sigma_bending
print(f"Actual Safety Factor: {actual_safety_factor:.3f}")
print()

# Check if design is safe
if sigma_bending <= allowable_stress:
    print("✓ DESIGN IS SAFE - Stress is within allowable limits")
    margin = ((allowable_stress - sigma_bending) / allowable_stress) * 100
    print(f"  Safety margin: {margin:.1f}%")
else:
    print("✗ DESIGN IS UNSAFE - Stress EXCEEDS allowable limits")
    overstress = ((sigma_bending - allowable_stress) / allowable_stress) * 100
    print(f"  Overstress: {overstress:.1f}%")
    print(f"  The rod will likely experience permanent deformation or failure!")
print()

print("=== SUMMARY ===")
print(f"Applied stress: {sigma_bending:.2f} MPa")
print(f"Allowable stress: {allowable_stress:.2f} MPa")
print(f"Ratio: {sigma_bending/allowable_stress:.2f} ({sigma_bending/allowable_stress*100:.1f}%)")
print()
print("Note: Verify C22 material properties with actual material specification sheet.")