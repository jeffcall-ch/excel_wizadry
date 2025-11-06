"""
Test the nested PIRS case specifically.
"""

import re


def remove_pirs_prefix(subject):
    """Recursive PIRS removal."""
    pattern = r'^(?:(?:FW|RE):\s*(?:\[EXT\]\s*)?)?[^,]+,\s+[^/]+/[^/]+/\d{5}:\s+(.*)'
    
    cleaned = subject
    max_iterations = 10
    iteration = 0
    
    print(f"Starting with: {subject}\n")
    
    while iteration < max_iterations:
        match = re.match(pattern, cleaned, re.IGNORECASE)
        if match:
            previous = cleaned
            cleaned = match.group(1)
            iteration += 1
            print(f"Iteration {iteration}: Removed PIRS prefix")
            print(f"  Result: {cleaned}")
            if cleaned == previous:
                break
        else:
            print(f"\nNo more PIRS patterns found after {iteration} iteration(s)")
            break
    
    return cleaned


def remove_email_tags(subject):
    """Remove email tags."""
    clean = subject.lower()
    tags = [" [EXTERNAL] ", " [EXT] ", " RE: ", " FW: ", " AW: ", " WG: ", 
            " TR: ", " RV: ", " SV: ", " FS: ", " VS: ", " VL: ", " ODP: ", 
            " PD: ", " R: ", " I: "]
    clean = " " + clean
    for tag in tags:
        clean = clean.replace(tag.lower(), " ")
    return clean.strip()


# Test cases
test_cases = [
    "Abu Dhabi, KVI SK/INOV/00608: Re: Abu Dhabi, KVI SK/INOV/00607: Re: Monthly Progress Report",
    "Abu Dhabi, KVI AG/INOV/00071: Re: Abu Dhabi | BOM List Painting Material - colour correction needed",
    "FW: [EXT] Abu Dhabi, KVI AG/SIK/00021: Re: Abu Dhabi - Sikla COR1 Material - QR Labels",
    "RE: Abu Dhabi, KVI SK/INOV/00608: Re: Abu Dhabi, KVI SK/INOV/00607: Re: Monthly Progress Report",
]

print("="*100)
print("TESTING NESTED PIRS REMOVAL")
print("="*100)

for i, test in enumerate(test_cases, 1):
    print(f"\n{'='*100}")
    print(f"TEST CASE {i}:")
    print(f"{'='*100}")
    
    after_pirs = remove_pirs_prefix(test)
    final = remove_email_tags(after_pirs)
    
    print(f"\nFINAL RESULT:")
    print(f"  Original:     {test}")
    print(f"  After PIRS:   {after_pirs}")
    print(f"  Final search: {final}")
