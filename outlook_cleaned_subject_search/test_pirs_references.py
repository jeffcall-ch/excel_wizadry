"""
Test PIRS reference removal within subjects.
"""

import re


def remove_pirs_prefix(subject):
    """Recursive PIRS removal + internal reference removal."""
    # First: Remove full PIRS prefix pattern
    pattern = r'^(?:(?:FW|RE):\s*(?:\[EXT\]\s*)?)?[^,]+,\s+[^/]+/[^/]+/\d{5}:\s+(.*)'
    
    cleaned = subject
    max_iterations = 10
    iteration = 0
    
    print(f"Starting: {subject}\n")
    
    # Keep removing PIRS prefix patterns
    while iteration < max_iterations:
        match = re.match(pattern, cleaned, re.IGNORECASE)
        if match:
            previous = cleaned
            cleaned = match.group(1)
            iteration += 1
            print(f"  Iteration {iteration}: Removed PIRS prefix -> {cleaned}")
            if cleaned == previous:
                break
        else:
            break
    
    # Second: Remove PIRS references within subject
    ref_pattern = r'[A-Z][A-Z0-9\s\-]+/[A-Z][A-Z0-9\s]+/\d{5}\s*[-:]\s*'
    before_ref_removal = cleaned
    cleaned = re.sub(ref_pattern, '', cleaned, flags=re.IGNORECASE)
    
    if before_ref_removal != cleaned:
        print(f"  Removed PIRS reference -> {cleaned}")
    
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


# Test cases with PIRS references
test_cases = [
    "Abu Dhabi - SIK, SIK-P/KVI SK/00045: Re: SIK-P/KVI AG/00006 - Abu Dhabi Document List",
    "Abu Dhabi - SIK, SIK-P/KVI SK/00048: Re: SIK-P/KVI AG/00006 - Abu Dhabi Document List",
    "Abu Dhabi, KVI SK/INOV/00608: Re: Abu Dhabi, KVI SK/INOV/00607: Re: Monthly Progress Report",
]

print("="*100)
print("TESTING PIRS REFERENCE REMOVAL")
print("="*100)

for i, test in enumerate(test_cases, 1):
    print(f"\n{'='*100}")
    print(f"TEST CASE {i}:")
    print(f"{'='*100}")
    
    after_pirs = remove_pirs_prefix(test)
    final = remove_email_tags(after_pirs)
    
    print(f"\nFINAL:")
    print(f"  Original:     {test}")
    print(f"  After PIRS:   {after_pirs}")
    print(f"  Final search: {final}")
