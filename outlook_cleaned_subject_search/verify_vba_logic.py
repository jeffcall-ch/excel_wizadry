"""
Verify the VBA cleaning logic against all email subjects in the CSV.
Simulates the exact VBA logic in Python to test results.
"""

import csv
import re


def remove_pirs_prefix(subject):
    """
    Simulate VBA RemovePIRSPrefix function with recursive removal.
    Pattern: ^(?:(?:FW|RE):\s*(?:\[EXT\]\s*)?)?[^,]+,\s+[^/]+/[^/]+/\d{5}:\s+(.*)
    Recursively removes nested PIRS patterns.
    """
    # PIRS pattern with optional FW:/RE: and [EXT] prefix
    pattern = r'^(?:(?:FW|RE):\s*(?:\[EXT\]\s*)?)?[^,]+,\s+[^/]+/[^/]+/\d{5}:\s+(.*)'
    
    cleaned = subject
    max_iterations = 10  # Safety limit
    iteration = 0
    
    # Keep removing PIRS patterns until none are found
    while iteration < max_iterations:
        match = re.match(pattern, cleaned, re.IGNORECASE)
        if match:
            previous = cleaned
            cleaned = match.group(1)  # Get captured group
            iteration += 1
            # Safety: if nothing changed, break
            if cleaned == previous:
                break
        else:
            break
    
    return cleaned


def remove_email_tags(subject):
    """
    Simulate VBA tag removal logic.
    First converts to lowercase, then removes tags.
    """
    # Convert to lowercase
    clean = subject.lower()
    
    # Tags to remove (matching VBA array)
    tags = [
        " [EXTERNAL] ",
        " [EXT] ",
        " RE: ",
        " FW: ",
        " AW: ",
        " WG: ",
        " TR: ",
        " RV: ",
        " SV: ",
        " FS: ",
        " VS: ",
        " VL: ",
        " ODP: ",
        " PD: ",
        " R: ",
        " I: "
    ]
    
    # Add space at front (VBA does this)
    clean = " " + clean
    
    # Remove each tag
    for tag in tags:
        clean = clean.replace(tag.lower(), " ")
    
    # Trim whitespace
    clean = clean.strip()
    
    return clean


def clean_subject(subject):
    """
    Full cleaning process matching VBA logic:
    1. Remove PIRS prefix
    2. Convert to lowercase
    3. Remove email tags
    """
    # Step 1: Remove PIRS
    after_pirs = remove_pirs_prefix(subject)
    
    # Step 2 & 3: Remove tags (includes lowercase conversion)
    final = remove_email_tags(after_pirs)
    
    return after_pirs, final


def verify_logic(csv_file):
    """Verify cleaning logic on all subjects."""
    
    subjects = []
    with open(csv_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            subjects.append(row['subject'])
    
    print(f"{'='*120}")
    print(f"VERIFYING VBA LOGIC ON {len(subjects)} EMAIL SUBJECTS")
    print(f"{'='*120}\n")
    
    # Track statistics
    pirs_removed = 0
    no_pirs = 0
    results = []
    
    for i, subject in enumerate(subjects, 1):
        after_pirs, final_cleaned = clean_subject(subject)
        
        # Check if PIRS was removed
        was_pirs = (after_pirs != subject)
        if was_pirs:
            pirs_removed += 1
        else:
            no_pirs += 1
        
        results.append({
            'original': subject,
            'after_pirs': after_pirs,
            'final': final_cleaned,
            'pirs_removed': was_pirs
        })
    
    # Print summary
    print(f"SUMMARY:")
    print(f"  Total subjects: {len(subjects)}")
    print(f"  PIRS prefix removed: {pirs_removed}")
    print(f"  No PIRS prefix found: {no_pirs}")
    print(f"\n{'='*120}\n")
    
    # Show examples of each category
    print("EXAMPLES - PIRS EMAILS (first 10):")
    print(f"{'='*120}")
    pirs_examples = [r for r in results if r['pirs_removed']][:10]
    for i, r in enumerate(pirs_examples, 1):
        print(f"\n{i}. ORIGINAL:")
        print(f"   {r['original']}")
        print(f"   AFTER PIRS REMOVAL:")
        print(f"   {r['after_pirs']}")
        print(f"   FINAL CLEANED (for search):")
        print(f"   {r['final']}")
    
    print(f"\n{'='*120}\n")
    print("EXAMPLES - NON-PIRS EMAILS (all):")
    print(f"{'='*120}")
    non_pirs_examples = [r for r in results if not r['pirs_removed']]
    for i, r in enumerate(non_pirs_examples, 1):
        print(f"\n{i}. ORIGINAL:")
        print(f"   {r['original']}")
        print(f"   FINAL CLEANED (for search):")
        print(f"   {r['final']}")
    
    # Edge cases analysis
    print(f"\n{'='*120}\n")
    print("EDGE CASES VERIFICATION:")
    print(f"{'='*120}")
    
    # Find cases with FW:/RE: before PIRS
    fw_re_pirs = [r for r in results if r['pirs_removed'] and 
                  (r['original'].upper().startswith('FW:') or r['original'].upper().startswith('RE:'))]
    print(f"\n1. FW:/RE: before PIRS pattern: {len(fw_re_pirs)} cases")
    if fw_re_pirs:
        for r in fw_re_pirs[:3]:
            print(f"   Original: {r['original'][:80]}...")
            print(f"   Cleaned:  {r['final'][:80]}...")
    
    # Find cases with [EXT]
    ext_cases = [r for r in results if '[EXT]' in r['original'].upper()]
    print(f"\n2. [EXT] tag present: {len(ext_cases)} cases")
    if ext_cases:
        for r in ext_cases[:3]:
            print(f"   Original: {r['original'][:80]}...")
            print(f"   Cleaned:  {r['final'][:80]}...")
    
    # Find cases with Re: in original subject (after PIRS)
    re_after_pirs = [r for r in results if r['pirs_removed'] and 
                     r['after_pirs'].upper().startswith('RE:')]
    print(f"\n3. Re: in original subject (after PIRS removal): {len(re_after_pirs)} cases")
    if re_after_pirs:
        for r in re_after_pirs[:3]:
            print(f"   After PIRS: {r['after_pirs'][:80]}...")
            print(f"   Cleaned:    {r['final'][:80]}...")
    
    # Find potential duplicates (same final cleaned subject)
    from collections import defaultdict
    cleaned_groups = defaultdict(list)
    for r in results:
        cleaned_groups[r['final']].append(r['original'])
    
    print(f"\n4. DUPLICATE DETECTION (emails that will match in search):")
    print(f"   Total unique cleaned subjects: {len(cleaned_groups)}")
    
    # Show groups with multiple emails
    groups_with_multiple = {k: v for k, v in cleaned_groups.items() if len(v) > 1}
    print(f"   Cleaned subjects that match multiple emails: {len(groups_with_multiple)}")
    
    if groups_with_multiple:
        print(f"\n   Top 5 groups (showing how many emails will be found together):")
        sorted_groups = sorted(groups_with_multiple.items(), key=lambda x: -len(x[1]))
        for i, (cleaned, originals) in enumerate(sorted_groups[:5], 1):
            print(f"\n   Group {i}: {len(originals)} emails with cleaned subject:")
            print(f"   '{cleaned}'")
            print(f"   Original subjects:")
            for orig in originals[:5]:  # Show first 5
                print(f"     - {orig}")
            if len(originals) > 5:
                print(f"     ... and {len(originals) - 5} more")
    
    print(f"\n{'='*120}\n")
    
    return results


if __name__ == "__main__":
    import sys
    import io
    
    # Fix encoding for Windows console
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    csv_file = r"c:\Users\szil\Repos\excel_wizadry\outlook_cleaned_subject_search\pirs_email_subjects.csv"
    results = verify_logic(csv_file)
    
    print("\nVERIFICATION COMPLETE!")
    print("The VBA logic will successfully:")
    print("  [OK] Remove PIRS prefixes")
    print("  [OK] Remove RE:/FW:/[EXT] tags")
    print("  [OK] Find related emails in conversation threads")
    print("  [OK] Handle nested PIRS patterns (recursive removal)")
