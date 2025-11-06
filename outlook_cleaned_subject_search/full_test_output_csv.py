"""
Full test of VBA logic - outputs CSV with original and cleaned subjects.
"""

import csv
import re


def remove_pirs_prefix(subject):
    """Recursive PIRS removal matching VBA logic."""
    # First: Remove full PIRS prefix pattern
    pattern = r'^(?:(?:FW|RE):\s*(?:\[EXT\]\s*)?)?[^,]+,\s+[^/]+/[^/]+/\d{5}:\s+(.*)'
    
    cleaned = subject
    max_iterations = 10
    iteration = 0
    
    # Keep removing PIRS prefix patterns
    while iteration < max_iterations:
        match = re.match(pattern, cleaned, re.IGNORECASE)
        if match:
            previous = cleaned
            cleaned = match.group(1)
            iteration += 1
            if cleaned == previous:
                break
        else:
            break
    
    # Second: Remove PIRS references within subject (e.g., "Re: CODE/CODE/00000 - Subject")
    # Pattern: CODE/CODE/00000 followed by optional " - " or ": "
    ref_pattern = r'[A-Z][A-Z0-9\s\-]+/[A-Z][A-Z0-9\s]+/\d{5}\s*[-:]\s*'
    cleaned = re.sub(ref_pattern, '', cleaned, flags=re.IGNORECASE)
    
    return cleaned


def remove_email_tags(subject):
    """Remove email tags matching VBA logic."""
    clean = subject.lower()
    
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
    
    clean = " " + clean
    
    for tag in tags:
        clean = clean.replace(tag.lower(), " ")
    
    clean = clean.strip()
    
    return clean


def clean_subject(subject):
    """Full cleaning process matching VBA logic."""
    # Step 1: Remove PIRS prefix (recursive)
    after_pirs = remove_pirs_prefix(subject)
    
    # Step 2 & 3: Remove tags (includes lowercase conversion)
    final = remove_email_tags(after_pirs)
    
    return final


def process_all_subjects(input_csv, output_csv):
    """Process all subjects and create output CSV."""
    
    # Read input
    subjects = []
    with open(input_csv, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            subjects.append({
                'filename': row['filename'],
                'original': row['subject'],
                'sender': row['sender'],
                'date': row['date']
            })
    
    print(f"Processing {len(subjects)} email subjects...")
    
    # Process each subject
    results = []
    for item in subjects:
        cleaned = clean_subject(item['original'])
        results.append({
            'original_subject': item['original'],
            'cleaned_subject': cleaned,
            'sender': item['sender'],
            'date': item['date']
        })
    
    # Write output CSV
    with open(output_csv, 'w', newline='', encoding='utf-8-sig') as f:
        fieldnames = ['original_subject', 'cleaned_subject', 'sender', 'date']
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        
        writer.writeheader()
        writer.writerows(results)
    
    print(f"\n✓ Output written to: {output_csv}")
    
    # Print statistics
    print(f"\nSTATISTICS:")
    print(f"  Total emails processed: {len(results)}")
    
    # Count unique cleaned subjects
    unique_cleaned = set(r['cleaned_subject'] for r in results)
    print(f"  Unique cleaned subjects: {len(unique_cleaned)}")
    
    # Count how many PIRS were removed
    pirs_removed = sum(1 for item in subjects 
                      if remove_pirs_prefix(item['original']) != item['original'])
    print(f"  PIRS prefixes removed: {pirs_removed}")
    
    # Show some examples
    print(f"\nEXAMPLES (first 10):")
    print("=" * 120)
    for i, result in enumerate(results[:10], 1):
        print(f"\n{i}.")
        print(f"  Original: {result['original_subject']}")
        print(f"  Cleaned:  {result['cleaned_subject']}")
    
    # Find and show nested PIRS cases
    print(f"\n{'='*120}")
    print("NESTED PIRS CASES:")
    print("=" * 120)
    nested_cases = []
    for item in subjects:
        original = item['original']
        # Check if after first removal there's still a PIRS pattern
        first_removal = remove_pirs_prefix(original)
        if first_removal != original:
            # Check if we can remove another PIRS
            second_removal = remove_pirs_prefix(first_removal)
            if second_removal != first_removal:
                nested_cases.append({
                    'original': original,
                    'first': first_removal,
                    'second': second_removal,
                    'final': clean_subject(original)
                })
    
    if nested_cases:
        print(f"Found {len(nested_cases)} nested PIRS cases:\n")
        for i, case in enumerate(nested_cases, 1):
            print(f"{i}. ORIGINAL:")
            print(f"   {case['original']}")
            print(f"   AFTER 1st PIRS REMOVAL:")
            print(f"   {case['first']}")
            print(f"   AFTER 2nd PIRS REMOVAL:")
            print(f"   {case['second']}")
            print(f"   FINAL CLEANED:")
            print(f"   {case['final']}\n")
    else:
        print("No nested PIRS cases found.")
    
    return results


if __name__ == "__main__":
    input_csv = r"c:\Users\szil\Repos\excel_wizadry\outlook_cleaned_subject_search\pirs_email_subjects.csv"
    output_csv = r"c:\Users\szil\Repos\excel_wizadry\outlook_cleaned_subject_search\cleaned_subjects_output.csv"
    
    print("="*120)
    print("FULL VBA LOGIC TEST - GENERATING CLEANED SUBJECTS CSV")
    print("="*120)
    print()
    
    results = process_all_subjects(input_csv, output_csv)
    
    print(f"\n{'='*120}")
    print("TEST COMPLETE!")
    print(f"Output file: {output_csv}")
    print("="*120)
