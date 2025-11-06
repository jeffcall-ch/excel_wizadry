"""
Analyze PIRS email subject patterns from the CSV file.
"""

import csv
import re
from collections import defaultdict

def analyze_subjects(csv_file):
    """Analyze all subject patterns."""
    
    subjects = []
    with open(csv_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            subjects.append(row['subject'])
    
    print(f"Total subjects: {len(subjects)}\n")
    
    # Pattern analysis
    pirs_pattern = defaultdict(list)
    non_pirs = []
    
    # PIRS pattern: <Project>, <Code1>/<Code2>/<Number>: <Original Subject>
    # Alternative: <Project> - <Code1>, <Code1>-<Code2>/<Code3>/<Number>: <Original Subject>
    
    for subject in subjects:
        # Check for different patterns
        
        # Pattern 1: "Project, Code1/Code2/Number: Subject"
        match1 = re.match(r'^([^,]+),\s+([^/]+)/([^/]+)/(\d+):\s+(.+)$', subject)
        if match1:
            pirs_pattern['Pattern1: Project, Code1/Code2/Number: Subject'].append({
                'subject': subject,
                'project': match1.group(1),
                'code1': match1.group(2),
                'code2': match1.group(3),
                'number': match1.group(4),
                'original': match1.group(5)
            })
            continue
        
        # Pattern 2: "Project - Code1, Code1-Code2/Code3/Number: Subject"
        match2 = re.match(r'^([^-]+)\s*-\s*([^,]+),\s+([^/]+)/([^/]+)/(\d+):\s+(.+)$', subject)
        if match2:
            pirs_pattern['Pattern2: Project - Code1, Code1-Code2/Code3/Number: Subject'].append({
                'subject': subject,
                'project': match2.group(1),
                'code1_first': match2.group(2),
                'code1': match2.group(3),
                'code2': match2.group(4),
                'number': match2.group(5),
                'original': match2.group(6)
            })
            continue
        
        # Pattern 3: Forwarded PIRS emails (FW: or RE: prefix)
        match3 = re.match(r'^(FW:|RE:)\s+(.+)$', subject, re.IGNORECASE)
        if match3:
            prefix = match3.group(1)
            rest = match3.group(2)
            
            # Check if the rest contains PIRS pattern
            if re.search(r',\s+[^/]+/[^/]+/\d+:', rest):
                pirs_pattern[f'Pattern3: {prefix} with PIRS'].append({
                    'subject': subject,
                    'prefix': prefix,
                    'rest': rest
                })
                continue
        
        # Pattern 4: [EXT] prefix
        match4 = re.match(r'^(.*)\[EXT\](.+)$', subject, re.IGNORECASE)
        if match4:
            # Check if it contains PIRS pattern after [EXT]
            if re.search(r',\s+[^/]+/[^/]+/\d+:', subject):
                pirs_pattern['Pattern4: [EXT] with PIRS'].append({
                    'subject': subject
                })
                continue
        
        # Not PIRS
        non_pirs.append(subject)
    
    # Print summary
    print("="*80)
    print("PIRS PATTERNS IDENTIFIED:")
    print("="*80)
    
    total_pirs = 0
    for pattern_name, items in pirs_pattern.items():
        print(f"\n{pattern_name}: {len(items)} occurrences")
        total_pirs += len(items)
        # Show first 3 examples
        for i, item in enumerate(items[:3]):
            print(f"  Example {i+1}: {item['subject'][:100]}...")
    
    print(f"\n{'='*80}")
    print(f"TOTAL PIRS EMAILS: {total_pirs}")
    print(f"NON-PIRS EMAILS: {len(non_pirs)}")
    print(f"{'='*80}")
    
    if non_pirs:
        print("\nNON-PIRS SUBJECTS:")
        for subject in non_pirs:
            print(f"  - {subject}")
    
    # Detailed analysis of PIRS identifiers
    print(f"\n{'='*80}")
    print("PIRS IDENTIFIER ANALYSIS:")
    print("="*80)
    
    all_identifiers = []
    for pattern_items in pirs_pattern.values():
        for item in pattern_items:
            if 'code1' in item and 'code2' in item and 'number' in item:
                identifier = f"{item['code1']}/{item['code2']}/{item['number']}"
                all_identifiers.append({
                    'identifier': identifier,
                    'code1': item['code1'],
                    'code2': item['code2'],
                    'number': item['number']
                })
    
    # Group by code1
    code1_groups = defaultdict(set)
    for item in all_identifiers:
        code1_groups[item['code1']].add(item['code2'])
    
    print("\nCode1 values and their Code2 combinations:")
    for code1, code2_set in sorted(code1_groups.items()):
        print(f"  {code1}: {', '.join(sorted(code2_set))}")
    
    # Number length analysis
    number_lengths = [len(item['number']) for item in all_identifiers]
    print(f"\nNumber lengths: min={min(number_lengths) if number_lengths else 0}, max={max(number_lengths) if number_lengths else 0}")
    print(f"Number length distribution: {set(number_lengths)}")


if __name__ == "__main__":
    csv_file = r"c:\Users\szil\Repos\excel_wizadry\outlook_cleaned_subject_search\pirs_email_subjects.csv"
    analyze_subjects(csv_file)
