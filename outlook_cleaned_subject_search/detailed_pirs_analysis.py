"""
Detailed PIRS pattern analysis including all edge cases.
"""

import csv
import re

def detailed_analysis(csv_file):
    """Detailed analysis of all subject patterns."""
    
    subjects = []
    with open(csv_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            subjects.append(row['subject'])
    
    print(f"DETAILED PIRS PATTERN ANALYSIS")
    print(f"="*100)
    print(f"Total subjects analyzed: {len(subjects)}\n")
    
    # Categories
    pirs_clean = []  # Pure PIRS without any prefix
    pirs_with_re_fw = []  # RE: or FW: before PIRS
    pirs_with_ext = []  # [EXT] tag somewhere in subject
    non_pirs = []
    
    # PIRS core pattern: <text>, <code>/<code>/<number>: <original subject>
    pirs_core_pattern = re.compile(r'^([^,]+),\s+([^/]+)/([^/]+)/(\d+):\s+(.+)$')
    
    for subject in subjects:
        # Check for prefixes
        has_re_fw = bool(re.match(r'^(RE:|FW:)\s+', subject, re.IGNORECASE))
        has_ext = '[EXT]' in subject.upper()
        
        # Remove common prefixes to check for PIRS pattern
        clean_subject = subject
        clean_subject = re.sub(r'^(RE:|FW:)\s+', '', clean_subject, flags=re.IGNORECASE)
        clean_subject = re.sub(r'\[EXT\]\s*', '', clean_subject, flags=re.IGNORECASE)
        
        # Check if it matches PIRS core pattern
        match = pirs_core_pattern.match(clean_subject)
        
        if match:
            project = match.group(1)
            code1 = match.group(2)
            code2 = match.group(3)
            number = match.group(4)
            original = match.group(5)
            
            info = {
                'full_subject': subject,
                'project': project,
                'code1': code1,
                'code2': code2,
                'number': number,
                'original': original,
                'pirs_prefix': f"{project}, {code1}/{code2}/{number}:",
                'has_re_fw': has_re_fw,
                'has_ext': has_ext
            }
            
            if has_re_fw:
                pirs_with_re_fw.append(info)
            elif has_ext:
                pirs_with_ext.append(info)
            else:
                pirs_clean.append(info)
        else:
            non_pirs.append(subject)
    
    # Print results
    print(f"1. CLEAN PIRS EMAILS (no RE:/FW:/[EXT]): {len(pirs_clean)}")
    print(f"   Example: {pirs_clean[0]['full_subject'][:90]}..." if pirs_clean else "")
    
    print(f"\n2. PIRS WITH RE:/FW: PREFIX: {len(pirs_with_re_fw)}")
    if pirs_with_re_fw:
        for i, item in enumerate(pirs_with_re_fw[:3]):
            print(f"   Example {i+1}: {item['full_subject'][:90]}...")
    
    print(f"\n3. PIRS WITH [EXT] TAG: {len(pirs_with_ext)}")
    if pirs_with_ext:
        for i, item in enumerate(pirs_with_ext[:3]):
            print(f"   Example {i+1}: {item['full_subject'][:90]}...")
    
    print(f"\n4. NON-PIRS EMAILS: {len(non_pirs)}")
    for subject in non_pirs:
        print(f"   - {subject}")
    
    # Pattern structure analysis
    print(f"\n{'='*100}")
    print("PIRS STRUCTURE BREAKDOWN:")
    print(f"{'='*100}")
    
    all_pirs = pirs_clean + pirs_with_re_fw + pirs_with_ext
    
    # Project name analysis
    projects = set(item['project'] for item in all_pirs)
    print(f"\nProject names found ({len(projects)}):")
    for proj in sorted(projects):
        count = sum(1 for item in all_pirs if item['project'] == proj)
        print(f"  - '{proj}' ({count} emails)")
    
    # Code combinations
    print(f"\nCode1/Code2 combinations:")
    combos = set((item['code1'], item['code2']) for item in all_pirs)
    for code1, code2 in sorted(combos):
        count = sum(1 for item in all_pirs if item['code1'] == code1 and item['code2'] == code2)
        print(f"  - {code1}/{code2} ({count} emails)")
    
    # Original subject starts with analysis
    print(f"\nOriginal subject patterns (first word after PIRS prefix):")
    first_words = {}
    for item in all_pirs:
        original = item['original']
        # Check if it starts with RE:/FW:
        if re.match(r'^(RE:|FW:)\s+', original, re.IGNORECASE):
            first_word = re.match(r'^(RE:|FW:)', original, re.IGNORECASE).group(1).upper()
        elif original.startswith('[EXT]'):
            first_word = '[EXT]'
        else:
            first_word = original.split()[0] if original.split() else ''
        
        first_words[first_word] = first_words.get(first_word, 0) + 1
    
    for word, count in sorted(first_words.items(), key=lambda x: -x[1])[:10]:
        print(f"  - '{word}': {count}")
    
    # Number analysis
    numbers = [item['number'] for item in all_pirs]
    print(f"\nNumber format:")
    print(f"  - All numbers are {len(numbers[0])} digits long: {all(len(n) == 5 for n in numbers)}")
    print(f"  - Range: {min(numbers)} to {max(numbers)}")
    
    # Separator analysis
    print(f"\nSeparator pattern:")
    print(f"  - Between project and codes: always ', ' (comma + space)")
    print(f"  - Between codes: always '/' (forward slash)")
    print(f"  - After number: always ': ' (colon + space)")
    
    return {
        'clean': pirs_clean,
        'with_re_fw': pirs_with_re_fw,
        'with_ext': pirs_with_ext,
        'non_pirs': non_pirs
    }


if __name__ == "__main__":
    csv_file = r"c:\Users\szil\Repos\excel_wizadry\outlook_cleaned_subject_search\pirs_email_subjects.csv"
    results = detailed_analysis(csv_file)
