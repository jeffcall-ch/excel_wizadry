# 🗂️ Script Cataloger Prompt
I need you to act as a **script cataloger**.

## Task Flow
Follow this exact sequence of steps:

1. **Scan the folder**: List all files in the given directory.  
2. **Exclude noise**: Ignore test scripts, helper files, and data files (e.g., Excel sheets).  
3. **Identify candidates**: Find the script(s) that could be the main one.  
4. **Read line by line**: Analyze the candidate script(s) thoroughly to fully understand:  
   - What the script does  
   - Whether it requires input files (and their type/format)  
   - Whether it produces output files (and their type/format)  
   - Where output is saved  
5. **Check README**:  
   - If the folder has a `README.md`, read it.  
   - Verify if the README description is accurate compared to the script’s actual functionality.  
   - Note any mismatches and include your findings in the final description.  
6. **Check existing descriptions**: If another description exists elsewhere, compare and update it if needed.  
7. **Decide main vs. obsolete**:  
   - Determine which script is the main one to be **kept**.  
   - Mark any old or redundant scripts as candidates for **deletion**.  
8. **Write to index**: Create a script index entry using the required format.  

## Writing Style
- Descriptions must be written in **plain, non-technical English** so that non-developers can easily understand the script’s purpose.  
- The description must be **structured for readability**, with clear subsections for:  
  - **Purpose** – what the script is for  
  - **Inputs** – input file types, formats, or parameters  
  - **Outputs** – output format and where results are saved  
  - **How to Run** – command or CLI syntax, options, wildcards, logical operators, or arguments if supported  
  - **Additional Notes** – mismatches with README, obsolete scripts, special considerations  

## Output Format
Use **this exact format**:

```markdown
## Folder Name
### main_script_name.py
Description:
- **Purpose:** [Explain clearly in plain English what the script does]  
- **Inputs:** [List input file types, formats, or parameters if any]  
- **Outputs:** [Describe output format, where results are saved]  
- **How to Run:** [Explain how to execute the script; include CLI syntax, options, arguments, wildcards, logical operators if applicable]  
- **Additional Notes:** [Mention obsolete scripts, mismatches with README, special considerations]  

Path: [full_path_to_the_file]  
Keywords: [keyword1], [keyword2], [keyword3]  

Write the result to: /root/repos/excel_wizadry/SCRIPT_INDEX.MD
Folder to analyze: instead of one single folder, I want you to go through all folders this time in the repo
 /root/repos/excel_wizadry/