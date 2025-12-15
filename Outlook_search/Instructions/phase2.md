This is Phase 2 - let's add advanced search operators and attachment filtering.

CONTEXT: We have a working basic search from Phase 1. Now enhance the search logic to support:

1. ADVANCED SEARCH OPERATORS:
   Modify the search term parser to support:
   
   a) Boolean operators (case-insensitive):
      - AND: "project AND budget" (both must exist)
      - OR: "morning OR evening" (either can exist)
      - NOT: "meeting NOT cancelled" (first yes, second no)
      - Parentheses: "(project OR proposal) AND budget"
   
   b) Exact phrase matching:
      - Quoted text: "night owl" (exact phrase)
      - Handle quotes within search string properly
   
   c) Proximity search:
      - Format: "word1 word2"~N
      - Example: "project deadline"~10 (words within 10 words of each other)
      - Count words in between, ignore punctuation
   
   d) Wildcard support:
      - "*" for any characters: "proj*" matches "project", "projection"
      - "?" for single character: "b?t" matches "bat", "bit", "but"

2. SEARCH PARSER CLASS:
   Create SearchQueryParser class with:
   - parse_query(query_string) → returns structured query object
   - tokenize(query) → breaks into terms and operators
   - evaluate(text, parsed_query) → returns True if text matches query
   - Handle nested parentheses properly
   - Support escaping quotes: \"

   Example parsing:
   Input: '(morning OR evening) AND "project update"~5 NOT cancelled'
   Output: Structured tree for evaluation

3. ATTACHMENT FILTERING UI:
   Add to GUI below "Search In:" checkbuttons:
   
   "Attachments:" 
   - Radio buttons: (○) Any  (○) With Attachments  (○) Without
   - When "With Attachments" selected, show checkboxes:
     [☑] PDF  [☑] Excel (.xls, .xlsx, .xlsm)  [☑] Word (.doc, .docx)
     [☑] Text  [☑] PowerPoint (.ppt, .pptx)  [☑] Images (.jpg, .png, .gif)
     [☑] ZIP/Archives  [☑] Other

4. ATTACHMENT SEARCH LOGIC:
   - Access mailItem.Attachments collection
   - Check attachment count for "With"/"Without"
   - When specific extensions selected, check each attachment.FileName
   - Handle embedded images vs actual file attachments

5. CASE SENSITIVITY:
   - Add checkbox: [☐] Case sensitive
   - Apply to all text matching (terms, phrases, proximity)
   - Update search_emails() to honor this flag

6. ENHANCED RESULTS DISPLAY:
   - Show matching snippet in new "Preview" column
   - Highlight matching text (first 100 chars of context)
   - For proximity matches, show the matched section
   - Display attachment names as tooltip on hover

7. PERFORMANCE:
   - Add progress bar updates during search
   - Search in background thread to keep GUI responsive
   - Add [Cancel] button that appears during search
   - Limit results to 500 by default (configurable)

8. TESTING EXAMPLES:
   Provide comments showing these working:
   - 'john AND (project OR proposal) NOT "out of office"'
   - '"status update"~5' (finds "status meeting update", "status for the update")
   - 'budget AND attachment:pdf' (if we want shorthand syntax)

Update the existing code from Phase 1 to integrate these features. Keep the code clean and well-commented.
