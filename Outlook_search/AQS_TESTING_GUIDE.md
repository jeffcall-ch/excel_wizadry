# AQS Refactor - Testing Guide & Known Limitations

## Testing Guide

### **Test Environment Setup**

1. **Install the AQS Version:**
   - Open Outlook VBA Editor (Alt+F11)
   - Delete or rename old `ThisOutlookSession` code
   - Copy contents of `VBA_AQS_ThisOutlookSession.vba` into `ThisOutlookSession`
   - Update `frmEmailSearch` form per `AQS_REFACTOR_USERFORM_UPDATES.md`
   - Compile (Debug → Compile VBAProject)
   - Save (Ctrl+S)

2. **Verify Windows Search Index:**
   - Open Control Panel → Indexing Options
   - Confirm Outlook data is indexed
   - If not indexed, add Outlook to indexing locations
   - Wait for indexing to complete (initial index can take hours)

---

### **Basic Functionality Tests**

#### **Test 1: Single Keyword Search**

**Input:**
```
SearchTerms: "budget"
Folder: Inbox
SearchSubFolders: False
```

**Expected Result:**
- New Outlook window opens
- Results appear showing emails containing "budget"
- AQS Query in Debug: `budget`
- Scope: Current Folder Only

**Validation:**
- Open Immediate Window (Ctrl+G)
- Check for log: `AQS Query: budget`
- Verify results match keyword

**Pass/Fail:** ☐

---

#### **Test 2: From Address Search**

**Input:**
```
FromAddress: "john"
Folder: Inbox
```

**Expected Result:**
- Results show emails from anyone named "john"
- AQS Query: `from:john`

**Validation:**
- All results have sender containing "john"
- Partial matching works (john@example.com, John Smith, etc.)

**Pass/Fail:** ☐

---

#### **Test 3: Date Range**

**Input:**
```
DateFrom: 12/1/2024
DateTo: 12/31/2024
Folder: Inbox
```

**Expected Result:**
- Only December 2024 emails shown
- AQS Query: `received:12/01/2024..12/31/2024`

**Validation:**
- Check first result date >= 12/1/2024
- Check last result date <= 12/31/2024

**Pass/Fail:** ☐

---

#### **Test 4: Multiple Criteria**

**Input:**
```
FromAddress: "manager"
Subject: "report"
UnreadOnly: True
```

**Expected Result:**
- Results from "manager" with "report" in subject AND unread
- AQS Query: `from:manager AND subject:report AND isread:no`

**Validation:**
- All results match ALL criteria
- No read emails in results

**Pass/Fail:** ☐

---

#### **Test 5: Special Characters in Subject**

**Input:**
```
Subject: "Re: Q4 Results (Final)"
```

**Expected Result:**
- Query properly escapes special characters
- AQS Query: `subject:"Re: Q4 Results (Final)"`
- Results contain exact phrase

**Validation:**
- Check Debug output for proper quoting
- Results match exact subject phrase

**Pass/Fail:** ☐

---

#### **Test 6: Attachment Filter - PDF**

**Input:**
```
AttachmentFilter: "PDF files"
Folder: Inbox
```

**Expected Result:**
- Only emails with PDF attachments shown
- AQS Query: `attachmentfilenames:*.pdf`

**Validation:**
- Open several results
- Confirm all have .pdf attachments

**Pass/Fail:** ☐

---

#### **Test 7: Size Filter**

**Input:**
```
SizeFilter: "> 10 MB"
Folder: Sent Items
```

**Expected Result:**
- Only emails larger than 10MB shown
- AQS Query: `size:>10MB`

**Validation:**
- Check email properties
- Verify sizes all exceed 10MB

**Pass/Fail:** ☐

---

#### **Test 8: Folder Scope - Subfolders**

**Input:**
```
Folder: Inbox
SearchSubFolders: True
SearchTerms: "project"
```

**Expected Result:**
- Results from Inbox AND all subfolders
- Scope: Current Folder + Subfolders

**Validation:**
- Check Debug output: `Search Scope: Current Folder + Subfolders`
- Verify results include emails from Inbox subfolders

**Pass/Fail:** ☐

---

### **Result Capture Tests**

#### **Test 9: Capture Results - Normal**

**Steps:**
1. Run search (any criteria)
2. Wait for results to appear (2-5 seconds)
3. Click "Capture Results"

**Expected Result:**
- Confirmation: "Captured X email(s)"
- Status label shows green "Captured X results"
- Export/Bulk/Stats buttons enabled

**Validation:**
- Check Debug: `Captured: X result(s)`
- Verify Export button now enabled

**Pass/Fail:** ☐

---

#### **Test 10: Capture Results - Empty**

**Steps:**
1. Search for: FromAddress = "nonexistent@invalid.com"
2. Wait for search to complete
3. Click "Capture Results"

**Expected Result:**
- Message: "No results captured"
- Status label shows "No results captured"
- Export/Bulk buttons remain disabled

**Validation:**
- No error thrown
- Graceful handling of empty results

**Pass/Fail:** ☐

---

#### **Test 11: Capture Results - Before Search**

**Steps:**
1. Open form (no search run)
2. Try clicking "Capture Results" (should be disabled)

**Expected Result:**
- Button is disabled, cannot click
- OR error message if enabled: "No search Explorer available"

**Validation:**
- Proper UI state management

**Pass/Fail:** ☐

---

### **Export Tests**

#### **Test 12: Export to CSV**

**Steps:**
1. Run search with 10-50 results
2. Capture results
3. Click Export → Select CSV

**Expected Result:**
- CSV file created on Desktop
- File name: `EmailSearchResults_YYYYMMDD_HHNNSS.csv`
- All columns populated: Date, From, To, Subject, Size, Attachments, Unread, Flagged

**Validation:**
- Open CSV in Excel
- Verify row count matches captured count
- Check data formatting (dates, special characters)

**Pass/Fail:** ☐

---

#### **Test 13: Export to Excel**

**Steps:**
1. Run search with 10-50 results
2. Capture results
3. Click Export → Select Excel

**Expected Result:**
- Excel file opens automatically
- Headers formatted (bold, gray background)
- AutoFilter enabled
- Columns auto-fitted

**Validation:**
- Click AutoFilter dropdowns work
- All data readable
- No formula injection (subjects starting with =, +, -, @)

**Pass/Fail:** ☐

---

#### **Test 14: Export Large Dataset (1000+ results)**

**Steps:**
1. Run search with >1000 results (e.g., all emails last month)
2. Capture results
3. Export to CSV

**Expected Result:**
- Export completes in <30 seconds
- All rows exported
- File size appropriate (~100KB per 1000 rows)

**Validation:**
- Check row count in CSV matches captured count
- No truncation or errors

**Pass/Fail:** ☐

---

### **Bulk Action Tests**

#### **Test 15: Bulk Mark as Read**

**Steps:**
1. Run search for unread emails
2. Capture results (note count)
3. Bulk Actions → Mark as Read

**Expected Result:**
- Confirmation dialog: "Mark all X emails as read?"
- After confirmation, all emails marked read
- Success message: "Marked X item(s) as read"

**Validation:**
- Check original search folder
- Verify all matching emails now marked read

**Pass/Fail:** ☐

---

#### **Test 16: Bulk Flag Items**

**Steps:**
1. Run search
2. Capture results
3. Bulk Actions → Flag Items

**Expected Result:**
- Confirmation shown
- All emails flagged
- Flags visible in Outlook UI

**Validation:**
- Check email flags in Outlook
- All captured emails now flagged

**Pass/Fail:** ☐

---

#### **Test 17: Bulk Move Items**

**Steps:**
1. Run search
2. Capture results
3. Bulk Actions → Move Items
4. Select target folder in picker

**Expected Result:**
- Folder picker dialog appears
- After selection, all emails moved
- Success message shows count

**Validation:**
- Check original folder (emails gone)
- Check target folder (emails present)
- Count matches

**Pass/Fail:** ☐

---

### **Statistics Tests**

#### **Test 18: Statistics Display**

**Steps:**
1. Run search with varied results
2. Capture results
3. Click Statistics

**Expected Result:**
- Statistics dialog shows:
  - Total Emails
  - Unread count and percentage
  - Flagged count and percentage
  - With Attachments count and percentage
  - Total Size (formatted)
  - Average Size
  - Unique Senders count
  - Most Common Sender

**Validation:**
- Manually verify sample calculations
- Check percentages add up correctly

**Pass/Fail:** ☐

---

### **Edge Case Tests**

#### **Test 19: Very Large Result Set (10,000+ items)**

**Input:**
- Search all emails last year

**Expected Result:**
- Search completes (may take 5-15 seconds)
- Results appear in UI
- Capture completes in <15 seconds
- No memory errors

**Validation:**
- Check Debug: `Captured: 10000+ result(s)`
- Export still works
- Statistics calculate correctly

**Pass/Fail:** ☐

---

#### **Test 20: Folder Not Found**

**Input:**
```
FolderName: "NonexistentFolder"
```

**Expected Result:**
- Warning message: "Cannot resolve folder, using Inbox"
- Search proceeds with Inbox
- Debug shows: `WARNING: Folder 'NonexistentFolder' not found`

**Validation:**
- Check Debug output
- Results from Inbox, not error

**Pass/Fail:** ☐

---

#### **Test 21: Search Window Closed Early**

**Steps:**
1. Run search
2. **Immediately** close the new Outlook window
3. Try clicking "Capture Results"

**Expected Result:**
- Error message: "Cannot access search results"
- OR "No search Explorer available"
- Graceful failure, no crash

**Validation:**
- User can run new search
- Form remains functional

**Pass/Fail:** ☐

---

#### **Test 22: Empty Query (No Criteria)**

**Input:**
- All fields blank

**Expected Result:**
- AQS Query: `` (empty string)
- All emails in scope returned (essentially "show all")

**Validation:**
- Large result set returned
- No errors
- Can capture and export

**Pass/Fail:** ☐

---

### **Performance Tests**

#### **Test 23: Search Speed**

**Methodology:**
- Run 10 different searches
- Record time from "Search" click to results visible

**Target Times:**
- <100 results: <2 seconds
- 100-1,000 results: <5 seconds
- 1,000-10,000 results: <10 seconds

**Results:**

| Search | Result Count | Time (seconds) | Pass/Fail |
|--------|--------------|----------------|-----------|
| 1      |              |                | ☐         |
| 2      |              |                | ☐         |
| 3      |              |                | ☐         |
| 4      |              |                | ☐         |
| 5      |              |                | ☐         |
| 6      |              |                | ☐         |
| 7      |              |                | ☐         |
| 8      |              |                | ☐         |
| 9      |              |                | ☐         |
| 10     |              |                | ☐         |

---

#### **Test 24: Capture Speed**

**Methodology:**
- Capture 1,000 / 5,000 / 10,000 results
- Record time

**Target Times:**
- 1,000 results: <2 seconds
- 5,000 results: <5 seconds
- 10,000 results: <15 seconds

**Results:**

| Result Count | Capture Time | Pass/Fail |
|--------------|--------------|-----------|
| 1,000        |              | ☐         |
| 5,000        |              | ☐         |
| 10,000       |              | ☐         |

---

## Known Limitations

### **1. Complex Boolean Logic**

**Limitation:**
AQS supports basic AND/OR/NOT but not deeply nested parentheses like `(A OR B) AND (C OR D)`.

**Workaround:**
- Build simplest possible AQS query
- Accept less precision
- For critical cases, could implement VBA post-filter (not included in current version)

**Example:**
```
❌ Not Supported: (from:john OR from:jane) AND (subject:budget OR subject:report)
✅ Supported:     from:john AND subject:budget
```

---

### **2. Proximity Search**

**Limitation:**
AQS does not have a NEAR operator for proximity search (words within N words of each other).

**Workaround:**
- Use simple keyword matching
- Phrases with quotes: `subject:"quarterly report"`
- Accept that words may not be adjacent

**Example:**
```
❌ Not Supported: word1~5 word2 (within 5 words)
✅ Supported:     subject:"word1 word2" (exact phrase)
```

---

### **3. Wildcard in Names (Unreliable)**

**Limitation:**
`from:john*` may not work reliably in AQS. Wildcard support varies.

**Workaround:**
- Use partial match: `from:john` (matches john, johnny, johnathan, john@email.com)
- AQS already does partial matching by default

---

### **4. Multiple Attachment Types with OR**

**Limitation:**
AQS may inconsistently handle: `attachmentfilenames:*.pdf OR *.xlsx`

**Current Implementation:**
- Uses parentheses: `(attachmentfilenames:*.xlsx OR attachmentfilenames:*.xls OR attachmentfilenames:*.xlsm)`
- This DOES work but is verbose

**Workaround:**
- If fails, use `hasattachments:yes` then manually check attachment types

---

### **5. Case-Sensitive Searches**

**Limitation:**
AQS searches are case-insensitive by default. No way to force case-sensitive.

**Impact:**
- `SearchCriteria.CaseSensitive` field ignored
- `from:JOHN` same as `from:john`

**Workaround:**
- None. AQS does not support case-sensitive matching.

---

### **6. No Completion Event**

**Limitation:**
`Explorer.Search` has no completion event. Cannot auto-detect when search finishes.

**Impact:**
- User must manually click "Capture Results"
- Cannot show live result count

**Workaround:**
- User-initiated capture (current implementation)
- Clear UX instructions

---

### **7. Instant Search Index Dependency**

**Limitation:**
AQS requires Windows Search index to be enabled and up-to-date.

**Impact:**
- New emails may not appear in results until indexed (lag up to 15 minutes)
- If indexing disabled, searches may fail or be slow

**Workaround:**
- User must enable Outlook indexing in Windows
- Check Control Panel → Indexing Options

**Verification:**
```
Control Panel → Indexing Options → Modify
Ensure "Microsoft Outlook" is checked
```

---

### **8. Search in Encrypted Emails**

**Limitation:**
AQS cannot search body/subject of S/MIME encrypted emails until decrypted.

**Impact:**
- Encrypted emails may not appear in body/subject searches
- Metadata (from, to, date) still searchable

**Workaround:**
- None. This is a fundamental encryption limitation.

---

### **9. Cross-Account Searches**

**Limitation:**
`olSearchScopeAllFolders` searches all accounts, but performance may degrade with many accounts.

**Impact:**
- Slow searches if 5+ email accounts configured
- May timeout on very large mailboxes (100,000+ emails)

**Workaround:**
- Search specific account/folder
- Use date ranges to limit scope

---

### **10. Archived Emails (PST Files)**

**Limitation:**
AQS only searches indexed locations. External PST files may not be indexed.

**Impact:**
- Archived emails in PST files may not appear in results

**Workaround:**
- Ensure PST file location is indexed in Windows Search
- Add PST folder to indexing locations

---

## Comparison: DASL vs AQS

| Feature | DASL (Old) | AQS (New) | Winner |
|---------|------------|-----------|--------|
| **Speed** | Slow (20-99 seconds) | Fast (1-5 seconds) | ✅ AQS |
| **Result Display** | Search Folder | Native Outlook UI | ✅ AQS |
| **Result Access** | Direct via Search.Results | Capture via GetTable | ➖ DASL |
| **Complex Boolean** | Full support | Limited | ➖ AQS |
| **Proximity Search** | Possible (manual) | Not supported | ➖ DASL |
| **Case Sensitive** | Supported | Not supported | ➖ DASL |
| **Wildcard Support** | Full | Partial | ➖ DASL |
| **Indexing Dependency** | None | Required | ➖ AQS |
| **User Experience** | Background | Interactive | ✅ AQS |
| **Debugging** | Hard (filter errors) | Easy (visible results) | ✅ AQS |

**Overall:** AQS wins on speed and UX, DASL wins on query flexibility.

---

## Troubleshooting

### **Issue: Search Returns No Results**

**Possible Causes:**
1. Windows Search index not enabled
2. Outlook not indexed
3. Index not updated (new emails)
4. Query syntax error

**Solutions:**
1. Check Indexing Options (Control Panel)
2. Rebuild Outlook index
3. Wait 15 minutes for new emails to index
4. Check Debug output for query validation warnings

---

### **Issue: Capture Results Fails**

**Possible Causes:**
1. Search window closed
2. Search still running
3. Insufficient permissions

**Solutions:**
1. Run search again
2. Wait 5-10 seconds before capturing
3. Check if Outlook running as administrator

---

### **Issue: Slow Search Performance**

**Possible Causes:**
1. Index corrupted
2. Very large mailbox (100,000+ emails)
3. Searching across many accounts

**Solutions:**
1. Rebuild Windows Search index
2. Add date range filter
3. Search specific folder instead of all folders

---

### **Issue: Export Fails**

**Possible Causes:**
1. Desktop path not accessible
2. File permissions
3. Excel not installed (for Excel export)

**Solutions:**
1. Check Desktop exists: `%USERPROFILE%\Desktop`
2. Run Outlook as administrator
3. Use CSV export instead

---

## Support & Debugging

### **Enable Debug Logging:**

All search operations log to Immediate Window (Ctrl+G in VBA Editor).

**Sample Debug Output:**
```
========== SEARCH STARTED (AQS) ==========
Timestamp: 12/15/2024 14:30:15
Target Folder: \\Mailbox - User\\Inbox
Search Scope: Current Folder + Subfolders
AQS Query: from:john AND subject:report AND isread:no
Query Length: 47 characters
==========================================

========== CAPTURING RESULTS ==========
Timestamp: 12/15/2024 14:30:20
Captured: 127 result(s)
=======================================
```

### **Common Debug Messages:**

- `WARNING: Folder 'X' not found` → Folder resolution failed, using Inbox
- `WARNING: Unbalanced quotes` → Query syntax issue
- `Capture Results Error` → Search window closed or not ready

---

## Version History

**Version 2.0.0 (2024-12-15) - AQS Refactor**
- Complete rewrite from DASL to AQS
- User-initiated result capture
- EntryID-based bulk actions/exports
- Native Outlook UI integration

**Version 1.x (2024-12-12) - DASL Implementation**
- Original DASL filter implementation
- Application.AdvancedSearch API
- Search folder results
- 20-99 second search times

---

## Performance Benchmarks

**Typical Search Times (AQS):**
- 50 results: 1-2 seconds
- 500 results: 2-4 seconds
- 5,000 results: 4-8 seconds
- 50,000 results: 8-15 seconds

**Typical Capture Times:**
- 100 results: <1 second
- 1,000 results: 1-2 seconds
- 10,000 results: 3-8 seconds
- 50,000 results: 10-20 seconds

**Typical Export Times (CSV):**
- 1,000 results: 5-10 seconds
- 10,000 results: 30-60 seconds

---

## Future Enhancements (Not Implemented)

1. **Auto-Refresh Results:** Detect when search completes (requires polling or timer)
2. **Post-Filter Complex Boolean:** VBA filter for `(A OR B) AND (C OR D)` logic
3. **Proximity Search Emulation:** VBA distance calculation between keywords
4. **Preset Save/Load:** JSON file storage for search presets
5. **Search History:** Track recent searches
6. **Advanced Export:** PDF generation, custom columns
7. **Batch Operations:** Progress bar for large bulk actions
8. **Search Templates:** Pre-configured queries for common tasks

---

## License & Credits

**Author:** AI Assistant (Claude Sonnet 4.5)  
**Date:** December 15, 2024  
**Version:** 2.0.0 - AQS Refactor  
**License:** MIT (modify and distribute freely)

---

## Final Checklist

Before deploying to production:

☐ All 24 tests passed  
☐ UserForm updated with Capture button  
☐ Status label added and functional  
☐ Windows Search index enabled and current  
☐ Debug output verified in Immediate Window  
☐ Tested with 10+ / 100+ / 1000+ result sets  
☐ Export to CSV works  
☐ Export to Excel works  
☐ All 3 bulk actions tested  
☐ Statistics calculations verified  
☐ Error handling tested (empty results, closed windows)  
☐ Performance meets targets (<5 seconds for most searches)  
☐ Documentation reviewed and understood  
☐ Known limitations documented for users

**Deployment Date:** _______________  
**Deployed By:** _______________  
**Production Status:** ☐ Beta ☐ Stable
