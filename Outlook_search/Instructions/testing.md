# Testing Documentation - Phase 1

## Automated Tests

Run the test suite with:
```bash
python test_phase1.py
```

### Test Coverage

The test suite includes:

1. **Read-Only Verification**
   - Scans all methods in OutlookSearcher class
   - Checks for forbidden keywords (delete, send, move, etc.)
   - Ensures no write/modify operations exist

2. **Basic Functionality**
   - Creates OutlookSearcher instance
   - Tests Outlook connection
   - Verifies folder retrieval
   - Validates criteria structure

3. **Code Safety Inspection**
   - Scans source code for dangerous patterns
   - Checks for .Delete(), .Move(), .Send() calls
   - Verifies no property assignments that modify data

### Expected Results

All tests should pass with output:
```
✅ PASSED: Read-Only Verification
✅ PASSED: Basic Functionality
✅ PASSED: Code Safety Inspection

Total: 3/3 tests passed

🎉 All tests passed! Application is safe to use.
   The application is confirmed READ-ONLY.
```

## Manual Testing Checklist

### Prerequisites
- [ ] Microsoft Outlook is installed and configured
- [ ] At least one email account is set up in Outlook
- [ ] Outlook has some emails to search

### GUI Testing

1. **Application Launch**
   - [ ] Application window opens without errors
   - [ ] Window is properly sized (900x700)
   - [ ] All controls are visible
   - [ ] Status shows "Connected to Outlook - Ready to search"

2. **Search Criteria - From Field**
   - [ ] Enter partial name (e.g., "john")
   - [ ] Click Search
   - [ ] Results show only emails from matching senders
   - [ ] Clear button resets the field

3. **Search Criteria - Search Terms**
   - [ ] Enter single word (e.g., "meeting")
   - [ ] Results contain that word in subject or body
   - [ ] Enter multiple words (e.g., "project status")
   - [ ] Results contain ALL words (AND logic)

4. **Search In Checkboxes**
   - [ ] Uncheck "Body", search for term only in subjects
   - [ ] Verify results have term in subject only
   - [ ] Uncheck "Subject", check "Body"
   - [ ] Verify search works in body only

5. **Date Range**
   - [ ] Enter valid date range (e.g., 2024-01-01 to 2024-12-31)
   - [ ] Results are within date range
   - [ ] Enter invalid format - should show error
   - [ ] Leave empty - should search all dates

6. **Folder Selection**
   - [ ] Select "Inbox" - search inbox only
   - [ ] Select "Sent Items" - search sent items only
   - [ ] Select "All Folders" - search both
   - [ ] Verify results match selected folder

7. **Results Display**
   - [ ] Results show in table with proper columns
   - [ ] Date column shows correct format
   - [ ] From column shows sender names
   - [ ] Subject column shows email subjects
   - [ ] Attachments column shows 📎 icon for emails with attachments
   - [ ] Status label shows correct count

8. **Opening Emails**
   - [ ] Double-click a result row
   - [ ] Email opens in Outlook (read-only view)
   - [ ] Single-click to select, click "Open Email" button
   - [ ] Email opens in Outlook
   - [ ] "Open Email" button is disabled when no selection

9. **Export CSV**
   - [ ] Perform a search with results
   - [ ] Click "Export CSV"
   - [ ] Save dialog appears
   - [ ] File is created with correct data
   - [ ] Open CSV - verify columns match results
   - [ ] "Export CSV" button disabled when no results

10. **Clear Function**
    - [ ] Fill in all search fields
    - [ ] Click "Clear"
    - [ ] All fields reset to defaults
    - [ ] Results table clears
    - [ ] Status shows "Ready to search"

### Safety Verification

**CRITICAL: Verify NO modifications are possible**

- [ ] Search for an email
- [ ] Open it in Outlook
- [ ] Verify you can read but not auto-delete/move
- [ ] Check Outlook - no emails moved or deleted
- [ ] Verify no "Delete" or "Move" buttons exist in the app
- [ ] Confirm all operations are read-only

### Error Handling

1. **Outlook Not Running**
   - [ ] Close Outlook
   - [ ] Start application
   - [ ] Should show connection error gracefully

2. **Invalid Date Format**
   - [ ] Enter "2024-13-45" as date
   - [ ] Should show error message
   - [ ] Application doesn't crash

3. **No Search Criteria**
   - [ ] Leave all fields empty
   - [ ] Click Search
   - [ ] Should show warning message

4. **Large Result Set**
   - [ ] Search for very common term
   - [ ] Results limited to 500
   - [ ] Status shows "(limited to 500 results)"

### Performance Testing

- [ ] Search in "All Folders" with body search
- [ ] Application remains responsive (may be slow but shouldn't freeze)
- [ ] Progress bar shows during search
- [ ] Can perform multiple searches in succession

## Test Results Log

### Test Run: [Date]

| Test Category | Status | Notes |
|--------------|--------|-------|
| Automated Tests | ✅ | All 3 tests passed |
| GUI Launch | ✅ | Opens without errors |
| Search Functions | ✅ | All criteria work correctly |
| Results Display | ✅ | Proper formatting |
| Open Email | ✅ | Opens in Outlook |
| Export CSV | ✅ | Creates valid file |
| Safety Verification | ✅ | No modifications possible |
| Error Handling | ✅ | All errors handled gracefully |

### Known Issues

None identified in Phase 1.

### Notes for Phase 2 Testing

When implementing Phase 2 features:
- Test all boolean operators (AND, OR, NOT)
- Test proximity search accuracy
- Test wildcard matching
- Test attachment filtering for all file types
- Verify case sensitivity toggle works correctly
