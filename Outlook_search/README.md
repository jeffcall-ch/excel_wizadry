# Outlook Advanced Email Search

A lightweight, powerful email search tool for Outlook with an ultra-compact interface and advanced query building capabilities.

![Version](https://img.shields.io/badge/version-2.4-blue)
![Platform](https://img.shields.io/badge/platform-Outlook%20VBA-orange)
![License](https://img.shields.io/badge/license-MIT-green)

---

## 🎯 What This Does

This tool provides a **visual search form** for Outlook that builds Advanced Query Syntax (AQS) search strings and executes them in Outlook's native search.

**Instead of typing complex search queries manually**, you fill in a simple form and press Enter.

### Before (Manual AQS):
```
from:"john smith" AND ((subject:budget AND subject:report) OR (body:budget AND body:report)) AND received:>=12/01/2024
```

### After (With This Tool):
```
From: john smith
Search terms: budget+report
Date from: 01.12.2024
[Press Enter]
```

---

## ✨ Features

### Search Capabilities
- ✅ **From/To/CC** - Filter by sender or recipients
- ✅ **Keywords** - Search terms with operators (OR, AND, wildcards)
- ✅ **Subject/Body** - Choose where to search
- ✅ **Date Range** - Filter by date (dd.mm.yyyy format)
- ✅ **Attachments** - Filter by file type (PDF, Excel, Word, etc.)
- ✅ **Scope** - Search all folders or current folder only
- ✅ **Exact Phrases** - Use quotes for exact matching

### Operators Supported
| Operator | Example | Result |
|----------|---------|--------|
| **Comma (,)** | `budget, report` | Finds emails with "budget" OR "report" |
| **Plus (+)** | `budget+report` | Finds emails with "budget" AND "report" |
| **Wildcard (*)** | `proj*` | Finds "project", "projector", "projects" |
| **Quotes ("")** | `"exact phrase"` | Finds exact phrase match |

### User Experience
- ✅ **Ultra-compact** - Only 480×340 pixels
- ✅ **Press Enter anywhere** - Works from any textbox, checkbox, or dropdown
- ✅ **Stay open** - Form stays open for refinement
- ✅ **Clean interface** - Modern, professional design
- ✅ **Fast** - Instant search execution

---

## 📸 Screenshot

```
┌─────────────────────────────────────────┐
│ Search mail                             │
├──────────────────┬──────────────────────┤
│ People           │ Keywords             │
│ From:[────]      │ ☐Subj ☐Body ☐Cur.Fld │
│ To:  [────]      │ Terms:[──────]       │
│ CC:  [────]      │ , = OR + = AND       │
├──────────────────┴──────────────────────┤
│ Filters                                 │
│ From:[──] To:[──]  Attachment:[──────]  │
├─────────────────────────────────────────┤
│ [Search] [Clear]                        │
└─────────────────────────────────────────┘
```

---

## 🚀 Installation

### Requirements
- Microsoft Outlook (Desktop version)
- Windows OS
- VBA enabled in Outlook

### Step 1: Open VBA Editor
1. Open **Outlook**
2. Press **`Alt + F11`** to open the VBA Editor
3. You should see "VbaProject.OTM" in the left pane (Project Explorer)

> **Tip:** If you don't see Project Explorer, press `Ctrl + R`

### Step 2: Create the UserForm

1. In VBA Editor, click **Insert** → **UserForm**
2. A new blank form appears
3. Press **`F4`** to open the Properties window
4. Find the **(Name)** property
5. Change it to: **`frmEmailSearch`**

   ⚠️ **IMPORTANT:** The name must be exactly `frmEmailSearch` (case-sensitive)

6. Right-click the form → **View Code**
7. Delete any default code in the code window
8. Open the file: **`VBA_AQS_UserForm_ULTRA_COMPACT_FIXED.vba`**
9. Copy **ALL** the code
10. Paste it into the UserForm code window

### Step 3: Add Backend Code

1. In Project Explorer (left pane), **double-click** **"ThisOutlookSession"**
2. A code window opens (might have existing code)
3. **Scroll to the very bottom** of any existing code
4. Open the file: **`VBA_AQS_ThisOutlookSession_WITH_INNER_PARENS.vba`**
5. Copy **ALL** the code
6. Paste it at the **bottom** of ThisOutlookSession

   > **Note:** Don't delete existing code in ThisOutlookSession, just add the new code at the end

### Step 4: Save and Close

1. Click **File** → **Save** (or press `Ctrl + S`)
2. Close the VBA Editor (press `Alt + Q` or click the X)

### Step 5: Create Quick Access Button (Optional but Recommended)

1. In Outlook, **right-click** the Quick Access Toolbar (top-left corner)
2. Choose **"Customize Quick Access Toolbar..."**
3. In the dropdown, select **"Macros"**
4. Find and select **"ShowEmailSearchForm"**
5. Click **"Add >>"**
6. Click **"OK"**

Now you have a button to launch the search form instantly!

### Step 6: Test It

**Option A:** Click your new Quick Access Toolbar button

**Option B:** 
1. Press **`Alt + F8`**
2. Type: `ShowEmailSearchForm`
3. Click **"Run"**

The search form should appear! 🎉

---

## 📖 How to Use

### Opening the Form

**Method 1:** Click your Quick Access Toolbar button

**Method 2:** Press `Alt + F8` → Type `ShowEmailSearchForm` → Run

### Basic Search

1. **Enter criteria** in any field:
   - **From/To/CC:** Email address or name
   - **Search terms:** Keywords (use operators)
   - **Dates:** Format as `dd.mm.yyyy`
   - **Attachment:** Select from dropdown
   - **Checkboxes:** Choose where to search

2. **Press Enter** or click **Search**

3. **Results appear** in Outlook's native search results

4. **Form stays open** - refine and search again

### Search Examples

#### Example 1: Find emails from someone
```
From: john.smith
[Press Enter]
```

#### Example 2: Find keywords with AND
```
Search terms: budget+report
☑ Subject  ☑ Body
[Press Enter]
```
Result: Finds emails with both "budget" AND "report"

#### Example 3: Find keywords with OR
```
Search terms: budget, report
☑ Subject
[Press Enter]
```
Result: Finds emails with "budget" OR "report" in subject

#### Example 4: Find exact phrase
```
Search terms: "project update"
☑ Subject  ☑ Body
[Press Enter]
```
Result: Finds exact phrase "project update"

#### Example 5: Date range with attachments
```
From: manager
Date from: 01.12.2024
Date to: 15.12.2024
Attachment: PDF Files
[Press Enter]
```
Result: Finds PDFs from manager in December 2024

#### Example 6: Current folder only
```
Search terms: invoice
☑ Current folder
[Press Enter]
```
Result: Searches only in the folder you're currently viewing

#### Example 7: Complex query
```
From: client
Search terms: proposal+final, draft
☑ Subject
Date from: 01.11.2024
Attachment: Word Files
[Press Enter]
```
Result: Word documents from client with (proposal AND final) OR draft in subject since Nov 2024

---

## 🎓 Search Operators Guide

### Comma (,) = OR
**Example:** `budget, report, invoice`  
**Finds:** Emails with "budget" OR "report" OR "invoice"

### Plus (+) = AND
**Example:** `budget+Q4+2024`  
**Finds:** Emails with all three terms: "budget" AND "Q4" AND "2024"

### Wildcard (*)
**Example:** `proj*`  
**Finds:** "project", "projector", "projects", etc.

**Note:** Only suffix wildcards work (`proj*`). Prefix wildcards (`*proj`) are not supported by Outlook AQS.

### Exact Phrases ("")
**Example:** `"quarterly report"`  
**Finds:** Exact phrase "quarterly report" (not "quarterly sales report")

### Combining Operators
**Example:** `budget+Q4, annual+review`  
**Finds:** (budget AND Q4) OR (annual AND review)

---

## ⚙️ Features Explained

### Search In (Subject/Body)
- **Subject only:** Searches only in email subject lines (faster)
- **Body only:** Searches only in email body text
- **Both:** Searches in both subject and body (default, most thorough)

### Current Folder Checkbox
- **Unchecked (default):** Searches **all folders** in your mailbox
- **Checked:** Searches **only the current folder** you're viewing (and its subfolders)

### Attachment Filters
| Option | What It Does |
|--------|--------------|
| **Any** | No filter (default) |
| **With Attachments** | Only emails that have attachments |
| **Without Attachments** | Only emails with no attachments |
| **PDF Files** | Only emails with PDF attachments |
| **Excel Files** | Only emails with .xls or .xlsx files |
| **Word Files** | Only emails with .doc or .docx files |
| **PowerPoint** | Only emails with .ppt or .pptx files |
| **Images** | Only emails with .jpg, .png, or .gif files |
| **Text Files** | Only emails with .txt files |
| **ZIP/Archives** | Only emails with .zip or .rar files |

### Date Format
Always use **dd.mm.yyyy** format:
- ✅ Correct: `15.12.2024`
- ❌ Wrong: `12/15/2024`
- ❌ Wrong: `15-12-2024`

---

## 🔧 Troubleshooting

### Form Doesn't Appear

**Problem:** Nothing happens when running the macro

**Solutions:**
1. Check UserForm name is exactly `frmEmailSearch` (case-sensitive)
   - Right-click UserForm → Properties → Check (Name) property
2. Verify code is pasted in correct locations:
   - UserForm code in the UserForm module
   - Backend code in ThisOutlookSession
3. Save the VBA project (Ctrl + S)
4. Restart Outlook
5. Try running the macro again

### Compile Error

**Problem:** "Compile error: Method or data member not found"

**Solutions:**
1. **Make sure you used the FIXED version:**
   - Use `VBA_AQS_UserForm_ULTRA_COMPACT_FIXED.vba`
   - NOT the old version with `KeyPreview`
2. Check code is in correct modules:
   - UserForm code → In UserForm module
   - Backend code → In ThisOutlookSession
3. Tools → References → Uncheck any marked as "MISSING"

### Enter Key Doesn't Work

**Problem:** Pressing Enter doesn't trigger search

**Solutions:**
1. Make sure you're using the FIXED version of the form
2. Try clicking in different fields and pressing Enter
3. The Enter key works from:
   - ✅ All textboxes
   - ✅ All checkboxes (press Enter when checkbox has focus)
   - ✅ Attachment dropdown
   - ❌ Form background (not supported in VBA)

### Invalid Date Format Error

**Problem:** "Invalid Date From format"

**Solutions:**
1. Use format: `dd.mm.yyyy`
2. Example: `15.12.2024` (not `12/15/2024`)
3. Make sure to use periods (.) not slashes (/)

### No Search Results

**Problem:** Search executes but returns no results

**Solutions:**
1. **Try simpler search** - Use fewer criteria
2. **Check date range** - Make sure dates are correct
3. **Check scope** - Uncheck "Current folder" to search everywhere
4. **Verify operators:**
   - Use `,` for OR
   - Use `+` for AND
5. **Check spelling** - Keywords must match exactly (unless using wildcards)

### Search Returns Wrong Results

**Problem:** Results don't match what you expected

**Solutions:**
1. **Check the generated query** - Enable debug mode:
   ```vba
   ' Add this line in ExecuteAQSSearch() after line 169:
   Debug.Print "AQS Query: " & aqsQuery
   ```
   Then press `Ctrl + G` in VBA Editor to see the Immediate Window
2. **Verify operator precedence:**
   - `budget+Q4, annual` = (budget AND Q4) OR annual
   - Use parentheses for complex queries (type manually if needed)
3. **Use exact phrases** for strict matching:
   - `"project budget"` instead of `project budget`

---

## 💡 Tips & Tricks

### Tip 1: Start Simple
Start with basic searches (From + one keyword), then add more criteria

### Tip 2: Use Wildcards for Names
Instead of `john`, try `john*` to catch "John", "Johnny", "Johnson"

### Tip 3: Clear Between Searches
Use the **Clear** button to reset all fields quickly

### Tip 4: Keep Form Open
The form stays open after searching - perfect for refining queries

### Tip 5: Current Folder for Focused Search
If you're in a specific project folder, check "Current folder" for faster results

### Tip 6: Combine Date + Attachments
Very effective: `Date from: 01.12.2024` + `Attachment: PDF Files`

### Tip 7: Test Operators
Try different operator combinations to see what works best for your needs

### Tip 8: Save Common Searches
Write down your most-used search patterns for quick reference

### Tip 9: Subject-Only for Speed
Searching subject only is much faster than searching body

### Tip 10: Use Debug Mode
Add `Debug.Print "AQS Query: " & aqsQuery` to see exact queries being generated

---

## 🔍 Advanced: Understanding the Generated Query

When you search, the tool builds an AQS (Advanced Query Syntax) string.

### Example Search:
```
From: john.smith
Search terms: budget+report
Date from: 01.12.2024
☑ Subject  ☑ Body
```

### Generated Query:
```
from:"john.smith" AND (((subject:budget AND subject:report) OR (body:budget AND body:report))) AND received:>=12/01/2024
```

### Query Breakdown:
```
from:"john.smith"           ← From filter
AND                          ← Top-level AND
(((                          ← Triple parentheses for proper grouping
  subject:budget AND subject:report   ← Subject search
) OR (                       ← OR between subject and body
  body:budget AND body:report         ← Body search
)))                          ← Close grouping
AND                          ← Top-level AND
received:>=12/01/2024       ← Date filter
```

The **triple parentheses** `(((...)))` ensure proper precedence when combining multiple criteria.

---

## 📁 File Structure

```
Outlook-Email-Search/
│
├── VBA_AQS_UserForm_ULTRA_COMPACT_FIXED.vba
│   └── UserForm code (paste into frmEmailSearch)
│
├── VBA_AQS_ThisOutlookSession_WITH_INNER_PARENS.vba
│   └── Backend code (paste into ThisOutlookSession)
│
├── README.md
│   └── This file
│
└── Documentation/
    ├── ULTRA_COMPACT_GUIDE.md
    ├── KEYPREVIEW_FIX_EXPLANATION.md
    ├── OPERATOR_GROUPING_EXPLAINED.md
    └── CODE_TRACE_GROUPING.md
```

---

## 🆚 Version History

### Version 2.4 (Current) - Ultra-Compact with Enter Key
- Ultra-compact form (480×340 pixels)
- Reduced textbox lengths (50-70% of original)
- Enter key works from all textboxes, checkboxes, and dropdown
- Fixed KeyPreview issue (VBA compatible)
- Improved query grouping with inner parentheses

### Version 2.3 - Compact Layout
- Reduced form size (580×340)
- Horizontal filter layout
- Tighter spacing

### Version 2.2 - Layout Improved
- Scope moved to Keywords panel
- Removed separate Scope panel
- Better visual organization

### Version 2.1 - Fixed Query Grouping
- Added inner parentheses for explicit AND group handling
- Fixed quoted phrase parsing bug
- Improved AQS compliance

### Version 2.0 - String Builder
- Complete refactor to AQS string builder approach
- Removed heavy backend (1200+ lines → 400 lines)
- Simplified architecture

### Version 1.0 - Original
- Full-featured search engine with result capturing
- Export functionality
- Bulk actions

---

## ❓ FAQ

### Q: Do I need to install anything?
**A:** No external software. Just paste VBA code into Outlook.

### Q: Will this work on Mac?
**A:** No, this requires Windows Outlook with VBA support.

### Q: Will this work with Outlook.com (web)?
**A:** No, this is for the desktop Outlook application only.

### Q: Does this modify my emails?
**A:** No, this is read-only. It only builds search queries.

### Q: Can I customize the form?
**A:** Yes! The code is fully editable. Modify colors, sizes, labels, etc.

### Q: Does this replace Outlook's search?
**A:** No, it uses Outlook's native search engine. This just builds better queries.

### Q: Will my IT department allow this?
**A:** It depends on your company's VBA policy. This doesn't access external resources or make system changes.

### Q: Can I share this with colleagues?
**A:** Yes! Share the VBA files and this README.

### Q: What if I find a bug?
**A:** Check the Troubleshooting section first, then review the debug output.

### Q: Can I search multiple mailboxes?
**A:** Yes, if you have multiple mailboxes in Outlook, searches will include all of them (when "All folders" is selected).

---

## 🤝 Contributing

Found a bug? Have a suggestion? Want to improve the code?

1. Document the issue/suggestion clearly
2. Test your changes thoroughly
3. Update documentation if needed
4. Share your improvements

---

## 📄 License

This project is provided as-is for personal and commercial use.

**MIT License** - Free to use, modify, and distribute.

---

## 🙏 Credits

**Created by:** NT  
**Version:** 2.4  
**Date:** 2025-12-16  
**Platform:** Outlook VBA

---

## 📞 Support

### Before Asking for Help:
1. ✅ Read this README completely
2. ✅ Check Troubleshooting section
3. ✅ Verify you used the FIXED version
4. ✅ Test with simple searches first
5. ✅ Check debug output (Ctrl+G in VBA)

### Need Help?
Include this information:
- Outlook version
- What you're trying to search
- Error message (if any)
- Screenshot of the form
- Debug output (if possible)

---

## 🎉 That's It!

You now have a powerful, compact email search tool in Outlook.

**Quick Start Checklist:**
- [x] Installed VBA code
- [x] Created Quick Access button
- [x] Tested with a simple search
- [x] Pressed Enter from a textbox
- [x] Read the operators guide
- [x] Bookmarked this README

**Happy Searching! 🔍**

---

*Last Updated: 2025-12-16*
