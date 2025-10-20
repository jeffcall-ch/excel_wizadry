# Cookie Converter - Usage Guide

## Overview
This tool converts browser cookie headers from Developer Console into Netscape format for use with `yt-dlp` and other tools.

## Quick Workflow

### Step 1: Copy Cookie Headers from Browser

**In Chrome/Edge/Firefox:**
1. Press `F12` to open Developer Tools
2. Go to **Network** tab
3. Refresh the page or navigate to the site
4. Click on any request
5. Find the **Cookie:** header in Request Headers
6. Copy everything after "Cookie: "

**Example of what you copy:**
```
SOCS=CAISHAgCEhJ...; SMSV=ADHTe-Aa9Nm...; ACCOUNT_CHOOSER=AFx_qI4kjTQk...
```

### Step 2: Paste into raw_cookie.txt

Open `raw_cookie.txt` and paste your cookie headers. You can:
- Paste multiple lines (one cookie header per line)
- Paste from multiple requests from different sites
- Mix Google, Microsoft, and other cookies

**Example `raw_cookie.txt`:**
```
SOCS=CAISHAgCEhJ...; SMSV=ADHTe-Aa9Nm...; ACCOUNT_CHOOSER=AFx_qI4kjTQk...
MSFPC=GUID=c37b3511...; MSPBack=0; MSCC=213.173.163.6-CH; PPLState=1
```

### Step 3: Run the Converter

```bash
python cookie_converter.py
```

When prompted, enter the domain:
- For Google: `.google.com`
- For Microsoft: `.microsoft.com`
- For OneDrive: `.onedrive.live.com`
- Or press Enter for `.example.com`

### Step 4: Validate (Optional)

```bash
python validate_cookies.py
```

This will check that all cookies are properly formatted.

### Step 5: Use with yt-dlp

```bash
yt-dlp --cookies cookies.txt "https://drive.google.com/file/d/YOUR_FILE_ID"
```

## Features

✅ **Handles multiple cookie headers** - Paste as many as you need
✅ **Supports multiple domains** - Mix different site cookies (though you'll need to choose one domain for output)
✅ **Flexible input format** - Semicolon-separated or line-separated
✅ **Netscape format** - Compatible with curl, wget, yt-dlp, and more
✅ **Validation included** - Verify your cookies before use

## File Structure

```
Cookies/
├── cookie_converter.py      # Main converter script
├── validate_cookies.py      # Validation script
├── raw_cookie.txt           # Your input (cookie headers from browser)
├── cookies.txt              # Generated output (Netscape format)
├── README.md                # Detailed documentation
└── USAGE_GUIDE.md          # This file
```

## Tips

### Domain Format
- Use **`.google.com`** (with leading dot) to include all Google subdomains
- Without dot (`google.com`) only matches exact domain
- Leading dot is usually what you want

### Multiple Sites
If you need cookies for multiple sites:
1. Run the converter once for each domain
2. Rename `cookies.txt` to `cookies_google.txt`, `cookies_microsoft.txt`, etc.
3. Use the appropriate file for each download

### Security
⚠️ **Important:** Cookie files contain authentication tokens!
- Don't share `raw_cookie.txt` or `cookies.txt`
- Don't commit them to version control
- Delete them when done

### Troubleshooting

**"No valid cookies found"**
- Check you copied the values (not the "Cookie:" label)
- Ensure format is `name=value; name2=value2`

**Validation errors**
- Run `validate_cookies.py` to see specific issues
- Check for proper tab separation (not spaces)

**yt-dlp authentication fails**
- Ensure domain matches the site (`.google.com` for Google Drive)
- Try refreshing cookies - they may have expired
- Check you copied **all** cookies from the request

## Example Session

```powershell
PS> cd Cookies

# Copy cookies from browser to raw_cookie.txt
PS> notepad raw_cookie.txt

# Convert to Netscape format
PS> python cookie_converter.py
Found 49 cookies
Enter domain (press Enter for '.example.com'): .google.com

✓ Successfully converted 49 cookies to Netscape format!
✓ Domain set to: .google.com
✓ Saved to cookies.txt

# Validate the output
PS> python validate_cookies.py
Validating Netscape cookie format...
==================================================
==================================================

Total cookies found: 49
✓ Validation successful: All cookies are in valid Netscape format!
✓ 49 cookie(s) validated

# Use with yt-dlp
PS> yt-dlp --cookies cookies.txt "https://drive.google.com/file/d/xxx"
```

## Advanced: Automating Domain Input

If you want to avoid typing the domain each time:

```powershell
echo ".google.com" | python cookie_converter.py
```

Or modify `cookie_converter.py` to set a default domain:

```python
# In main() function, change:
domain = input("Enter domain (press Enter for '.example.com'): ").strip()
if not domain:
    domain = '.example.com'

# To:
domain = '.google.com'  # Your preferred default
```

## Next Steps

1. ✅ Scripts are ready to use
2. ✅ Copy cookie headers from browser
3. ✅ Paste into `raw_cookie.txt`
4. ✅ Run `python cookie_converter.py`
5. ✅ Use `cookies.txt` with your tools

Happy downloading! 🎉
