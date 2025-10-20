# Cookie Converter

This directory contains scripts to convert browser cookies into Netscape format cookies that can be used with tools like `yt-dlp`.

## Quick Start

1. **Copy Cookie Headers from Browser**:
   - Open Developer Tools (F12) in your browser
   - Go to the Network tab
   - Click on any request to the website
   - Find the "Cookie:" header in the Request Headers section
   - Copy the entire cookie string (everything after "Cookie: ")
   - You can copy multiple cookie headers from different requests

2. **Paste into `raw_cookie.txt`**:
   - Paste your copied cookie headers into `raw_cookie.txt`
   - You can paste multiple lines - each line can be a separate cookie header
   - Or paste one long line with semicolon-separated cookies

3. **Run the Converter**:
   ```bash
   python cookie_converter.py
   ```
   - Enter the domain (e.g., `.google.com`, `.microsoft.com`)
   - The script will generate `cookies.txt` in Netscape format

4. **Validate (Optional)**:
   ```bash
   python validate_cookies.py
   ```

5. **Use with yt-dlp or other tools**:
   ```bash
   yt-dlp --cookies cookies.txt <URL>
   ```

## Detailed Instructions

### How to Extract Cookies from Browser Developer Console

#### Method 1: From Network Tab (Recommended)
1. Open Developer Tools (F12) in Microsoft Edge, Chrome, or Firefox
2. Go to the **Network** tab
3. Navigate to or refresh the webpage
4. Click on any request to the domain you want cookies from
5. In the **Headers** section, find **Request Headers**
6. Locate the `Cookie:` header
7. Copy everything **after** "Cookie: " (the entire value)
8. Paste into `raw_cookie.txt`

Example of what you'll copy:
```
SOCS=CAISHAgCEhJ...; SMSV=ADHTe-Aa9Nm...; ACCOUNT_CHOOSER=AFx_qI4kjTQk...
```

#### Method 2: From Application/Storage Tab
1. Open Developer Tools (F12)
2. Go to the **Application** tab (Chrome/Edge) or **Storage** tab (Firefox)
3. Under **Storage**, click on **Cookies**
4. Select the domain (e.g., `https://drive.google.com`)
5. You'll see a table with all cookies
6. Right-click and select **Export** to save as `.txt`
   - Or manually copy cookie names and values

**Note**: If using this method, format as: `name=value; name2=value2; ...`

### How the Script Works

The `cookie_converter.py` script:
1. Reads `raw_cookie.txt`
2. Parses cookie headers (handles both semicolon-separated and newline-separated)
3. Extracts cookie name-value pairs
4. Converts to Netscape format with these fields:
   - Domain (you specify)
   - Flag (TRUE for subdomain inclusion)
   - Path (/)
   - Secure (FALSE)
   - Expiration (0 for session cookie)
   - Name (from cookie)
   - Value (from cookie)
5. Writes to `cookies.txt`

### Validation

The `validate_cookies.py` script checks:
- Correct Netscape format (7 tab-separated fields)
- Proper TRUE/FALSE flags
- Valid expiration timestamps
- Skips comments and empty lines

### Example Workflow

```bash
# 1. Copy cookies from browser and paste into raw_cookie.txt

# 2. Convert to Netscape format
python cookie_converter.py
# Enter domain when prompted: .google.com

# 3. Validate (optional)
python validate_cookies.py

# 4. Use with yt-dlp
yt-dlp --cookies cookies.txt "https://drive.google.com/file/d/xxx"
```

## Notes

- **Domain Format**: Use `.example.com` (with leading dot) to include subdomains
- **Multiple Cookie Headers**: You can paste multiple cookie headers, one per line
- **Semicolon-Separated**: The script handles semicolon-separated cookies automatically
- **Session Cookies**: By default, cookies are set with expiration=0 (session cookies)
- **Security**: Keep `cookies.txt` and `raw_cookie.txt` private - they contain authentication data

## Troubleshooting

**"No valid cookies found"**:
- Make sure you copied the cookie **values**, not the "Cookie:" label
- Check that cookies are formatted as `name=value`

**Validation fails**:
- Run `validate_cookies.py` to see specific errors
- Check that the domain format is correct

**Cookies not working with yt-dlp**:
- Ensure the domain matches the website you're downloading from
- Try with a leading dot (`.google.com` vs `google.com`)