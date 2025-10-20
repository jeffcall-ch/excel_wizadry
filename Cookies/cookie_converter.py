import re

def parse_raw_headers(raw_input):

    """

    Parse raw HTTP request headers from browser developer console.def detect_domain_from_cookies(cookies):

    Extracts domain from Host/:authority: header and cookies from Cookie: header.    """

        Try to auto-detect the domain based on cookie names.

    Args:    

        raw_input (str): Raw HTTP headers from browser    Args:

            cookies (list): List of (name, value) tuples

    Returns:        

        tuple: (domain, cookies_list, is_secure)    Returns:

            - domain (str): Extracted domain or None        str: Detected domain or None

            - cookies_list (list): List of tuples (name, value)    """

            - is_secure (bool): Whether connection is HTTPS    cookie_names = {name.lower() for name, _ in cookies}

    """    

    domain = None    # Google cookies (Drive, Gmail, YouTube, etc.)

    cookies = []    google_cookies = {'sid', 'hsid', 'ssid', 'apisid', 'sapisid', 'nid', 'socs', 'account_chooser', '__secure-1psid', '__secure-3psid'}

    is_secure = False    if cookie_names & google_cookies:

    cookie_line = None        return '.google.com'

        

    lines = raw_input.strip().split('\n')    # Microsoft/Outlook/OneDrive cookies

        microsoft_cookies = {'mspauth', 'mspcid', 'wlssc', 'msfpc', 'anon', 'mspok', 'msppre', 'msprequ'}

    for line in lines:    if cookie_names & microsoft_cookies:

        line = line.strip()        return '.microsoft.com'

        if not line:    

            continue    # Facebook cookies

            if 'c_user' in cookie_names or 'xs' in cookie_names:

        # Check for domain in Host or :authority: header        return '.facebook.com'

        if line.lower().startswith('host:') or line.lower().startswith(':authority:'):    

            # Extract domain from "Host: example.com" or ":authority: example.com"    # Twitter/X cookies

            parts = line.split(':', 1)    if 'auth_token' in cookie_names or 'ct0' in cookie_names:

            if len(parts) >= 2:        return '.twitter.com'

                domain = parts[-1].strip()    

                # Add leading dot for subdomain inclusion    # YouTube cookies (specific Google domain)

                if domain and not domain.startswith('.'):    if 'ysc' in cookie_names or 'pref' in cookie_names:

                    domain = '.' + domain        return '.youtube.com'

            

        # Check for Cookie header    # Amazon cookies

        elif line.lower().startswith('cookie:'):    if 'session-id' in cookie_names or 'ubid-main' in cookie_names:

            # Extract everything after "Cookie: "        return '.amazon.com'

            cookie_line = line.split(':', 1)[1].strip()    

            return None

        # Check for HTTPS indicator

        elif line.lower().startswith(':scheme:'):

            if 'https' in line.lower():def parse_cookie_headers(raw_input, domain=None):

                is_secure = True    """

        elif 'https://' in line.lower():    Parse cookie headers from browser developer console.

            is_secure = True    Handles multiple lines of cookie headers (either one per line or semicolon-separated).

        

    # Parse cookies if found    Args:

    if cookie_line:        raw_input (str): Raw cookie string from browser (can be multiple lines)

        cookies = parse_cookie_string(cookie_line)        domain (str): Optional domain to use. If None, defaults to '.example.com'

    else:    

        # Fallback: treat entire input as cookie string if no headers found    Returns:

        cookies = parse_cookie_string(raw_input)        list: List of tuples (name, value)

        """

    return domain, cookies, is_secure    cookies = []

    

    # Split by newlines first to handle multiple cookie headers

def parse_cookie_string(cookie_string):    lines = raw_input.strip().split('\n')

    """    

    Parse a cookie string into name-value pairs.    for line in lines:

            line = line.strip()

    Args:        if not line:

        cookie_string (str): Cookie string (semicolon-separated)            continue

                

    Returns:        # Split by semicolons to get individual cookies

        list: List of tuples (name, value)        cookie_parts = line.split(';')

    """        

    cookies = []        for part in cookie_parts:

                part = part.strip()

    # Split by newlines first to handle multiple cookie lines            if not part or '=' not in part:

    lines = cookie_string.strip().split('\n')                continue

                    

    for line in lines:            # Split only on first = to handle values with =

        line = line.strip()            name, value = part.split('=', 1)

        if not line:            name = name.strip()

            continue            value = value.strip()

                    

        # Remove "Cookie:" prefix if present            if name and value:

        if line.lower().startswith('cookie:'):                cookies.append((name, value))

            line = line.split(':', 1)[1].strip()    

                return cookies

        # Split by semicolons to get individual cookies

        cookie_parts = line.split(';')

        def convert_to_netscape_format(cookies, domain='.example.com', is_secure=False):

        for part in cookie_parts:    """

            part = part.strip()    Convert parsed cookies to Netscape format.

            if not part or '=' not in part:    

                continue    Args:

                        cookies (list): List of tuples (name, value)

            # Split only on first = to handle values with =        domain (str): Domain to use for all cookies

            name, value = part.split('=', 1)        is_secure (bool): Whether cookies were sent over HTTPS

            name = name.strip()    

            value = value.strip()    Returns:

                    str: Cookies in Netscape format

            if name and value:    """

                cookies.append((name, value))    netscape_cookies = []

        

    return cookies    # Add Netscape header

    netscape_cookies.append("# Netscape HTTP Cookie File")

    netscape_cookies.append("# This file was generated by cookie_converter.py")

def convert_to_netscape_format(cookies, domain='.example.com', is_secure=False):    netscape_cookies.append(f"# Domain: {domain}")

    """    netscape_cookies.append(f"# Secure: {is_secure}")

    Convert parsed cookies to Netscape format.    netscape_cookies.append("")

        

    Args:    for name, value in cookies:

        cookies (list): List of tuples (name, value)        # Netscape format: domain, flag, path, secure, expiration, name, value

        domain (str): Domain to use for all cookies        flag = 'TRUE'  # Include subdomains

        is_secure (bool): Whether cookies were sent over HTTPS        path = '/'

            secure = 'TRUE' if is_secure else 'FALSE'

    Returns:        expiration = '0'  # Session cookie

        str: Cookies in Netscape format        

    """        netscape_cookie = f"{domain}\t{flag}\t{path}\t{secure}\t{expiration}\t{name}\t{value}"

    netscape_cookies = []        netscape_cookies.append(netscape_cookie)

        

    # Add Netscape header    return '\n'.join(netscape_cookies)

    netscape_cookies.append("# Netscape HTTP Cookie File")

    netscape_cookies.append("# This file was generated by cookie_converter.py")

    netscape_cookies.append(f"# Domain: {domain}")def main():

    netscape_cookies.append(f"# Secure: {is_secure}")    """Main function to convert cookies from raw_cookie.txt to cookies.txt"""

    netscape_cookies.append("")    

        # Read the raw cookie from the file

    for name, value in cookies:    try:

        # Netscape format: domain, flag, path, secure, expiration, name, value        with open("raw_cookie.txt", "r", encoding='utf-8') as file:

        flag = 'TRUE'  # Include subdomains            raw_cookie = file.read()

        path = '/'    except FileNotFoundError:

        secure = 'TRUE' if is_secure else 'FALSE'        print("Error: raw_cookie.txt not found!")

        expiration = '0'  # Session cookie        print("Please create raw_cookie.txt and paste your cookie headers from the browser.")

                return

        netscape_cookie = f"{domain}\t{flag}\t{path}\t{secure}\t{expiration}\t{name}\t{value}"    

        netscape_cookies.append(netscape_cookie)    if not raw_cookie.strip():

            print("Error: raw_cookie.txt is empty!")

    return '\n'.join(netscape_cookies)        print("Please paste your cookie headers from the browser into raw_cookie.txt")

        return

    

def main():    # Parse cookies

    """Main function to convert cookies from raw_cookie.txt to cookies.txt"""    cookies = parse_cookie_headers(raw_cookie)

        

    # Read the raw input from the file    if not cookies:

    try:        print("No valid cookies found in raw_cookie.txt!")

        with open("raw_cookie.txt", "r", encoding='utf-8') as file:        return

            raw_input = file.read()    

    except FileNotFoundError:    print(f"Found {len(cookies)} cookies")

        print("Error: raw_cookie.txt not found!")    print("Sample cookies:", ', '.join([name for name, _ in cookies[:5]]))

        print("Please create raw_cookie.txt and paste your request headers from the browser.")    print()

        return    

        # Try to auto-detect domain

    if not raw_input.strip():    detected_domain = detect_domain_from_cookies(cookies)

        print("Error: raw_cookie.txt is empty!")    

        print("\nPlease paste into raw_cookie.txt:")    if detected_domain:

        print("  OPTION 1 (Recommended): Entire raw request headers from Network tab")        print(f"🔍 Auto-detected domain: {detected_domain}")

        print("  OPTION 2: Just the cookie string (name=value; name2=value2; ...)")        use_detected = input(f"Use this domain? (Y/n): ").strip().lower()

        return        if use_detected in ('', 'y', 'yes'):

                domain = detected_domain

    # Parse headers and extract domain and cookies        else:

    domain, cookies, is_secure = parse_raw_headers(raw_input)            domain = input("Enter domain (e.g., .google.com, .example.com): ").strip()

                if not domain:

    if not cookies:                domain = '.example.com'

        print("No valid cookies found in raw_cookie.txt!")    else:

        print("\nMake sure you copied either:")        print("⚠️  Could not auto-detect domain from cookies")

        print("  - The entire raw request headers (Right-click request → Copy → Copy as cURL)")        domain = input("Enter domain (press Enter for '.example.com'): ").strip()

        print("  - Or the Cookie header value (name=value; name2=value2; ...)")        if not domain:

        return            domain = '.example.com'

        

    print(f"✓ Found {len(cookies)} cookies")    # Convert to Netscape format

        netscape_cookies = convert_to_netscape_format(cookies, domain)

    # If domain was extracted, show it; otherwise prompt    

    if domain:    # Write the Netscape format cookies to a file

        print(f"✓ Auto-detected domain: {domain}")    with open("cookies.txt", "w", encoding='utf-8') as file:

        confirm = input(f"Use this domain? (Y/n): ").strip().lower()        file.write(netscape_cookies)

        if confirm and confirm != 'y' and confirm != 'yes' and confirm != '':    

            domain = input("Enter domain: ").strip()    print(f"\n✓ Successfully converted {len(cookies)} cookies to Netscape format!")

            if not domain.startswith('.'):    print(f"✓ Domain set to: {domain}")

                domain = '.' + domain    print(f"✓ Saved to cookies.txt")

    else:    print("\nYou can now use cookies.txt with tools like yt-dlp:")

        print("⚠ Could not auto-detect domain from headers")    print("  yt-dlp --cookies cookies.txt <URL>")

        domain = input("Enter domain (e.g., .google.com, .microsoft.com): ").strip()

        if not domain:

            domain = '.example.com'if __name__ == "__main__":

            print(f"Using default: {domain}")    main()
        elif not domain.startswith('.'):
            domain = '.' + domain
    
    if is_secure:
        print(f"✓ Detected HTTPS connection - cookies will be marked as secure")
    
    # Convert to Netscape format
    netscape_cookies = convert_to_netscape_format(cookies, domain, is_secure)
    
    # Write the Netscape format cookies to a file
    with open("cookies.txt", "w", encoding='utf-8') as file:
        file.write(netscape_cookies)
    
    print(f"\n✓ Successfully converted {len(cookies)} cookies to Netscape format!")
    print(f"✓ Domain: {domain}")
    print(f"✓ Secure: {is_secure}")
    print(f"✓ Saved to cookies.txt")
    print("\nYou can now use cookies.txt with tools like yt-dlp:")
    print(f"  yt-dlp --cookies cookies.txt <URL>")


if __name__ == "__main__":
    main()


