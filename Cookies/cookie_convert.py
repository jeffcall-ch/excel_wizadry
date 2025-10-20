"""
Cookie Converter - Converts raw HTTP headers to Netscape cookie format

Supports:
- Raw HTTP headers from browser Network tab (Recommended)
- Just cookie strings (fallback)

Usage:
    python cookie_convert.py
"""

import re


def parse_raw_headers(raw_input):
    """
    Parse raw HTTP request headers from browser developer console.
    Can handle MULTIPLE request headers in one file.
    
    Args:
        raw_input (str): Raw HTTP headers from browser (can contain multiple requests)
    
    Returns:
        list: List of tuples (domain, cookies_list, is_secure)
            - domain (str): Extracted domain with leading dot
            - cookies_list (list): List of tuples (name, value)
            - is_secure (bool): Whether connection uses HTTPS
    """
    results = []
    current_domain = None
    current_cookies = []
    current_is_secure = False
    
    lines = raw_input.strip().split('\n')
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Handle key-value pairs on separate lines (like :authority followed by the value)
        if line == ':authority' and i + 1 < len(lines):
            # Save previous domain's cookies if any
            if current_domain and current_cookies:
                results.append((current_domain, current_cookies, current_is_secure))
                current_cookies = []
                current_is_secure = False
            
            current_domain = lines[i + 1].strip()
            if current_domain and not current_domain.startswith('.'):
                current_domain = '.' + current_domain
            i += 2
            continue
        
        # Handle traditional "Host: domain" format
        if line.lower().startswith('host:'):
            # Save previous domain's cookies if any
            if current_domain and current_cookies:
                results.append((current_domain, current_cookies, current_is_secure))
                current_cookies = []
                current_is_secure = False
            
            current_domain = line.split(':', 1)[1].strip()
            if current_domain and not current_domain.startswith('.'):
                current_domain = '.' + current_domain
        
        # Handle :scheme or scheme value
        elif line == ':scheme' and i + 1 < len(lines):
            scheme_value = lines[i + 1].strip().lower()
            if scheme_value == 'https':
                current_is_secure = True
            i += 2
            continue
        elif line.lower().startswith(':scheme:'):
            if 'https' in line.lower():
                current_is_secure = True
        
        # Handle cookie on separate line or inline
        elif line == 'cookie' and i + 1 < len(lines):
            cookie_line = lines[i + 1].strip()
            current_cookies.extend(parse_cookie_string(cookie_line))
            i += 2
            continue
        elif line.lower().startswith('cookie:'):
            cookie_line = line.split(':', 1)[1].strip()
            current_cookies.extend(parse_cookie_string(cookie_line))
        
        i += 1
    
    # Don't forget the last domain's cookies
    if current_domain and current_cookies:
        results.append((current_domain, current_cookies, current_is_secure))
    
    # Fallback: if no headers found, treat entire input as cookie string
    if not results:
        cookies = parse_cookie_string(raw_input)
        if cookies:
            results.append((None, cookies, False))
    
    return results


def parse_cookie_string(cookie_string):
    """
    Parse a cookie string into name-value pairs.
    
    Args:
        cookie_string (str): Cookie string (semicolon-separated)
    
    Returns:
        list: List of tuples (name, value)
    """
    cookies = []
    
    # Split by newlines first to handle multiple cookie lines
    lines = cookie_string.strip().split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Remove "cookie:" prefix if present
        if line.lower().startswith('cookie:'):
            line = line.split(':', 1)[1].strip()
            
        # Split by semicolons to get individual cookies
        cookie_parts = line.split(';')
        
        for part in cookie_parts:
            part = part.strip()
            if not part or '=' not in part:
                continue
                
            # Split only on first = to handle values with =
            name, value = part.split('=', 1)
            name = name.strip()
            value = value.strip()
            
            if name and value:
                cookies.append((name, value))
    
    return cookies


def convert_to_netscape_format(domain_cookies_list):
    """
    Convert parsed cookies from multiple domains to Netscape format.
    
    Args:
        domain_cookies_list (list): List of tuples (domain, cookies, is_secure)
            - domain (str): Domain for the cookies
            - cookies (list): List of tuples (name, value)
            - is_secure (bool): Whether cookies were sent over HTTPS
    
    Returns:
        str: Cookies in Netscape format
    """
    netscape_cookies = []
    
    # Add Netscape header
    netscape_cookies.append("# Netscape HTTP Cookie File")
    netscape_cookies.append("# This file was generated by cookie_convert.py")
    
    # Add domain info
    domains = [domain for domain, _, _ in domain_cookies_list if domain]
    if domains:
        netscape_cookies.append(f"# Domains: {', '.join(domains)}")
    
    netscape_cookies.append("")
    
    # Process each domain's cookies
    for domain, cookies, is_secure in domain_cookies_list:
        for name, value in cookies:
            # Netscape format: domain, flag, path, secure, expiration, name, value
            flag = 'TRUE'  # Include subdomains
            path = '/'
            secure = 'TRUE' if is_secure else 'FALSE'
            expiration = '0'  # Session cookie (expires at end of session)
            
            netscape_cookie = f"{domain}\t{flag}\t{path}\t{secure}\t{expiration}\t{name}\t{value}"
            netscape_cookies.append(netscape_cookie)
    
    return '\n'.join(netscape_cookies)


def main():
    """Main function to convert cookies from raw_cookie.txt to cookies.txt"""
    
    print("=" * 70)
    print("Cookie Converter - Raw HTTP Headers → Netscape Format")
    print("=" * 70)
    print()
    
    # Read the raw input from the file
    try:
        with open("raw_cookie.txt", "r", encoding='utf-8') as file:
            raw_input = file.read()
    except FileNotFoundError:
        print("❌ Error: raw_cookie.txt not found!")
        print()
        print("Please create raw_cookie.txt and paste your request headers:")
        print("  1. Open browser Developer Tools (F12)")
        print("  2. Go to Network tab")
        print("  3. Click on any request")
        print("  4. Copy the RAW request headers")
        print("  5. Paste into raw_cookie.txt")
        return
    
    if not raw_input.strip():
        print("❌ Error: raw_cookie.txt is empty!")
        print()
        print("Please paste into raw_cookie.txt:")
        print("  OPTION 1 (Recommended): Entire raw request headers from Network tab")
        print("  OPTION 2: Just the cookie string (name=value; name2=value2; ...)")
        return
    
    # Parse headers and extract domain and cookies
    print("📖 Parsing raw_cookie.txt...")
    domain_cookies_list = parse_raw_headers(raw_input)
    
    if not domain_cookies_list:
        print("❌ No valid cookies found in raw_cookie.txt!")
        print()
        print("Make sure you copied either:")
        print("  - The entire raw request headers from Network tab")
        print("  - Or the Cookie header value (name=value; name2=value2; ...)")
        return
    
    # Count total cookies and domains
    total_cookies = sum(len(cookies) for _, cookies, _ in domain_cookies_list)
    domains_found = [domain for domain, _, _ in domain_cookies_list if domain]
    
    print(f"✓ Found {total_cookies} cookies from {len(domain_cookies_list)} domain(s)")
    
    for domain, cookies, is_secure in domain_cookies_list:
        if domain:
            print(f"  • {domain}: {len(cookies)} cookies (HTTPS: {is_secure})")
        else:
            print(f"  • Unknown domain: {len(cookies)} cookies")
    
    print()
    
    # Handle domains that need confirmation or input
    final_list = []
    for domain, cookies, is_secure in domain_cookies_list:
        if domain:
            # Domain was auto-detected, just confirm
            confirm = input(f"Use domain '{domain}' for its {len(cookies)} cookies? (Y/n): ").strip().lower()
            if confirm and confirm not in ['y', 'yes', '']:
                domain = input("Enter correct domain: ").strip()
                if domain and not domain.startswith('.'):
                    domain = '.' + domain
        else:
            # No domain detected, ask for it
            print(f"⚠ Could not auto-detect domain for {len(cookies)} cookies")
            domain = input("Enter domain (e.g., .google.com): ").strip()
            if not domain:
                domain = '.example.com'
                print(f"Using default: {domain}")
            elif not domain.startswith('.'):
                domain = '.' + domain
        
        final_list.append((domain, cookies, is_secure))
    
    print()
    print("🔄 Converting to Netscape format...")
    
    # Convert to Netscape format
    netscape_cookies = convert_to_netscape_format(final_list)
    
    # Write the Netscape format cookies to a file
    with open("cookies.txt", "w", encoding='utf-8') as file:
        file.write(netscape_cookies)
    
    print()
    print("=" * 70)
    print("✅ SUCCESS!")
    print("=" * 70)
    print(f"✓ Converted {total_cookies} cookies to Netscape format")
    for domain, cookies, is_secure in final_list:
        print(f"  • {domain}: {len(cookies)} cookies (HTTPS: {is_secure})")
    print(f"✓ Output: cookies.txt")
    print()
    print("You can now use cookies.txt with tools like yt-dlp:")
    print(f"  yt-dlp --cookies cookies.txt <URL>")
    print()


if __name__ == "__main__":
    main()
