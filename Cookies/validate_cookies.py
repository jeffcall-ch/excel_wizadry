import re

def validate_netscape_format(file_path):
    """
    Validates if the cookies in the given file adhere to the Netscape cookie format.

    Args:
        file_path (str): Path to the cookies.txt file.

    Returns:
        tuple: (bool, int, int) - (is_valid, total_cookies, errors)
    """
    try:
        with open(file_path, "r", encoding='utf-8') as file:
            lines = file.readlines()
    except FileNotFoundError:
        print(f"Error: {file_path} not found!")
        return False, 0, 0

    # Regular expression to match Netscape cookie format
    # Format: domain \t flag \t path \t secure \t expiration \t name \t value
    pattern = re.compile(r"^.+\t(TRUE|FALSE)\t.+\t(TRUE|FALSE)\t\d+\t.+\t.*$")

    valid = True
    cookie_count = 0
    error_count = 0
    
    for i, line in enumerate(lines, start=1):
        stripped_line = line.strip()
        
        # Skip empty lines and comments
        if not stripped_line or stripped_line.startswith('#'):
            continue
        
        cookie_count += 1
        
        if not pattern.match(stripped_line):
            print(f"❌ Line {i} does not match Netscape format:")
            print(f"   {stripped_line[:100]}{'...' if len(stripped_line) > 100 else ''}")
            
            # Try to provide helpful feedback
            parts = stripped_line.split('\t')
            if len(parts) != 7:
                print(f"   Expected 7 tab-separated fields, found {len(parts)}")
            
            valid = False
            error_count += 1

    return valid, cookie_count, error_count


def main():
    """Main function to validate cookies.txt"""
    file_path = "cookies.txt"
    
    print("Validating Netscape cookie format...")
    print("=" * 50)
    
    valid, cookie_count, error_count = validate_netscape_format(file_path)
    
    print("=" * 50)
    print(f"\nTotal cookies found: {cookie_count}")
    
    if valid:
        print("✓ Validation successful: All cookies are in valid Netscape format!")
        print(f"✓ {cookie_count} cookie(s) validated")
        return True
    else:
        print(f"✗ Validation failed: {error_count} cookie(s) have invalid format")
        print("\nNetscape format requires 7 tab-separated fields:")
        print("  domain \\t flag \\t path \\t secure \\t expiration \\t name \\t value")
        print("\nExample:")
        print("  .google.com\\tTRUE\\t/\\tFALSE\\t0\\tNID\\t123abc")
        return False


if __name__ == "__main__":
    main()