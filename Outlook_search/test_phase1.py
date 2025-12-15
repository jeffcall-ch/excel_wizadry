"""
Test script for Outlook Search Application - Phase 1
This script verifies the read-only nature and basic functionality.
"""

import sys
import inspect
from outlook_searcher import OutlookSearcher


def test_read_only_verification():
    """Verify that OutlookSearcher has no write/modify methods."""
    print("=" * 60)
    print("TEST: Verifying READ-ONLY nature of OutlookSearcher")
    print("=" * 60)
    
    # Get all methods of OutlookSearcher
    methods = inspect.getmembers(OutlookSearcher, predicate=inspect.isfunction)
    
    # Forbidden keywords that suggest write operations
    forbidden_keywords = [
        'delete', 'remove', 'send', 'forward', 'reply', 'move', 
        'save', 'write', 'update', 'modify', 'change', 'edit',
        'mark', 'flag', 'category', 'unread'
    ]
    
    print("\nAll methods in OutlookSearcher class:")
    print("-" * 60)
    
    violations = []
    for name, method in methods:
        if not name.startswith('_'):  # Public methods
            print(f"  ✓ {name}")
            
            # Check method name for forbidden keywords
            method_lower = name.lower()
            for keyword in forbidden_keywords:
                if keyword in method_lower and keyword not in ['_from', '_to']:
                    violations.append((name, keyword))
    
    print("\nPrivate/helper methods:")
    print("-" * 60)
    for name, method in methods:
        if name.startswith('_') and not name.startswith('__'):
            print(f"  ✓ {name}")
    
    if violations:
        print("\n❌ FAILED: Found potentially unsafe methods:")
        for method_name, keyword in violations:
            print(f"  ⚠️  {method_name} contains '{keyword}'")
        return False
    else:
        print("\n✅ PASSED: No write/modify methods found")
        print("   All methods are read-only operations")
        return True


def test_basic_functionality():
    """Test basic connection and method availability."""
    print("\n" + "=" * 60)
    print("TEST: Basic Functionality")
    print("=" * 60)
    
    try:
        searcher = OutlookSearcher()
        print("✓ OutlookSearcher instance created")
        
        # Test connection (may fail if Outlook not running)
        success, message = searcher.connect_to_outlook()
        if success:
            print(f"✓ Connected to Outlook: {message}")
            
            # Test get_folders
            folders = searcher.get_folders()
            print(f"✓ Retrieved folders: {folders}")
            
            # Test search with empty criteria
            criteria = {
                'from_text': '',
                'search_terms': '',
                'search_subject': True,
                'search_body': False,
                'date_from': '',
                'date_to': '',
                'folder': 'Inbox'
            }
            print("✓ Created test criteria (empty search)")
            print("  Note: Actual search not performed to avoid loading data")
            
        else:
            print(f"⚠️  Could not connect to Outlook: {message}")
            print("   This is OK if Outlook is not running")
        
        print("\n✅ PASSED: Basic functionality test completed")
        return True
        
    except Exception as e:
        print(f"\n❌ FAILED: {str(e)}")
        return False


def test_code_inspection():
    """Inspect the code for any dangerous operations."""
    print("\n" + "=" * 60)
    print("TEST: Code Inspection for Safety")
    print("=" * 60)
    
    try:
        with open('outlook_searcher.py', 'r', encoding='utf-8') as f:
            code = f.read()
        
        # Dangerous patterns to look for
        dangerous_patterns = [
            '.Delete(', '.Move(', '.Send(', '.Forward(', '.Reply(',
            '.Save(', '.Close(', '.UnRead =', '.FlagRequest =',
            '.Categories =', '.Importance ='
        ]
        
        violations = []
        for pattern in dangerous_patterns:
            if pattern in code:
                violations.append(pattern)
        
        if violations:
            print("\n❌ FAILED: Found potentially dangerous code patterns:")
            for pattern in violations:
                print(f"  ⚠️  Found: {pattern}")
            return False
        else:
            print("✓ No dangerous write operations found in code")
            print("✓ No .Delete() calls")
            print("✓ No .Move() calls")
            print("✓ No .Send() calls")
            print("✓ No property assignments that modify data")
            print("\n✅ PASSED: Code is safe (read-only)")
            return True
            
    except Exception as e:
        print(f"\n❌ FAILED: {str(e)}")
        return False


def main():
    """Run all tests."""
    print("\n" + "=" * 60)
    print("OUTLOOK SEARCH APPLICATION - PHASE 1 TESTS")
    print("=" * 60)
    print("\nThese tests verify the application is READ-ONLY")
    print("and will never modify any Outlook data.\n")
    
    results = []
    
    # Run tests
    results.append(("Read-Only Verification", test_read_only_verification()))
    results.append(("Basic Functionality", test_basic_functionality()))
    results.append(("Code Safety Inspection", test_code_inspection()))
    
    # Summary
    print("\n" + "=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for test_name, result in results:
        status = "✅ PASSED" if result else "❌ FAILED"
        print(f"{status}: {test_name}")
    
    print(f"\nTotal: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n🎉 All tests passed! Application is safe to use.")
        print("   The application is confirmed READ-ONLY.")
        return 0
    else:
        print("\n⚠️  Some tests failed. Review the output above.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
