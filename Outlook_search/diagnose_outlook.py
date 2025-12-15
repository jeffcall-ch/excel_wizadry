"""
Diagnostic script to verify Outlook connection and data retrieval
"""

import sys
import win32com.client
from datetime import datetime


def test_outlook_connection():
    """Test basic Outlook connection."""
    print("=" * 70)
    print("TEST 1: Outlook Connection")
    print("=" * 70)
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("✅ Successfully created Outlook Application object")
        
        namespace = outlook.GetNamespace("MAPI")
        print("✅ Successfully got MAPI namespace")
        
        return outlook, namespace
    except Exception as e:
        print(f"❌ Failed to connect to Outlook: {e}")
        return None, None


def test_folder_access(namespace):
    """Test accessing Outlook folders."""
    print("\n" + "=" * 70)
    print("TEST 2: Folder Access")
    print("=" * 70)
    
    if not namespace:
        print("❌ No namespace available")
        return None, None
    
    try:
        inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
        print(f"✅ Successfully accessed Inbox")
        print(f"   Folder name: {inbox.Name}")
        print(f"   Total items: {inbox.Items.Count}")
        
        sent_items = namespace.GetDefaultFolder(5)  # olFolderSentMail
        print(f"✅ Successfully accessed Sent Items")
        print(f"   Folder name: {sent_items.Name}")
        print(f"   Total items: {sent_items.Items.Count}")
        
        return inbox, sent_items
    except Exception as e:
        print(f"❌ Failed to access folders: {e}")
        return None, None


def test_read_emails(folder, max_emails=5):
    """Test reading emails from a folder."""
    print("\n" + "=" * 70)
    print(f"TEST 3: Reading Emails from {folder.Name}")
    print("=" * 70)
    
    if not folder:
        print("❌ No folder available")
        return False
    
    try:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)  # Sort by newest first
        
        count = min(max_emails, items.Count)
        print(f"📧 Attempting to read {count} most recent emails...\n")
        
        emails_read = 0
        for i, item in enumerate(items):
            if i >= max_emails:
                break
            
            # Skip non-mail items
            if not hasattr(item, 'Subject'):
                continue
            
            try:
                print(f"Email #{i+1}:")
                print(f"  Date: {item.ReceivedTime}")
                print(f"  From: {getattr(item, 'SenderName', 'Unknown')}")
                print(f"  Email: {getattr(item, 'SenderEmailAddress', 'Unknown')}")
                print(f"  Subject: {getattr(item, 'Subject', '(No Subject)')[:60]}")
                
                # Check attachments
                try:
                    att_count = item.Attachments.Count
                    print(f"  Attachments: {att_count}")
                except:
                    print(f"  Attachments: 0")
                
                # Check body (first 100 chars)
                try:
                    body = getattr(item, 'Body', '')
                    body_preview = body[:100].replace('\n', ' ').replace('\r', '')
                    print(f"  Body preview: {body_preview}...")
                except:
                    print(f"  Body preview: (unable to read)")
                
                print()
                emails_read += 1
                
            except Exception as e:
                print(f"  ⚠️ Error reading email #{i+1}: {e}\n")
                continue
        
        if emails_read > 0:
            print(f"✅ Successfully read {emails_read} emails")
            return True
        else:
            print(f"⚠️ No emails could be read")
            return False
            
    except Exception as e:
        print(f"❌ Failed to read emails: {e}")
        return False


def test_search_functionality(namespace):
    """Test basic search functionality."""
    print("\n" + "=" * 70)
    print("TEST 4: Search Functionality")
    print("=" * 70)
    
    if not namespace:
        print("❌ No namespace available")
        return False
    
    try:
        inbox = namespace.GetDefaultFolder(6)
        items = inbox.Items
        
        print("Searching for emails with 'the' in subject or body...")
        print("(This is a common word, should find results)\n")
        
        found_count = 0
        for i, item in enumerate(items):
            if i >= 50:  # Check first 50 emails
                break
            
            if not hasattr(item, 'Subject'):
                continue
            
            subject = getattr(item, 'Subject', '').lower()
            body = getattr(item, 'Body', '').lower()
            
            if 'the' in subject or 'the' in body:
                found_count += 1
                if found_count <= 3:  # Show first 3 matches
                    print(f"Match #{found_count}:")
                    print(f"  Subject: {getattr(item, 'Subject', '(No Subject)')[:60]}")
                    print(f"  Date: {item.ReceivedTime}")
                    print()
        
        print(f"✅ Found {found_count} emails containing 'the' in first 50 emails")
        return True
        
    except Exception as e:
        print(f"❌ Search test failed: {e}")
        return False


def test_entry_id_retrieval(folder):
    """Test getting EntryID for opening emails later."""
    print("\n" + "=" * 70)
    print("TEST 5: EntryID Retrieval (for opening emails)")
    print("=" * 70)
    
    if not folder:
        print("❌ No folder available")
        return False
    
    try:
        items = folder.Items
        if items.Count == 0:
            print("⚠️ No emails in folder to test")
            return True
        
        item = items.GetFirst()
        if not hasattr(item, 'Subject'):
            print("⚠️ First item is not an email")
            return True
        
        entry_id = item.EntryID
        print(f"✅ Successfully retrieved EntryID")
        print(f"   EntryID length: {len(entry_id)} characters")
        print(f"   Sample: {entry_id[:50]}...")
        
        # Test retrieving item by EntryID
        # Get outlook application to access Session
        outlook_app = folder.Application
        retrieved_item = outlook_app.Session.GetItemFromID(entry_id)
        print(f"✅ Successfully retrieved email by EntryID")
        print(f"   Subject: {getattr(retrieved_item, 'Subject', '(No Subject)')[:60]}")
        
        return True
        
    except Exception as e:
        print(f"❌ EntryID test failed: {e}")
        return False


def main():
    """Run all diagnostic tests."""
    print("\n" + "=" * 70)
    print("OUTLOOK CONNECTION AND DATA RETRIEVAL DIAGNOSTIC")
    print("=" * 70)
    print("\nThis script will verify:")
    print("1. Connection to Outlook")
    print("2. Access to folders (Inbox, Sent Items)")
    print("3. Reading email data")
    print("4. Search functionality")
    print("5. EntryID retrieval for opening emails")
    print()
    
    results = []
    
    # Test 1: Connection
    outlook, namespace = test_outlook_connection()
    results.append(("Outlook Connection", outlook is not None))
    
    if not namespace:
        print("\n❌ Cannot continue without Outlook connection")
        print("\nPossible issues:")
        print("  - Outlook is not installed")
        print("  - Outlook is not configured with an email account")
        print("  - Outlook is not running (try opening it first)")
        print("  - Security settings blocking COM access")
        return 1
    
    # Test 2: Folder Access
    inbox, sent_items = test_folder_access(namespace)
    results.append(("Folder Access", inbox is not None))
    
    # Test 3: Read Emails
    if inbox and inbox.Items.Count > 0:
        success = test_read_emails(inbox, max_emails=3)
        results.append(("Read Emails", success))
    else:
        print("\n⚠️ Inbox is empty, skipping email read test")
        results.append(("Read Emails", None))
    
    # Test 4: Search
    search_success = test_search_functionality(namespace)
    results.append(("Search Functionality", search_success))
    
    # Test 5: EntryID
    if inbox:
        entry_id_success = test_entry_id_retrieval(inbox)
        results.append(("EntryID Retrieval", entry_id_success))
    
    # Summary
    print("\n" + "=" * 70)
    print("DIAGNOSTIC SUMMARY")
    print("=" * 70)
    
    for test_name, result in results:
        if result is True:
            status = "✅ PASSED"
        elif result is False:
            status = "❌ FAILED"
        else:
            status = "⚠️ SKIPPED"
        print(f"{status}: {test_name}")
    
    passed = sum(1 for _, result in results if result is True)
    failed = sum(1 for _, result in results if result is False)
    
    print(f"\nResults: {passed} passed, {failed} failed")
    
    if failed == 0 and passed > 0:
        print("\n🎉 All tests passed! Outlook connection and data retrieval working.")
        print("\nThe application should work correctly.")
        print("If the GUI search didn't work, please describe what happened:")
        print("  - Did you get an error message?")
        print("  - Did it show 'Found 0 emails'?")
        print("  - Did it freeze or crash?")
        return 0
    else:
        print("\n⚠️ Some tests failed. Review the output above.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
