"""
Test script to enumerate all Outlook folders
"""

import win32com.client

def enumerate_folders(folder, level=0):
    """Recursively enumerate folders"""
    indent = "  " * level
    try:
        print(f"{indent}📁 {folder.Name} ({folder.Folders.Count} subfolders)")
        
        # Try to enumerate subfolders
        try:
            for subfolder in folder.Folders:
                enumerate_folders(subfolder, level + 1)
        except Exception as e:
            print(f"{indent}  ⚠️ Cannot access subfolders: {e}")
    except Exception as e:
        print(f"{indent}❌ Error: {e}")

def main():
    print("=" * 60)
    print("OUTLOOK FOLDER ENUMERATION TEST")
    print("=" * 60)
    
    try:
        # Connect to Outlook
        print("\n1. Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("✓ Connected successfully\n")
        
        # Test 1: Get default folders
        print("2. Testing default folders:")
        print("-" * 60)
        try:
            inbox = namespace.GetDefaultFolder(6)
            print(f"✓ Inbox: {inbox.Name}")
            
            sent = namespace.GetDefaultFolder(5)
            print(f"✓ Sent Items: {sent.Name}")
        except Exception as e:
            print(f"❌ Error getting default folders: {e}")
        
        # Test 2: Enumerate all stores
        print("\n3. Enumerating all stores:")
        print("-" * 60)
        try:
            store_count = namespace.Stores.Count
            print(f"Found {store_count} store(s)\n")
            
            for i, store in enumerate(namespace.Stores, 1):
                print(f"\nStore {i}: {store.DisplayName}")
                try:
                    root = store.GetRootFolder()
                    print(f"  Root folder: {root.Name}")
                    print(f"  Subfolders:")
                    enumerate_folders(root, level=2)
                except Exception as e:
                    print(f"  ❌ Cannot access root: {e}")
                    
        except Exception as e:
            print(f"❌ Error enumerating stores: {e}")
        
        # Test 3: Enumerate Inbox subfolders
        print("\n4. Detailed Inbox subfolder enumeration:")
        print("-" * 60)
        try:
            inbox = namespace.GetDefaultFolder(6)
            enumerate_folders(inbox, level=0)
        except Exception as e:
            print(f"❌ Error: {e}")
        
        # Test 4: Enumerate Sent Items subfolders
        print("\n5. Detailed Sent Items subfolder enumeration:")
        print("-" * 60)
        try:
            sent = namespace.GetDefaultFolder(5)
            enumerate_folders(sent, level=0)
        except Exception as e:
            print(f"❌ Error: {e}")
            
        print("\n" + "=" * 60)
        print("TEST COMPLETE")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n❌ FATAL ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
    input("\nPress Enter to exit...")
