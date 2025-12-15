"""
Outlook Searcher - COM interface for reading Outlook emails
WARNING: This module is READ-ONLY. It will never modify any Outlook data.
"""

import win32com.client
from datetime import datetime
from typing import List, Dict, Optional, Tuple


class OutlookSearcher:
    """
    Handles all Outlook COM interactions for searching emails.
    This class is strictly READ-ONLY and will never modify Outlook data.
    """
    
    def __init__(self, log_callback=None):
        self.outlook = None
        self.namespace = None
        self.log_callback = log_callback
        self.mail_folders = []  # Cache of actual mail folders
    
    def _log(self, message, level='INFO'):
        """Log a message either to callback or console."""
        if self.log_callback:
            self.log_callback(message, level)
        else:
            print(f"[{level}] {message}")
        
    def connect_to_outlook(self) -> Tuple[bool, str]:
        """
        Establish connection to Outlook via COM and cache mail folders.
        
        Returns:
            Tuple[bool, str]: (success, message)
        """
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            # Cache mail folders on startup
            self._cache_mail_folders()
            
            return True, "Connected to Outlook successfully"
        except Exception as e:
            return False, f"Failed to connect to Outlook: {str(e)}"
    
    def _cache_mail_folders(self):
        """
        Enumerate and cache all actual mail folders (skip system folders).
        """
        self.mail_folders = []
        
        # Folders to skip (not mail folders)
        skip_folders = {'Calendar', 'Contacts', 'Tasks', 'Notes', 'Journal',
                       'Sync Issues', 'RSS Feeds', 'Yammer Root', 'Files',
                       'Conversation Action Settings', 'Quick Step Settings',
                       'Social Activity Notifications', 'ExternalContacts',
                       'Birthdays', 'Organizational Contacts',
                       'GAL Contacts', 'Companies', 'Recipient Cache',
                       'PeopleCentricConversation Buddies'}
        
        try:
            for store in self.namespace.Stores:
                # Skip Public Folders
                if "Public Folders" in store.DisplayName:
                    continue
                
                try:
                    root_folder = store.GetRootFolder()
                    all_folders = self._get_all_folders_recursive(root_folder)
                    
                    # Filter to only mail folders
                    for folder in all_folders:
                        try:
                            # Skip system folders
                            if folder.Name in skip_folders or folder.Name.startswith('{'):
                                continue
                            
                            # Test if it's a mail folder (has items with ReceivedTime)
                            items = folder.Items
                            items.Sort("[ReceivedTime]", True)
                            
                            # It's a mail folder - add to cache
                            self.mail_folders.append(folder)
                        except:
                            # Not a mail folder or can't access
                            pass
                except:
                    pass
                    
        except Exception as e:
            self._log(f"Error caching mail folders: {str(e)}", "WARNING")
    
    def get_folders(self) -> List[str]:
        """
        Get list of available Outlook folders.
        
        Returns:
            List[str]: List of folder names (simplified for performance)
        """
        if not self.namespace:
            return []
        
        try:
            # Return simplified list - full search will still work with "All Folders"
            return [
                "All Folders (including subfolders)",
                "---",
                "Inbox",
                "Sent Items", 
                "---",
                "01 Projects",
                "02 Office",
                "99 Misc",
                "Archive"
            ]
        except Exception as e:
            print(f"Error getting folders: {e}")
            return ["All Folders (including subfolders)", "Inbox", "Sent Items"]
    
    def _build_folder_list(self, folder, folder_list, prefix="", level=0, max_folders=100):
        """
        Recursively build folder list for dropdown with indentation.
        
        Args:
            folder: Current folder
            folder_list: List to append to
            prefix: Prefix for display
            level: Nesting level
            max_folders: Maximum folders to add (prevent freezing)
        """
        # Stop if we have too many folders
        if len(folder_list) > max_folders:
            return
            
        try:
            # Skip certain system folders
            folder_name = folder.Name
            skip_folders = ["Sync Issues", "Recoverable Items", "Calendar", 
                          "Contacts", "Tasks", "Notes", "Journal", "Deleted Items",
                          "Junk Email", "Drafts", "Outbox", "RSS Feeds", "Conversation History"]
            
            if folder_name not in skip_folders and level < 5:  # Limit depth
                indent = "  " * level
                display_name = f"{indent}{folder_name}"
                
                # Avoid duplicates
                if display_name not in folder_list:
                    folder_list.append(display_name)
                
                # Recursively add subfolders
                try:
                    if hasattr(folder, 'Folders') and folder.Folders.Count > 0:
                        for subfolder in folder.Folders:
                            self._build_folder_list(subfolder, folder_list, prefix, level + 1, max_folders)
                except:
                    pass
        except Exception as e:
            pass
    
    def _sanitize_dasl_input(self, text: str) -> str:
        """
        Escape single quotes for DASL queries.
        
        Args:
            text: Input text
            
        Returns:
            Sanitized text
        """
        if not text:
            return ""
        return str(text).replace("'", "''")
    
    def _build_date_filter(self, criteria: Dict) -> str:
        """
        Build date-only filter for Items.Restrict(). 
        Only use filters that work reliably across all Outlook versions.
        Text matching (sender, subject) will be done manually for reliability.
        
        Args:
            criteria: Search criteria dictionary
            
        Returns:
            str: Date filter string or empty string
        """
        filters = []
        
        # Only use date filters - these work reliably in Restrict()
        date_from = criteria.get('date_from', '').strip()
        date_to = criteria.get('date_to', '').strip()
        
        if date_from:
            try:
                from datetime import datetime
                dt = datetime.strptime(date_from, '%Y-%m-%d')
                formatted_date = dt.strftime('%m/%d/%Y 12:00 AM')
                filters.append(f"[ReceivedTime] >= '{formatted_date}'")
            except:
                pass
        
        if date_to:
            try:
                from datetime import datetime
                dt = datetime.strptime(date_to, '%Y-%m-%d')
                formatted_date = dt.strftime('%m/%d/%Y 11:59 PM')
                filters.append(f"[ReceivedTime] <= '{formatted_date}'")
            except:
                pass
        
        if filters:
            return ' AND '.join(filters)
        
        return ''
    
    def _get_all_folders_recursive(self, folder, folder_list=None, level=0):
        """
        Recursively get all subfolders.
        
        Args:
            folder: Parent folder
            folder_list: List to accumulate folders
            level: Recursion depth for logging
            
        Returns:
            List of all folders including subfolders
        """
        if folder_list is None:
            folder_list = []
        
        try:
            # Add this folder to the list
            folder_list.append(folder)
            indent = "  " * level
            self._log(f"{indent}+ {folder.Name}", "DEBUG")
            
            # Try to get subfolders
            try:
                if hasattr(folder, 'Folders'):
                    subfolder_count = folder.Folders.Count
                    if subfolder_count > 0:
                        # Recursively process subfolders
                        for subfolder in folder.Folders:
                            self._get_all_folders_recursive(subfolder, folder_list, level + 1)
            except Exception as e:
                # No subfolders or access denied
                self._log(f"{indent}  (no subfolders or access denied)", "DEBUG")
                
        except Exception as e:
            # Some folders may not be accessible
            self._log(f"Cannot access folder: {str(e)}", "WARNING")
        
        return folder_list
    
    def _get_folder_by_name(self, folder_name: str):
        """
        Get folder object by name.
        
        Args:
            folder_name: Name of the folder (may include indentation from dropdown)
            
        Returns:
            Folder object or None
        """
        try:
            # Remove indentation from folder name
            clean_name = folder_name.strip()
            
            if clean_name == "All Folders (including subfolders)":
                return None  # Will search all recursively
            elif clean_name == "---":
                return None
            else:
                # Search for folder by name recursively
                return self._find_folder_by_name(clean_name)
        except Exception as e:
            print(f"Error getting folder: {e}")
            return None
    
    def _find_folder_by_name(self, name: str):
        """
        Find folder by name in all stores.
        
        Args:
            name: Folder name to find
            
        Returns:
            Folder object or None
        """
        try:
            for store in self.namespace.Stores:
                try:
                    root_folder = store.GetRootFolder()
                    result = self._search_folder_recursive(root_folder, name)
                    if result:
                        return result
                except:
                    pass
        except:
            pass
        
        # Fallback to default folders
        if name == "Inbox":
            return self.namespace.GetDefaultFolder(6)
        elif name == "Sent Items":
            return self.namespace.GetDefaultFolder(5)
        
        return None
    
    def _search_folder_recursive(self, folder, name):
        """
        Recursively search for folder by name.
        
        Args:
            folder: Current folder to search
            name: Name to match
            
        Returns:
            Folder object or None
        """
        try:
            if folder.Name == name:
                return folder
            
            if hasattr(folder, 'Folders'):
                for subfolder in folder.Folders:
                    result = self._search_folder_recursive(subfolder, name)
                    if result:
                        return result
        except:
            pass
        
        return None
    
    def search_emails(self, criteria: Dict) -> List[Dict]:
        """
        Search emails based on criteria using Outlook's Restrict method. READ-ONLY operation.
        
        Args:
            criteria: Dictionary with search parameters
                - from_text: str - partial match for sender
                - search_terms: str - terms to search for (AND logic)
                - search_subject: bool - search in subject
                - search_body: bool - search in body
                - date_from: str - start date (YYYY-MM-DD)
                - date_to: str - end date (YYYY-MM-DD)
                - folder: str - folder to search in
        
        Returns:
            List[Dict]: List of matching emails with metadata
        """
        if not self.namespace:
            self._log("Not connected to Outlook namespace", "ERROR")
            return []
        
        results = []
        
        try:
            # Get folder to search
            folder_name = criteria.get('folder', 'Inbox')
            self._log(f"Searching in folder: {folder_name}", "INFO")
            
            if folder_name == "All Folders (including subfolders)":
                # Use cached mail folders
                folders_to_search = self.mail_folders
                self._log(f"Using {len(folders_to_search)} cached mail folders", "INFO")
                    
            elif folder_name == "---":
                return []
            else:
                # Search specific folder
                folder = self._find_folder_by_name(folder_name.strip())
                if folder:
                    folders_to_search = [folder]
                    self._log(f"Found folder: {folder.Name}", "INFO")
                else:
                    self._log(f"Folder '{folder_name}' not found", "ERROR")
                    return []
            
            # Hybrid approach: date Restrict (reliable) + manual text filtering
            import time
            search_start = time.time()
            
            date_filter = self._build_date_filter(criteria)
            
            # Extract text search criteria
            from_text = criteria.get('from_text', '').strip().lower()
            search_terms = criteria.get('search_terms', '').strip().lower()
            search_subject = criteria.get('search_subject', True)
            search_body = criteria.get('search_body', False)
            
            # Log search plan
            filter_parts = []
            if date_filter:
                filter_parts.append(f"dates: {date_filter}")
            if from_text:
                filter_parts.append(f"sender: '{from_text}'")
            if search_terms:
                where = []
                if search_subject:
                    where.append("subject")
                if search_body:
                    where.append("body")
                filter_parts.append(f"'{search_terms}' in {' or '.join(where)}")
            
            if filter_parts:
                self._log(f"Filters: {'; '.join(filter_parts)}", "DEBUG")
            else:
                self._log("No filters - returning recent emails", "INFO")
            
            self._log(f"Searching {len(folders_to_search)} folder(s)...", "INFO")
            
            # Search each folder
            total_folders_searched = 0
            
            for search_folder in folders_to_search:
                try:
                    folder_start = time.time()
                    items = search_folder.Items
                    
                    # Sort by received time FIRST (before filtering) for faster access
                    try:
                        items.Sort("[ReceivedTime]", True)
                    except:
                        pass
                    
                    count_in_folder = 0
                    checked = 0
                    
                    # Use optimized iteration - only check recent emails if no date filter
                    max_to_check = 1000 if not date_filter else 999999
                    
                    # Manual filtering for sender/subject/body
                    for item in items:
                        checked += 1
                        if checked > max_to_check:
                            break
                        
                        # Skip non-mail items
                        if not hasattr(item, 'SenderEmailAddress'):
                            continue
                        
                        # Apply date filter manually if specified
                        if date_filter:
                            try:
                                received_time = getattr(item, 'ReceivedTime', None)
                                if received_time:
                                    # Check against date criteria
                                    date_from_str = criteria.get('date_from', '').strip()
                                    date_to_str = criteria.get('date_to', '').strip()
                                    
                                    if date_from_str:
                                        from datetime import datetime
                                        date_from_obj = datetime.strptime(date_from_str, '%Y-%m-%d')
                                        if received_time.date() < date_from_obj.date():
                                            continue
                                    
                                    if date_to_str:
                                        from datetime import datetime
                                        date_to_obj = datetime.strptime(date_to_str, '%Y-%m-%d')
                                        if received_time.date() > date_to_obj.date():
                                            continue
                            except:
                                pass
                        
                        # Check sender if specified
                        if from_text:
                            sender_name = getattr(item, 'SenderName', '').lower()
                            sender_email = getattr(item, 'SenderEmailAddress', '').lower()
                            
                            # Try to get SMTP address
                            smtp_address = ''
                            try:
                                sender = getattr(item, 'Sender', None)
                                if sender:
                                    smtp_address = sender.PropertyAccessor.GetProperty(
                                        "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                                    ).lower()
                            except:
                                pass
                            
                            # Check if any sender field contains the search text
                            if not (from_text in sender_name or from_text in sender_email or from_text in smtp_address):
                                continue
                        
                        # Check subject/body if specified
                        if search_terms:
                            found = False
                            
                            # Check subject
                            if search_subject:
                                subject = getattr(item, 'Subject', '').lower()
                                if search_terms in subject:
                                    found = True
                            
                            # Check body if not found in subject (or if subject search disabled)
                            if not found and search_body:
                                try:
                                    body_text = getattr(item, 'Body', '').lower()
                                    if search_terms in body_text:
                                        found = True
                                except:
                                    pass
                            
                            # Skip if not found anywhere
                            if not found:
                                continue
                        
                        # Extract email data
                        email_data = self._extract_email_data(item)
                        if email_data:
                            email_data['folder_name'] = search_folder.Name
                            results.append(email_data)
                            count_in_folder += 1
                        
                        # Limit results
                        if len(results) >= 500:
                            self._log("Reached 500 result limit", "WARNING")
                            break
                    
                    folder_time = time.time() - folder_start
                    
                    if count_in_folder > 0:
                        self._log(f"✓ '{search_folder.Name}': {count_in_folder} emails ({folder_time:.2f}s)", "SUCCESS")
                    
                    total_folders_searched += 1
                    
                    if len(results) >= 500:
                        break
                        
                except Exception as e:
                    # Skip problematic folders
                    self._log(f"Error in folder '{search_folder.Name}': {str(e)}", "DEBUG")
            
            search_time = time.time() - search_start
            self._log(f"🚀 Search complete in {search_time:.2f} seconds! Found {len(results)} emails from {total_folders_searched} folders", "SUCCESS")
                    
        except Exception as e:
            self._log(f"Search error: {str(e)}", "ERROR")
            import traceback
            traceback.print_exc()
        
        return results
    
    def _matches_criteria(self, item, criteria: Dict) -> bool:
        """
        Check if email matches search criteria.
        
        Args:
            item: Outlook mail item
            criteria: Search criteria dictionary
            
        Returns:
            bool: True if matches
        """
        try:
            # From field filter
            from_text = criteria.get('from_text', '').strip().lower()
            if from_text:
                sender_email = getattr(item, 'SenderEmailAddress', '').lower()
                sender_name = getattr(item, 'SenderName', '').lower()
                
                # Also try to get SMTP address from Sender property
                smtp_address = ''
                try:
                    sender = item.Sender
                    if sender:
                        # Try to get the SMTP address
                        smtp_address = getattr(sender, 'Address', '').lower()
                        if not smtp_address or 'exchangelabs' in smtp_address:
                            # Try PropertyAccessor for SMTP address
                            try:
                                smtp_address = sender.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E").lower()
                            except:
                                pass
                except:
                    pass
                
                # Debug: Log the first few comparisons
                if not hasattr(self, '_debug_count'):
                    self._debug_count = 0
                
                if self._debug_count < 3:
                    self._log(f"Checking: from_text='{from_text}'", "DEBUG")
                    self._log(f"  sender_name='{sender_name}'", "DEBUG")
                    self._log(f"  sender_email='{sender_email}'", "DEBUG")
                    self._log(f"  smtp_address='{smtp_address}'", "DEBUG")
                    self._debug_count += 1
                
                # Check all possible email fields
                if (from_text not in sender_email and 
                    from_text not in sender_name and 
                    from_text not in smtp_address):
                    return False
                else:
                    # Found a match!
                    if self._debug_count < 10:
                        self._log(f"MATCH! '{from_text}' found in sender (name='{sender_name}', email='{sender_email}', smtp='{smtp_address}')", "DEBUG")
            
            # Search terms filter (AND logic)
            search_terms = criteria.get('search_terms', '').strip().lower()
            if search_terms:
                terms = search_terms.split()
                search_subject = criteria.get('search_subject', True)
                search_body = criteria.get('search_body', True)
                
                subject = getattr(item, 'Subject', '').lower()
                body = getattr(item, 'Body', '').lower() if search_body else ''
                
                search_text = ''
                if search_subject:
                    search_text += subject + ' '
                if search_body:
                    search_text += body
                
                # All terms must be present (AND logic)
                for term in terms:
                    if term not in search_text:
                        return False
            
            # Date range filter
            date_from = criteria.get('date_from', '')
            date_to = criteria.get('date_to', '')
            
            if date_from or date_to:
                try:
                    received_time = item.ReceivedTime
                    # Convert to datetime if it's a PyTime object
                    if hasattr(received_time, 'date'):
                        received_date = received_time.date()
                    else:
                        received_date = datetime.strptime(str(received_time).split()[0], '%Y-%m-%d').date()
                    
                    if date_from:
                        from_date = datetime.strptime(date_from, '%Y-%m-%d').date()
                        if received_date < from_date:
                            return False
                    
                    if date_to:
                        to_date = datetime.strptime(date_to, '%Y-%m-%d').date()
                        if received_date > to_date:
                            return False
                except Exception as e:
                    print(f"Error parsing date: {e}")
                    pass
            
            return True
            
        except Exception as e:
            print(f"Error matching criteria: {e}")
            return False
    
    def _extract_email_data(self, item) -> Optional[Dict]:
        """
        Extract data from email item (READ-ONLY).
        
        Args:
            item: Outlook mail item
            
        Returns:
            Dict: Email metadata
        """
        try:
            # Get attachment info (READ-ONLY)
            attachment_count = 0
            attachment_names = []
            try:
                attachments = item.Attachments
                for attachment in attachments:
                    # Skip embedded images/hidden attachments
                    if hasattr(attachment, 'FileName'):
                        attachment_count += 1
                        attachment_names.append(attachment.FileName)
            except:
                pass
            
            # Format received time
            received_time = item.ReceivedTime
            if hasattr(received_time, 'strftime'):
                date_str = received_time.strftime('%Y-%m-%d %H:%M')
            else:
                date_str = str(received_time)
            
            # Get proper SMTP email address
            sender_email = getattr(item, 'SenderEmailAddress', '')
            # Try to get SMTP address if we have Exchange format
            if not sender_email or 'exchangelabs' in sender_email.lower() or '/o=' in sender_email.lower():
                try:
                    sender = item.Sender
                    if sender:
                        # Try PropertyAccessor for SMTP address
                        try:
                            sender_email = sender.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                        except:
                            sender_email = getattr(sender, 'Address', sender_email)
                except:
                    pass
            
            return {
                'date': date_str,
                'from': getattr(item, 'SenderName', 'Unknown'),
                'from_email': sender_email,
                'subject': getattr(item, 'Subject', '(No Subject)'),
                'attachments': f"📎 {attachment_count}" if attachment_count > 0 else "",
                'attachment_count': attachment_count,
                'attachment_names': ', '.join(attachment_names) if attachment_names else '',
                'entry_id': item.EntryID,  # For opening email later
            }
        except Exception as e:
            print(f"Error extracting email data: {e}")
            return None
    
    def open_email(self, entry_id: str) -> Tuple[bool, str]:
        """
        Open email in Outlook (READ-ONLY - just displays it).
        
        Args:
            entry_id: Email's EntryID
            
        Returns:
            Tuple[bool, str]: (success, message)
        """
        if not self.namespace:
            return False, "Not connected to Outlook"
        
        try:
            # Get the email item by EntryID
            # Use the Outlook application's Session instead of namespace
            item = self.outlook.Session.GetItemFromID(entry_id)
            # Display the email (READ-ONLY operation)
            item.Display()
            return True, "Email opened successfully"
        except Exception as e:
            return False, f"Failed to open email: {str(e)}"
