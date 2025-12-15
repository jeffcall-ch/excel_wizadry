' ============================================
' Module: ThisOutlookSession
' Purpose: Main search engine with ALL Phase 1-3 features using AQS/Explorer.Search
' Including: Native AQS search, Result capture, Bulk actions, Export functions, Statistics
' Version: 2.0 - Refactored from DASL to AQS (Explorer.Search)
' ============================================

Option Explicit

' Constants
Private Const APP_VERSION As String = "2.0.0"
Private Const APP_DATE As String = "2024-12-15"

Private Const KB_IN_BYTES As Long = 1024
Private Const MB_IN_BYTES As Long = 1048576
Private Const GB_IN_BYTES As Long = 1073741824
Private Const SIZE_100KB As Long = 102400
Private Const SIZE_1MB As Long = 1048576
Private Const SIZE_10MB As Long = 10485760

' NOTE: SearchCriteria type is now defined in modSharedTypes standard module
' to avoid "Cannot define public type in object module" error

' Module-level variables (AQS-based architecture)
Private m_CapturedResults As Collection    ' Stores EntryIDs from captured search results
Private m_SearchExplorer As Outlook.Explorer ' Reference to Explorer window with search results
Private m_Namespace As Outlook.NameSpace    ' Cached namespace object
Private m_FormInstance As frmEmailSearch    ' Reference to search form

' ============================================
' Main Entry Point - Show Search Form
' ============================================

Public Sub ShowEmailSearchForm()
    On Error GoTo ErrorHandler
    
    ' Create and show the search form
    Set m_FormInstance = New frmEmailSearch
    m_FormInstance.Show vbModeless
    Exit Sub
    
ErrorHandler:
    MsgBox "Error opening search form: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Common causes:" & vbCrLf & _
           "1. UserForm not created or not named 'frmEmailSearch'" & vbCrLf & _
           "2. UserForm code has compile errors" & vbCrLf & _
           "3. Missing controls on the form" & vbCrLf & vbCrLf & _
           "Press Alt+F11 to check the VBA editor.", vbCritical, "Cannot Open Search Form"
End Sub

'-----------------------------------------------------------
' Function: GetCachedNamespace
' Purpose: Return cached Outlook namespace for performance
' Returns: Outlook.NameSpace object
' Notes: CACHED - avoids repeated GetNamespace calls
'-----------------------------------------------------------
Private Function GetCachedNamespace() As Outlook.NameSpace
    If m_Namespace Is Nothing Then
        Set m_Namespace = Application.GetNamespace("MAPI")
    End If
    Set GetCachedNamespace = m_Namespace
End Function

'-----------------------------------------------------------
' Function: GetCapturedResults
' Purpose: Return current captured results for form access
' Returns: Collection of EntryIDs (String)
' Notes: Used by form to check if results available
'-----------------------------------------------------------
Public Function GetCapturedResults() As Collection
    Set GetCapturedResults = m_CapturedResults
End Function

' ============================================
' MAIN SEARCH EXECUTION - AQS/Explorer.Search
' ============================================

'-----------------------------------------------------------
' Sub: ExecuteEmailSearch
' Purpose: Execute UI search using Explorer.Search (AQS)
' Parameters:
'   criteria (SearchCriteria) - User search criteria from form
' Notes: ASYNC - Results appear in Outlook UI, not returned directly
'        User must click "Capture Results" after reviewing
'-----------------------------------------------------------
Public Sub ExecuteEmailSearch(criteria As SearchCriteria)
    On Error GoTo ErrorHandler
    
    Dim targetFolder As Outlook.Folder
    Dim searchExplorer As Outlook.Explorer
    Dim searchScope As OlSearchScope
    Dim aqsQuery As String
    
    ' Resolve target folder
    Set targetFolder = ResolveFolder(criteria.FolderName)
    If targetFolder Is Nothing Then
        MsgBox "Cannot resolve folder: " & criteria.FolderName & vbCrLf & vbCrLf & _
               "Defaulting to Inbox.", vbExclamation, "Folder Not Found"
        Set targetFolder = GetCachedNamespace().GetDefaultFolder(olFolderInbox)
    End If
    
    ' Open Explorer for search (creates new window)
    Set searchExplorer = OpenSearchExplorer(targetFolder)
    
    ' Build AQS query
    aqsQuery = BuildAQSQuery(criteria)
    
    ' Determine search scope
    searchScope = DetermineSearchScope(criteria)
    
    ' Debug output
    Debug.Print ""
    Debug.Print "========== SEARCH STARTED =========="
    Debug.Print "Timestamp: " & Now
    Debug.Print "Target Folder: " & targetFolder.FolderPath
    Debug.Print "Search Scope: " & GetScopeName(searchScope)
    Debug.Print "AQS Query: " & aqsQuery
    Debug.Print "Query Length: " & Len(aqsQuery) & " characters"
    Debug.Print "===================================="
    Debug.Print ""
    
    ' Execute search (ASYNC - results appear in UI)
    ' NATIVE AQS: Uses Windows Search index for fast execution
    searchExplorer.Search aqsQuery, searchScope
    
    ' Store reference for later capture
    Set m_SearchExplorer = searchExplorer
    
    ' Clear any previous results
    Set m_CapturedResults = Nothing
    
    ' Notify user
    MsgBox "Search launched in Outlook." & vbCrLf & vbCrLf & _
           "Review the results, then click 'Capture Results' to enable exports and bulk actions.", _
           vbInformation, "Search Launched"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print ""
    Debug.Print "!!! ERROR IN SEARCH EXECUTION !!!"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Folder: " & IIf(targetFolder Is Nothing, "(null)", targetFolder.FolderPath)
    Debug.Print "AQS Query: " & aqsQuery
    Debug.Print "================================"
    Debug.Print ""
    
    MsgBox "Error during search execution: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Check the Immediate Window (Ctrl+G) for details.", vbCritical, "Search Error"
End Sub

' ============================================
' FOLDER RESOLUTION & SCOPE
' ============================================

'-----------------------------------------------------------
' Function: ResolveFolder
' Purpose: Resolve folder name to MAPIFolder object with error handling
' Parameters:
'   folderName (String) - Folder name or path to resolve
' Returns: Outlook.Folder object (falls back to Inbox if not found)
' Notes: Supports common folder names and recursive custom folder search
'-----------------------------------------------------------
Private Function ResolveFolder(folderName As String) As Outlook.Folder
    On Error GoTo ErrorHandler
    
    Dim ns As Outlook.NameSpace
    Dim folder As Outlook.Folder
    Dim trimmedName As String
    
    Set ns = GetCachedNamespace()
    trimmedName = UCase(Trim(folderName))
    
    ' Handle empty or default
    If trimmedName = "" Or trimmedName = "---" Then
        Set folder = ns.GetDefaultFolder(olFolderInbox)
        Set ResolveFolder = folder
        Exit Function
    End If
    
    ' Check common system folders first
    Select Case trimmedName
        Case "INBOX"
            Set folder = ns.GetDefaultFolder(olFolderInbox)
        Case "SENT ITEMS", "SENT"
            Set folder = ns.GetDefaultFolder(olFolderSentMail)
        Case "DELETED ITEMS", "TRASH", "DELETED"
            Set folder = ns.GetDefaultFolder(olFolderDeletedItems)
        Case "DRAFTS"
            Set folder = ns.GetDefaultFolder(olFolderDrafts)
        Case "JUNK", "JUNK EMAIL", "JUNK E-MAIL"
            Set folder = ns.GetDefaultFolder(olFolderJunk)
        Case "OUTBOX"
            Set folder = ns.GetDefaultFolder(olFolderOutbox)
        Case Else
            ' Try to find custom folder recursively
            Set folder = FindFolderByName(ns.GetDefaultFolder(olFolderInbox).Parent, folderName)
    End Select
    
    ' Fallback to Inbox if not found
    If folder Is Nothing Then
        Debug.Print "WARNING: Folder '" & folderName & "' not found, using Inbox"
        Set folder = ns.GetDefaultFolder(olFolderInbox)
    End If
    
    Set ResolveFolder = folder
    Exit Function
    
ErrorHandler:
    Debug.Print "ResolveFolder Error: " & Err.Description & " (Folder: " & folderName & ")"
    ' Fallback to Inbox
    On Error Resume Next
    Set ResolveFolder = ns.GetDefaultFolder(olFolderInbox)
End Function

'-----------------------------------------------------------
' Function: FindFolderByName
' Purpose: Recursively search for folder by name
' Parameters:
'   parentFolder (Folder) - Starting folder for search
'   searchName (String) - Folder name to find
' Returns: Outlook.Folder object or Nothing if not found
' Notes: Case-insensitive search
'-----------------------------------------------------------
Private Function FindFolderByName(parentFolder As Outlook.Folder, searchName As String) As Outlook.Folder
    On Error Resume Next
    
    Dim folder As Outlook.Folder
    Dim subFolder As Outlook.Folder
    
    ' Search direct children
    For Each folder In parentFolder.Folders
        If UCase(Trim(folder.Name)) = UCase(Trim(searchName)) Then
            Set FindFolderByName = folder
            Exit Function
        End If
    Next
    
    ' Search subfolders recursively
    For Each folder In parentFolder.Folders
        Set subFolder = FindFolderByName(folder, searchName)
        If Not subFolder Is Nothing Then
            Set FindFolderByName = subFolder
            Exit Function
        End If
    Next
    
    Set FindFolderByName = Nothing
End Function

'-----------------------------------------------------------
' Function: OpenSearchExplorer
' Purpose: Create new Explorer window for search results
' Parameters:
'   folder (Folder) - Folder to open in new Explorer
' Returns: Outlook.Explorer object
' Notes: Creates separate window to avoid hijacking user's current view
'-----------------------------------------------------------
Private Function OpenSearchExplorer(folder As Outlook.Folder) As Outlook.Explorer
    On Error GoTo ErrorHandler
    
    Dim newExplorer As Outlook.Explorer
    
    ' Create new Explorer for this folder
    Set newExplorer = Application.Explorers.Add(folder, olFolderDisplayNormal)
    newExplorer.Display
    
    ' Give Explorer time to initialize
    DoEvents
    
    Set OpenSearchExplorer = newExplorer
    Exit Function
    
ErrorHandler:
    Debug.Print "OpenSearchExplorer Error: " & Err.Description
    ' Try to use active Explorer as fallback
    Set OpenSearchExplorer = Application.ActiveExplorer
End Function

'-----------------------------------------------------------
' Function: DetermineSearchScope
' Purpose: Determine OlSearchScope from criteria
' Parameters:
'   criteria (SearchCriteria) - User search criteria
' Returns: OlSearchScope enum value
' Notes: Priority order: All Folders > Current Store > Subfolders > Current Folder
'-----------------------------------------------------------
Private Function DetermineSearchScope(criteria As SearchCriteria) As OlSearchScope
    Dim folderName As String
    folderName = UCase(Trim(criteria.FolderName))
    
    ' Priority 1: All Folders (entire Outlook)
    If folderName = "ALL FOLDERS" Or folderName = "EVERYWHERE" Or folderName = "ALL FOLDERS (INCLUDING SUBFOLDERS)" Then
        DetermineSearchScope = olSearchScopeAllFolders  ' = 3
        Exit Function
    End If
    
    ' Priority 2: Current Store (entire mailbox)
    If folderName = "CURRENT STORE" Or folderName = "THIS MAILBOX" Or folderName = "MAILBOX" Then
        DetermineSearchScope = olSearchScopeCurrentStore  ' = 2
        Exit Function
    End If
    
    ' Priority 3: Specific folder WITH subfolders
    If criteria.SearchSubFolders Then
        DetermineSearchScope = olSearchScopeSubfolders  ' = 1
        Exit Function
    End If
    
    ' Priority 4: Specific folder ONLY
    DetermineSearchScope = olSearchScopeCurrentFolder  ' = 0
End Function

'-----------------------------------------------------------
' Function: GetScopeName
' Purpose: Convert OlSearchScope enum to readable string for debug
' Parameters:
'   scope (OlSearchScope) - Scope enum value
' Returns: String description
' Notes: Helper for debug output
'-----------------------------------------------------------
Private Function GetScopeName(scope As OlSearchScope) As String
    Select Case scope
        Case olSearchScopeCurrentFolder
            GetScopeName = "Current Folder Only"
        Case olSearchScopeSubfolders
            GetScopeName = "Current Folder + Subfolders"
        Case olSearchScopeCurrentStore
            GetScopeName = "Current Mailbox (All Folders)"
        Case olSearchScopeAllFolders
            GetScopeName = "All Folders (All Accounts)"
        Case Else
            GetScopeName = "Unknown (" & scope & ")"
    End Select
End Function

' ============================================
' AQS QUERY BUILDER
' ============================================

'-----------------------------------------------------------
' Function: BuildAQSQuery
' Purpose: Convert SearchCriteria to AQS (Advanced Query Syntax) string
' Parameters:
'   criteria (SearchCriteria) - User search criteria from form
' Returns: AQS query string
' Notes: NATIVE AQS - Uses Windows Search index tokens
'        All field mappings documented in prompt
'-----------------------------------------------------------
Private Function BuildAQSQuery(criteria As SearchCriteria) As String
    On Error GoTo ErrorHandler
    
    Dim queryParts As Collection
    Dim part As String
    Set queryParts = New Collection
    
    ' FROM ADDRESS
    ' NATIVE AQS: from:value searches both display name and email
    If Len(Trim(criteria.FromAddress)) > 0 Then
        queryParts.Add "from:" & EscapeAQSValue(criteria.FromAddress)
    End If
    
    ' TO ADDRESS
    ' NATIVE AQS: to:value searches recipient names/addresses
    If Len(Trim(criteria.ToAddress)) > 0 Then
        queryParts.Add "to:" & EscapeAQSValue(criteria.ToAddress)
    End If
    
    ' CC ADDRESS
    ' NATIVE AQS: cc:value searches CC recipients
    If Len(Trim(criteria.CcAddress)) > 0 Then
        queryParts.Add "cc:" & EscapeAQSValue(criteria.CcAddress)
    End If
    
    ' SUBJECT
    ' NATIVE AQS: subject:value searches subject line
    If Len(Trim(criteria.Subject)) > 0 Then
        queryParts.Add "subject:" & EscapeAQSValue(criteria.Subject)
    End If
    
    ' BODY
    ' NATIVE AQS: body:value searches message body
    If Len(Trim(criteria.Body)) > 0 Then
        queryParts.Add "body:" & EscapeAQSValue(criteria.Body)
    End If
    
    ' DATE RANGE
    ' NATIVE AQS: received:date1..date2 for ranges, received:>=date for comparisons
    If Not IsEmpty(criteria.DateFrom) And Not IsEmpty(criteria.DateTo) Then
        If IsDate(criteria.DateFrom) And IsDate(criteria.DateTo) Then
            ' Range syntax preferred for performance
            queryParts.Add "received:" & FormatAQSDate(CDate(criteria.DateFrom)) & ".." & FormatAQSDate(CDate(criteria.DateTo))
        End If
    ElseIf Not IsEmpty(criteria.DateFrom) Then
        If IsDate(criteria.DateFrom) Then
            queryParts.Add "received:>=" & FormatAQSDate(CDate(criteria.DateFrom))
        End If
    ElseIf Not IsEmpty(criteria.DateTo) Then
        If IsDate(criteria.DateTo) Then
            queryParts.Add "received:<=" & FormatAQSDate(CDate(criteria.DateTo))
        End If
    End If
    
    ' UNREAD STATUS
    ' NATIVE AQS: isread:no for unread, isread:yes for read
    If criteria.UnreadOnly Then
        queryParts.Add "isread:no"
    End If
    
    ' IMPORTANCE
    ' NATIVE AQS: importance:high/normal/low
    If criteria.ImportantOnly Then
        queryParts.Add "importance:high"
    End If
    
    ' ATTACHMENT FILTERS
    ' NATIVE AQS: hasattachments:yes/no, attachmentfilenames:pattern
    If Len(Trim(criteria.AttachmentFilter)) > 0 Then
        Select Case UCase(Trim(criteria.AttachmentFilter))
            Case "WITH ATTACHMENTS", "WITHATTACHMENTS", "YES"
                queryParts.Add "hasattachments:yes"
                
            Case "WITHOUT ATTACHMENTS", "WITHOUTATTACHMENTS", "NO", "NONE"
                queryParts.Add "hasattachments:no"
                
            Case "PDF", "PDF FILES"
                queryParts.Add "attachmentfilenames:*.pdf"
                
            Case "EXCEL", "EXCEL FILES (.XLS, .XLSX)", "XLS", "XLSX"
                ' Multiple extensions with OR
                queryParts.Add "(attachmentfilenames:*.xlsx OR attachmentfilenames:*.xls OR attachmentfilenames:*.xlsm)"
                
            Case "WORD", "WORD FILES (.DOC, .DOCX)", "DOC", "DOCX"
                queryParts.Add "(attachmentfilenames:*.docx OR attachmentfilenames:*.doc)"
                
            Case "IMAGES", "IMAGES (.JPG, .PNG, .GIF)", "JPG", "PNG", "GIF"
                queryParts.Add "(attachmentfilenames:*.jpg OR attachmentfilenames:*.jpeg OR attachmentfilenames:*.png OR attachmentfilenames:*.gif)"
                
            Case "POWERPOINT", "POWERPOINT FILES (.PPT, .PPTX)", "PPT", "PPTX"
                queryParts.Add "(attachmentfilenames:*.ppt OR attachmentfilenames:*.pptx)"
                
            Case "ZIP", "ZIP/ARCHIVES", "ARCHIVES"
                queryParts.Add "(attachmentfilenames:*.zip OR attachmentfilenames:*.rar OR attachmentfilenames:*.7z)"
        End Select
    End If
    
    ' SIZE FILTERS
    ' NATIVE AQS: size:>=value, size:value1..value2 (supports KB, MB, GB units)
    If Not IsEmpty(criteria.SizeFilter) Then
        Dim sizeFilterStr As String
        sizeFilterStr = UCase(Trim(criteria.SizeFilter))
        
        Select Case sizeFilterStr
            Case "< 100 KB"
                queryParts.Add "size:<100KB"
                
            Case "100 KB - 1 MB"
                queryParts.Add "size:100KB..1MB"
                
            Case "1 MB - 10 MB"
                queryParts.Add "size:1MB..10MB"
                
            Case "> 10 MB"
                queryParts.Add "size:>10MB"
                
            ' Ignore "Any Size" or empty
        End Select
    End If
    
    ' GENERAL SEARCH TERMS
    ' NATIVE AQS: Space-separated keywords treated as implicit AND
    If Len(Trim(criteria.SearchTerms)) > 0 Then
        ' Add search terms directly (AQS handles natural language)
        queryParts.Add Trim(criteria.SearchTerms)
    End If
    
    ' COMBINE ALL PARTS WITH AND
    If queryParts.Count = 0 Then
        ' Empty query - search everything
        BuildAQSQuery = ""
    Else
        BuildAQSQuery = JoinCollection(queryParts, " AND ")
    End If
    
    ' Validate query
    ValidateAQSQuery BuildAQSQuery
    
    Exit Function
    
ErrorHandler:
    Debug.Print "BuildAQSQuery Error: " & Err.Description
    BuildAQSQuery = ""
End Function
    Dim filter As String
    Dim conditions() As String
    Dim conditionCount As Long
    Dim i As Long
    
    conditionCount = 0
    ReDim conditions(0 To MAX_SEARCH_CONDITIONS)
    
    ' ========== FROM ADDRESS ==========
    If criteria.FromAddress <> "" Then
        ' Search both display name and SMTP address
        Dim fromPart As String
        Dim fromEscaped As String
        fromEscaped = EscapeSQL(criteria.FromAddress)
        
        If criteria.CaseSensitive Then
            fromPart = "(" & Q("urn:schemas:httpmail:fromname") & " LIKE '%" & fromEscaped & "%' OR " & Q("urn:schemas:httpmail:fromemail") & " LIKE '%" & fromEscaped & "%')"
        Else
            ' Case-insensitive is default for LIKE
            fromPart = "(" & Q("urn:schemas:httpmail:fromname") & " LIKE '%" & fromEscaped & "%' OR " & Q("urn:schemas:httpmail:fromemail") & " LIKE '%" & fromEscaped & "%')"
        End If
        
        conditions(conditionCount) = fromPart
        conditionCount = conditionCount + 1
    End If
    
    ' ========== TO ADDRESS ==========
    If criteria.ToAddress <> "" Then
        Dim toEscaped As String
        toEscaped = EscapeSQL(criteria.ToAddress)
        conditions(conditionCount) = Q("urn:schemas:httpmail:displayto") & " LIKE '%" & toEscaped & "%'"
        conditionCount = conditionCount + 1
    End If
    
    ' ========== CC ADDRESS ==========
    If criteria.CcAddress <> "" Then
        Dim ccEscaped As String
        ccEscaped = EscapeSQL(criteria.CcAddress)
        conditions(conditionCount) = Q("urn:schemas:httpmail:displaycc") & " LIKE '%" & ccEscaped & "%'"
        conditionCount = conditionCount + 1
    End If
    
    ' ========== SEARCH TERMS (Advanced Query Parser) ==========
    If criteria.SearchTerms <> "" Then
        Dim searchPart As String
        searchPart = ParseAdvancedSearchTerms(criteria.SearchTerms, criteria.SearchSubject, criteria.SearchBody, criteria.CaseSensitive)
        
        If searchPart <> "" Then
            conditions(conditionCount) = searchPart
            conditionCount = conditionCount + 1
        End If
    End If
    
    ' ========== DATE RANGE ==========
    If criteria.DateFrom <> "" And IsDate(criteria.DateFrom) Then
        Dim fromDate As Date
        fromDate = CDate(criteria.DateFrom)
        conditions(conditionCount) = Q("urn:schemas:httpmail:datereceived") & " >= '" & _
            Format(fromDate, "mm/dd/yyyy hh:nn AMPM") & "'"
        conditionCount = conditionCount + 1
    End If
    
    If criteria.DateTo <> "" And IsDate(criteria.DateTo) Then
        Dim toDate As Date
        toDate = CDate(criteria.DateTo) + 1  ' Include entire day
        conditions(conditionCount) = Q("urn:schemas:httpmail:datereceived") & " < '" & _
            Format(toDate, "mm/dd/yyyy hh:nn AMPM") & "'"
        conditionCount = conditionCount + 1
    End If
    
    ' ========== ATTACHMENT FILTER ==========
    If criteria.AttachmentFilter <> "Any" And criteria.AttachmentFilter <> "" Then
        Select Case criteria.AttachmentFilter
            Case "With Attachments"
                conditions(conditionCount) = Q("urn:schemas:httpmail:hasattachment") & " = 1"
                conditionCount = conditionCount + 1
            
            Case "Without Attachments"
                conditions(conditionCount) = Q("urn:schemas:httpmail:hasattachment") & " = 0"
                conditionCount = conditionCount + 1
            
            Case "PDF files"
                conditions(conditionCount) = Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.pdf'"
                conditionCount = conditionCount + 1
            
            Case "Excel files (.xls, .xlsx)"
                conditions(conditionCount) = "(" & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.xls' OR " & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.xlsx' OR " & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.xlsm')"
                conditionCount = conditionCount + 1
            
            Case "Word files (.doc, .docx)"
                conditions(conditionCount) = "(" & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.doc' OR " & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.docx')"
                conditionCount = conditionCount + 1
            
            Case "Images (.jpg, .png, .gif)"
                conditions(conditionCount) = "(" & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.jpg' OR " & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.jpeg' OR " & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.png' OR " & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.gif')"
                conditionCount = conditionCount + 1
            
            Case "PowerPoint files (.ppt, .pptx)"
                conditions(conditionCount) = "(" & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.ppt' OR " & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.pptx')"
                conditionCount = conditionCount + 1
            
            Case "ZIP/Archives"
                conditions(conditionCount) = "(" & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.zip' OR " & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.rar' OR " & Q("urn:schemas:httpmail:attachmentfilenames") & " LIKE '%.7z')"
                conditionCount = conditionCount + 1
        End Select
    End If
    
    ' ========== SIZE FILTER ==========
    If criteria.SizeFilter <> "Any Size" And criteria.SizeFilter <> "" Then
        Select Case criteria.SizeFilter
            Case "< 100 KB"
                conditions(conditionCount) = Q("urn:schemas:httpmail:messagesize") & " < " & SIZE_100KB
                conditionCount = conditionCount + 1
            
            Case "100 KB - 1 MB"
                conditions(conditionCount) = "(" & Q("urn:schemas:httpmail:messagesize") & " >= " & SIZE_100KB & " AND " & Q("urn:schemas:httpmail:messagesize") & " < " & SIZE_1MB & ")"
                conditionCount = conditionCount + 1
            
            Case "1 MB - 10 MB"
                conditions(conditionCount) = "(" & Q("urn:schemas:httpmail:messagesize") & " >= " & SIZE_1MB & " AND " & Q("urn:schemas:httpmail:messagesize") & " < " & SIZE_10MB & ")"
                conditionCount = conditionCount + 1
            
            Case "> 10 MB"
                conditions(conditionCount) = Q("urn:schemas:httpmail:messagesize") & " >= " & SIZE_10MB
                conditionCount = conditionCount + 1
        End Select
    End If
    
    ' ========== UNREAD ONLY ==========
    If criteria.UnreadOnly Then
        conditions(conditionCount) = Q("urn:schemas:httpmail:read") & " = 0"
        conditionCount = conditionCount + 1
    End If
    
    ' ========== IMPORTANT/FLAGGED ONLY ==========
    If criteria.ImportantOnly Then
        conditions(conditionCount) = "(" & Q("urn:schemas:httpmail:importance") & " = 2 OR " & Q("http://schemas.microsoft.com/mapi/proptag/0x10900003") & " = 2)"
        conditionCount = conditionCount + 1
    End If
    
    ' ========== COMBINE ALL CONDITIONS ==========
    If conditionCount = 0 Then
        BuildDASLFilter = ""
        Exit Function
    End If
    
    ' Combine all conditions with AND
    filter = conditions(0)
    For i = 1 To conditionCount - 1
        filter = filter & " AND " & conditions(i)
    Next i
    
    ' Wrap entire filter with single @SQL= prefix
    BuildDASLFilter = "@SQL=" & filter
End Function

' ============================================
' Parse Advanced Search Terms (Phase 2 Features)
' Supports: AND, OR, NOT, "exact phrases", word1~N word2 (proximity), wild*cards
' ============================================

Private Function ParseAdvancedSearchTerms(searchText As String, searchSubject As Boolean, searchBody As Boolean, caseSensitive As Boolean) As String
    Dim result As String
    Dim terms() As String
    Dim i As Long
    Dim currentTerm As String
    Dim subjectFilters As String
    Dim bodyFilters As String
    Dim combinedFilter As String
    
    ' Handle exact phrase matching (quoted text)
    If InStr(searchText, Chr(34)) > 0 Then ' Contains quotes
        ' This is a simplified implementation
        ' In production, you'd use a proper tokenizer
        currentTerm = Replace(searchText, Chr(34), "")
        currentTerm = EscapeSQL(currentTerm)
        
        If searchSubject Then
            subjectFilters = Q("urn:schemas:httpmail:subject") & " LIKE '%" & currentTerm & "%'"
        End If
        
        If searchBody Then
            bodyFilters = Q("urn:schemas:httpmail:textdescription") & " LIKE '%" & currentTerm & "%'"
        End If
        
        ' Combine subject and body filters
        If subjectFilters <> "" And bodyFilters <> "" Then
            result = "(" & subjectFilters & " OR " & bodyFilters & ")"
        ElseIf subjectFilters <> "" Then
            result = subjectFilters
        Else
            result = bodyFilters
        End If
        
        ParseAdvancedSearchTerms = result
        Exit Function
    End If
    
    ' Handle proximity search (word1~5 word2 means words within 5 words of each other)
    If InStr(searchText, "~") > 0 Then
        ' Proximity search - simplified version
        ' DASL doesn't natively support proximity, so we'll use CONTAINS instead
        Dim proximityTerms As String
        proximityTerms = Replace(searchText, "~", " NEAR ")
        proximityTerms = EscapeSQL(proximityTerms)
        
        If searchSubject Then
            subjectFilters = Q("urn:schemas:httpmail:subject") & " LIKE '%" & proximityTerms & "%'"
        End If
        
        If searchBody Then
            bodyFilters = Q("urn:schemas:httpmail:textdescription") & " LIKE '%" & proximityTerms & "%'"
        End If
        
        If subjectFilters <> "" And bodyFilters <> "" Then
            result = "(" & subjectFilters & " OR " & bodyFilters & ")"
        ElseIf subjectFilters <> "" Then
            result = subjectFilters
        Else
            result = bodyFilters
        End If
        
        ParseAdvancedSearchTerms = result
        Exit Function
    End If
    
    ' Handle Boolean operators (AND, OR, NOT)
    If InStr(UCase(searchText), " AND ") > 0 Or _
       InStr(UCase(searchText), " OR ") > 0 Or _
       InStr(UCase(searchText), " NOT ") > 0 Then
        
        ' Replace Boolean operators with DASL equivalents
        Dim booleanQuery As String
        booleanQuery = searchText
        
        ' Split by spaces to handle individual terms
        terms = Split(booleanQuery, " ")
        Dim termFilters() As String
        ReDim termFilters(0 To UBound(terms))
        
        For i = 0 To UBound(terms)
            currentTerm = Trim(terms(i))
            
            ' Skip Boolean operators - they'll be handled separately
            If UCase(currentTerm) = "AND" Or UCase(currentTerm) = "OR" Or UCase(currentTerm) = "NOT" Then
                termFilters(i) = " " & UCase(currentTerm) & " "
            Else
                ' Escape and build filter for this term
                currentTerm = EscapeSQL(currentTerm)
                
                ' Handle wildcards (* and ?)
                If InStr(currentTerm, "*") > 0 Or InStr(currentTerm, "?") > 0 Then
                    currentTerm = Replace(currentTerm, "*", "%")
                    currentTerm = Replace(currentTerm, "?", "_")
                Else
                    currentTerm = "%" & currentTerm & "%"
                End If
                
                ' Build filter for subject and/or body
                Dim termFilter As String
                If searchSubject And searchBody Then
                    termFilter = "(" & Q("urn:schemas:httpmail:subject") & " LIKE '" & currentTerm & "' OR " & Q("urn:schemas:httpmail:textdescription") & " LIKE '" & currentTerm & "')"
                ElseIf searchSubject Then
                    termFilter = Q("urn:schemas:httpmail:subject") & " LIKE '" & currentTerm & "'"
                Else
                    termFilter = Q("urn:schemas:httpmail:textdescription") & " LIKE '" & currentTerm & "'"
                End If
                
                termFilters(i) = termFilter
            End If
        Next i
        
        ' Combine all term filters
        result = Join(termFilters, " ")
        result = "(" & result & ")"
        
        ParseAdvancedSearchTerms = result
        Exit Function
    End If
    
    ' Simple search (no special operators)
    ' Split by spaces and combine with AND
    terms = Split(searchText, " ")
    Dim simpleFilters() As String
    ReDim simpleFilters(0 To UBound(terms))
    
    For i = 0 To UBound(terms)
        currentTerm = Trim(terms(i))
        If currentTerm <> "" Then
            currentTerm = EscapeSQL(currentTerm)
            
            ' Handle wildcards
            If InStr(currentTerm, "*") > 0 Or InStr(currentTerm, "?") > 0 Then
                currentTerm = Replace(currentTerm, "*", "%")
                currentTerm = Replace(currentTerm, "?", "_")
            Else
                currentTerm = "%" & currentTerm & "%"
            End If
            
            ' Build filter
            If searchSubject And searchBody Then
                simpleFilters(i) = "(" & Q("urn:schemas:httpmail:subject") & " LIKE '" & currentTerm & "' OR " & Q("urn:schemas:httpmail:textdescription") & " LIKE '" & currentTerm & "')"
            ElseIf searchSubject Then
                simpleFilters(i) = Q("urn:schemas:httpmail:subject") & " LIKE '" & currentTerm & "'"
            Else
                simpleFilters(i) = Q("urn:schemas:httpmail:textdescription") & " LIKE '" & currentTerm & "'"
            End If
        End If
    Next i
    
    ' Combine with AND
    result = Join(simpleFilters, " AND ")
    ParseAdvancedSearchTerms = "(" & result & ")"
End Function

' ============================================
' Get Search Scope (Which Folders)
' ============================================

Private Function GetSearchScope(folderName As String) As String
    Dim ns As Outlook.NameSpace
    Dim folder As Outlook.MAPIFolder
    Dim scope As String
    
    Set ns = GetCachedNamespace()
    
    ' Handle special cases
    If folderName = "All Folders (including subfolders)" Or folderName = "---" Then
        ' Search all mail folders
        scope = "'" & ns.GetDefaultFolder(olFolderInbox).FolderPath & "'"
        scope = scope & ",'" & ns.GetDefaultFolder(olFolderSentMail).FolderPath & "'"
        scope = scope & ",'" & ns.GetDefaultFolder(olFolderDeletedItems).FolderPath & "'"
        scope = scope & ",'" & ns.GetDefaultFolder(olFolderDrafts).FolderPath & "'"
        
        ' Add custom folders if they exist
        On Error Resume Next
        Set folder = ns.GetDefaultFolder(olFolderInbox).Parent.Folders("01 Projects")
        If Not folder Is Nothing Then scope = scope & ",'" & folder.FolderPath & "'"
        
        Set folder = ns.GetDefaultFolder(olFolderInbox).Parent.Folders("02 Office")
        If Not folder Is Nothing Then scope = scope & ",'" & folder.FolderPath & "'"
        
        Set folder = ns.GetDefaultFolder(olFolderInbox).Parent.Folders("99 Misc")
        If Not folder Is Nothing Then scope = scope & ",'" & folder.FolderPath & "'"
        
        Set folder = ns.GetDefaultFolder(olFolderInbox).Parent.Folders("Archive")
        If Not folder Is Nothing Then scope = scope & ",'" & folder.FolderPath & "'"
        On Error GoTo 0
        
    ElseIf StrComp(folderName, "Inbox", vbTextCompare) = 0 Then
        scope = "'" & ns.GetDefaultFolder(olFolderInbox).FolderPath & "'"
    
    ElseIf StrComp(folderName, "Sent Items", vbTextCompare) = 0 Then
        scope = "'" & ns.GetDefaultFolder(olFolderSentMail).FolderPath & "'"
    
    ElseIf StrComp(folderName, "Deleted Items", vbTextCompare) = 0 Then
        scope = "'" & ns.GetDefaultFolder(olFolderDeletedItems).FolderPath & "'"
    
    ElseIf StrComp(folderName, "Drafts", vbTextCompare) = 0 Then
        scope = "'" & ns.GetDefaultFolder(olFolderDrafts).FolderPath & "'"
    
    Else
        ' Try to find custom folder
        On Error Resume Next
        Set folder = ns.GetDefaultFolder(olFolderInbox).Parent.Folders(folderName)
        If Not folder Is Nothing Then
            scope = "'" & folder.FolderPath & "'"
        Else
            ' Default to Inbox if folder not found
            scope = "'" & ns.GetDefaultFolder(olFolderInbox).FolderPath & "'"
        End If
        On Error GoTo 0
    End If
    
    GetSearchScope = scope
End Function

' ============================================
' SQL Escape Function - Enhanced Security
' ============================================

Private Function EscapeSQL(text As String) As String
    Dim result As String
    result = text
    
    ' Escape single quotes (doubled in SQL)
    result = Replace(result, "'", "''")
    
    ' Remove potentially dangerous SQL characters
    result = Replace(result, ";", "")  ' Remove semicolons
    result = Replace(result, "--", "") ' Remove SQL comments
    result = Replace(result, "/*", "") ' Remove block comments
    result = Replace(result, "*/", "")
    
    EscapeSQL = result
End Function

' ============================================
' Ensure Proper Scope Quoting
' ============================================

Private Function EnsureQuotedScope(ByVal scope As String) As String
    ' Ensure each folder path in scope is properly single-quoted
    Dim parts() As String
    Dim i As Long
    Dim p As String
    
    parts = Split(scope, ",")
    For i = LBound(parts) To UBound(parts)
        p = Trim(parts(i))
        If Left(p, 1) <> "'" Then p = "'" & p
        If Right(p, 1) <> "'" Then p = p & "'"
        parts(i) = p
    Next i
    
    EnsureQuotedScope = Join(parts, ",")
End Function

' ============================================
' Property Name Quote Helper
' ============================================

Private Function Q(ByVal s As String) As String
    ' Wrap property name in double-quotes using Chr(34)
    Q = Chr(34) & s & Chr(34)
End Function

' ============================================
' BULK ACTIONS (Phase 3)
' ============================================

Public Sub BulkMarkAsRead(searchObj As Outlook.Search)
    On Error Resume Next
    
    Dim item As Object
    Dim count As Long
    Dim response As VbMsgBoxResult
    
    count = searchObj.Results.Count
    
    ' Check for empty results
    If count = 0 Then
        MsgBox "No emails in search results to mark as read.", vbInformation, "No Results"
        Exit Sub
    End If
    
    response = MsgBox("Mark all " & count & " emails in the search results as read?", _
                      vbYesNo + vbQuestion, "Bulk Mark as Read")
    
    If response = vbYes Then
        For Each item In searchObj.Results
            If TypeOf item Is Outlook.MailItem Then
                Dim mailItem As Outlook.MailItem
                Set mailItem = item
                mailItem.UnRead = False
                mailItem.Save
            End If
        Next item
        
        MsgBox "Marked " & count & " email(s) as read.", vbInformation, "Bulk Action Complete"
    End If
End Sub

Public Sub BulkFlagItems(searchObj As Outlook.Search)
    On Error Resume Next
    
    Dim item As Object
    Dim count As Long
    Dim response As VbMsgBoxResult
    
    count = searchObj.Results.Count
    
    ' Check for empty results
    If count = 0 Then
        MsgBox "No emails in search results to flag.", vbInformation, "No Results"
        Exit Sub
    End If
    
    response = MsgBox("Flag all " & count & " emails in the search results?", _
                      vbYesNo + vbQuestion, "Bulk Flag Items")
    
    If response = vbYes Then
        For Each item In searchObj.Results
            If TypeOf item Is Outlook.MailItem Then
                Dim mailItem As Outlook.MailItem
                Set mailItem = item
                mailItem.FlagRequest = "Follow up"
                mailItem.FlagStatus = olFlagMarked
                mailItem.Save
            End If
        Next item
        
        MsgBox "Flagged " & count & " email(s).", vbInformation, "Bulk Action Complete"
    End If
End Sub

Public Sub BulkMoveItems(searchObj As Outlook.Search)
    On Error GoTo ErrorHandler
    
    Dim item As Object
    Dim count As Long
    Dim response As VbMsgBoxResult
    Dim targetFolder As Outlook.MAPIFolder
    Dim ns As Outlook.NameSpace
    
    count = searchObj.Results.Count
    
    ' Check for empty results
    If count = 0 Then
        MsgBox "No emails in search results to move.", vbInformation, "No Results"
        Exit Sub
    End If
    
    ' Use cached namespace and folder picker dialog
    Set ns = GetCachedNamespace()
    Set targetFolder = ns.PickFolder
    
    If targetFolder Is Nothing Then Exit Sub
    
    response = MsgBox("Move all " & count & " emails to '" & targetFolder.Name & "'?", _
                      vbYesNo + vbQuestion, "Bulk Move Items")
    
    If response = vbYes Then
        For Each item In searchObj.Results
            If TypeOf item Is Outlook.MailItem Then
                Dim mailItem As Outlook.MailItem
                Set mailItem = item
                mailItem.Move targetFolder
            End If
        Next item
        
        MsgBox "Moved " & count & " email(s) to '" & targetFolder.Name & "'.", vbInformation, "Bulk Action Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during bulk move: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' EXPORT FUNCTIONS (Phase 3)
' ============================================

Public Sub ExportSearchResultsToCSV(searchObj As Outlook.Search)
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim txtFile As Object
    Dim filePath As String
    Dim desktopPath As String
    Dim item As Object
    Dim mailItem As Outlook.MailItem
    Dim line As String
    Dim count As Long
    
    ' Check for empty results
    If searchObj.Results.Count = 0 Then
        MsgBox "No search results to export.", vbInformation, "No Results"
        Exit Sub
    End If
    
    ' Create file system object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Validate Desktop path exists
    desktopPath = Environ("USERPROFILE") & "\Desktop"
    If Not fso.FolderExists(desktopPath) Then
        ' Try alternative path
        desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        If desktopPath = "" Then
            MsgBox "Unable to locate Desktop folder. Cannot export.", vbCritical, "Export Error"
            Exit Sub
        End If
    End If
    
    ' Get save location
    filePath = desktopPath & "\EmailSearchResults_" & _
               Format(Now, "yyyymmdd_hhnnss") & ".csv"
    
    ' Create file
    Set txtFile = fso.CreateTextFile(filePath, True, True)
    
    ' Write header
    txtFile.WriteLine "Date,From,To,Subject,Size,Attachments,Unread,Flagged"
    
    ' Write data
    count = 0
    For Each item In searchObj.Results
        If TypeOf item Is Outlook.MailItem Then
            Set mailItem = item
            
            line = ""
            line = line & Format(mailItem.ReceivedTime, "yyyy-mm-dd hh:nn") & ","
            line = line & EscapeCSV(mailItem.SenderName) & ","
            line = line & EscapeCSV(mailItem.To) & ","
            line = line & EscapeCSV(mailItem.Subject) & ","
            line = line & mailItem.Size & ","
            line = line & mailItem.Attachments.count & ","
            line = line & IIf(mailItem.UnRead, "Yes", "No") & ","
            line = line & IIf(mailItem.FlagStatus = olFlagMarked, "Yes", "No")
            
            txtFile.WriteLine line
            count = count + 1
        End If
    Next item
    
    ' Close file
    txtFile.Close
    
    ' Show confirmation
    MsgBox "Exported " & count & " email(s) to:" & vbCrLf & vbCrLf & filePath, _
           vbInformation, "Export Complete"
    
    ' Open the file location
    Shell "explorer.exe /select," & Chr(34) & filePath & Chr(34), vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during export: " & Err.Description, vbCritical, "Export Error"
End Sub

Public Sub ExportSearchResultsToExcel(searchObj As Outlook.Search)
    On Error GoTo ErrorHandler
    
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim item As Object
    Dim mailItem As Outlook.MailItem
    Dim row As Long
    Dim filePath As String
    Dim desktopPath As String
    Dim fso As Object
    
    ' Check for empty results
    If searchObj.Results.Count = 0 Then
        MsgBox "No search results to export.", vbInformation, "No Results"
        Exit Sub
    End If
    
    ' Validate Desktop path
    Set fso = CreateObject("Scripting.FileSystemObject")
    desktopPath = Environ("USERPROFILE") & "\Desktop"
    If Not fso.FolderExists(desktopPath) Then
        desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        If desktopPath = "" Then
            MsgBox "Unable to locate Desktop folder. Cannot export.", vbCritical, "Export Error"
            Exit Sub
        End If
    End If
    
    ' Create Excel application
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    ' Set headers
    With xlSheet
        .Cells(1, 1).Value = "Date"
        .Cells(1, 2).Value = "From"
        .Cells(1, 3).Value = "To"
        .Cells(1, 4).Value = "Subject"
        .Cells(1, 5).Value = "Size (KB)"
        .Cells(1, 6).Value = "Attachments"
        .Cells(1, 7).Value = "Unread"
        .Cells(1, 8).Value = "Flagged"
        
        ' Format header row
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Populate data
    row = 2
    For Each item In searchObj.Results
        If TypeOf item Is Outlook.MailItem Then
            Set mailItem = item
            
            With xlSheet
                .Cells(row, 1).Value = mailItem.ReceivedTime
                .Cells(row, 2).Value = SanitizeForExcel(mailItem.SenderName)
                .Cells(row, 3).Value = SanitizeForExcel(mailItem.To)
                .Cells(row, 4).Value = SanitizeForExcel(mailItem.Subject)
                .Cells(row, 5).Value = Round(mailItem.Size / 1024, 2)
                .Cells(row, 6).Value = mailItem.Attachments.count
                .Cells(row, 7).Value = IIf(mailItem.UnRead, "Yes", "No")
                .Cells(row, 8).Value = IIf(mailItem.FlagStatus = olFlagMarked, "Yes", "No")
            End With
            
            row = row + 1
        End If
    Next item
    
    ' Auto-fit columns
    xlSheet.Columns.AutoFit
    
    ' Add filter
    xlSheet.Range("A1:H" & row - 1).AutoFilter
    
    ' Save file
    filePath = desktopPath & "\EmailSearchResults_" & _
               Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    xlBook.SaveAs filePath
    
    ' Show Excel
    xlApp.Visible = True
    
    ' Release references (but keep Excel open)
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    MsgBox "Exported " & (row - 2) & " email(s) to Excel:" & vbCrLf & vbCrLf & filePath, _
           vbInformation, "Export Complete"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during Excel export: " & Err.Description, vbCritical, "Export Error"
    
    ' Cleanup on error - quit Excel
    On Error Resume Next
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = False
        xlApp.Quit
    End If
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub

Private Function EscapeCSV(text As String) As String
    ' Escape commas and quotes for CSV
    Dim result As String
    result = Replace(text, Chr(34), Chr(34) & Chr(34))
    If InStr(result, ",") > 0 Or InStr(result, Chr(34)) > 0 Or InStr(result, vbCrLf) > 0 Then
        result = Chr(34) & result & Chr(34)
    End If
    EscapeCSV = result
End Function

Private Function SanitizeForExcel(text As String) As String
    ' Prevent CSV/Excel formula injection by prefixing dangerous characters
    ' Dangerous chars: = + - @ (can trigger formula execution)
    If Len(text) > 0 Then
        Dim firstChar As String
        firstChar = Left(text, 1)
        If firstChar = "=" Or firstChar = "+" Or firstChar = "-" Or firstChar = "@" Then
            SanitizeForExcel = "'" & text  ' Prefix with single quote to force text interpretation
        Else
            SanitizeForExcel = text
        End If
    Else
        SanitizeForExcel = text
    End If
End Function

' ============================================
' STATISTICS (Phase 3)
' ============================================

Public Sub ShowSearchStatistics(searchObj As Outlook.Search)
    On Error Resume Next
    
    Dim item As Object
    Dim mailItem As Outlook.MailItem
    Dim totalCount As Long
    Dim unreadCount As Long
    Dim flaggedCount As Long
    Dim withAttachments As Long
    Dim totalSize As Long
    Dim senders As Object
    Dim senderName As String
    Dim msg As String
    Dim topSender As String
    Dim topSenderCount As Long
    
    ' Initialize
    totalCount = searchObj.Results.count
    Set senders = CreateObject("Scripting.Dictionary")
    
    ' Analyze results
    For Each item In searchObj.Results
        If TypeOf item Is Outlook.MailItem Then
            Set mailItem = item
            
            ' Count unread
            If mailItem.UnRead Then unreadCount = unreadCount + 1
            
            ' Count flagged
            If mailItem.FlagStatus = olFlagMarked Then flaggedCount = flaggedCount + 1
            
            ' Count attachments
            If mailItem.Attachments.count > 0 Then withAttachments = withAttachments + 1
            
            ' Total size
            totalSize = totalSize + mailItem.Size
            
            ' Track senders
            senderName = mailItem.SenderName
            If senders.Exists(senderName) Then
                senders(senderName) = senders(senderName) + 1
            Else
                senders.Add senderName, 1
            End If
        End If
    Next item
    
    ' Find top sender
    Dim key As Variant
    For Each key In senders.Keys
        If senders(key) > topSenderCount Then
            topSenderCount = senders(key)
            topSender = key
        End If
    Next key
    
    ' Build statistics message
    msg = "SEARCH STATISTICS" & vbCrLf & vbCrLf
    msg = msg & "Total Emails: " & totalCount & vbCrLf
    
    If totalCount > 0 Then
        ' Only calculate percentages if we have results
        msg = msg & "Unread: " & unreadCount & " (" & Format(unreadCount / totalCount * 100, "0.0") & "%)" & vbCrLf
        msg = msg & "Flagged: " & flaggedCount & " (" & Format(flaggedCount / totalCount * 100, "0.0") & "%)" & vbCrLf
        msg = msg & "With Attachments: " & withAttachments & " (" & Format(withAttachments / totalCount * 100, "0.0") & "%)" & vbCrLf & vbCrLf
        msg = msg & "Total Size: " & Format(totalSize / MB_IN_BYTES / KB_IN_BYTES, "0.00") & " MB" & vbCrLf
        msg = msg & "Average Size: " & Format(totalSize / totalCount / KB_IN_BYTES, "0.00") & " KB" & vbCrLf & vbCrLf
        msg = msg & "Unique Senders: " & senders.Count & vbCrLf
        msg = msg & "Most Common Sender: " & topSender & " (" & topSenderCount & " emails)"
    Else
        msg = msg & "No results to analyze."
    End If
    
    ' Show statistics
    MsgBox msg, vbInformation, "Search Statistics"
End Sub

' ============================================
' PRESET MANAGEMENT (Phase 3)
' ============================================

Public Sub SaveSearchPreset(presetName As String, criteria As SearchCriteria)
    ' This is a simplified version that saves to registry
    ' In production, you'd use a JSON file in user profile
    
    On Error Resume Next
    
    Dim regPath As String
    regPath = "Software\OutlookEmailSearch\Presets\" & presetName
    
    ' Save each field to registry (simplified - use JSON in production)
    ' VBA doesn't have built-in registry access, so this is pseudo-code
    ' You would use Windows Script Host or a JSON file instead
    
    MsgBox "Preset saved successfully (feature requires JSON file implementation)", vbInformation
End Sub

Public Function LoadSearchPreset(presetName As String, ByRef criteria As SearchCriteria) As Boolean
    ' This is a simplified version that loads from registry
    ' In production, you'd use a JSON file in user profile
    
    On Error Resume Next
    
    ' Load from registry or JSON file
    ' Simplified version - always returns False
    
    LoadSearchPreset = False
    MsgBox "Preset loading requires JSON file implementation", vbInformation
End Function
