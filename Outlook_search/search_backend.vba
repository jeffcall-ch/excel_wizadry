' ============================================
' Module: ThisOutlookSession (AQS STRING BUILDER VERSION - FIXED)
' Purpose: Build AQS search string and insert into Outlook search field
' Features: Simple search string generation with proper quoted phrase handling
' Date: 2025-12-16
' Version: 2.1 - Fixed quoted phrase handling
' ============================================

Option Explicit

' Module-level variables
Private m_FormInstance As frmEmailSearch    ' Reference to search form

' ============================================
' MAIN ENTRY POINT
' ============================================

Public Sub ShowEmailSearchForm()
    On Error GoTo ErrorHandler
    
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

' ============================================
' SEARCH EXECUTION
' ============================================

'-----------------------------------------------------------
' Sub: ExecuteAQSSearch
' Purpose: Build AQS search string and execute search in Outlook
' Parameters:
'   fromAddr - From address/name
'   toAddr - To address/name
'   ccAddr - CC address/name
'   searchTerms - Search terms (with OR/AND operators)
'   dateFrom - Date from (dd.mm.yyyy format)
'   dateTo - Date to (dd.mm.yyyy format)
'   attachmentType - Attachment type index (0=Any, 1=With, 2=Without, 3+=File types)
'   searchInSubject - Search in subject field
'   searchInBody - Search in body field
'   currentFolderOnly - True: current folder only, False: all folders
'-----------------------------------------------------------
Public Sub ExecuteAQSSearch(ByVal fromAddr As String, _
                           ByVal toAddr As String, _
                           ByVal ccAddr As String, _
                           ByVal searchTerms As String, _
                           ByVal dateFrom As String, _
                           ByVal dateTo As String, _
                           ByVal attachmentType As Integer, _
                           ByVal searchInSubject As Boolean, _
                           ByVal searchInBody As Boolean, _
                           ByVal currentFolderOnly As Boolean)
    On Error GoTo ErrorHandler
    
    Dim aqsQuery As String
    Dim queryParts As Collection
    Dim searchScope As OlSearchScope
    Dim myOlApp As Outlook.Application
    
    Set queryParts = New Collection
    Set myOlApp = Application
    
    ' Build AQS query parts
    
    ' 1. Add From/To/CC if provided
    If Len(fromAddr) > 0 Then
        queryParts.Add "from:" & WrapIfNeeded(fromAddr)
    End If
    
    If Len(toAddr) > 0 Then
        queryParts.Add "to:" & WrapIfNeeded(toAddr)
    End If
    
    If Len(ccAddr) > 0 Then
        queryParts.Add "cc:" & WrapIfNeeded(ccAddr)
    End If
    
    ' 2. Add Search Terms with field prefixes (subject/body)
    If Len(searchTerms) > 0 Then
        Dim whereParts As Collection
        Set whereParts = New Collection
        
        ' Parse search terms (convert comma to OR, plus to AND)
        Dim parsedTerms As String
        parsedTerms = ParseSearchTerms(searchTerms)
        
        ' Build field-specific queries
        If searchInSubject Then
            whereParts.Add ApplyFieldToTerms("subject", parsedTerms)
        End If
        
        If searchInBody Then
            whereParts.Add ApplyFieldToTerms("body", parsedTerms)
        End If
        
        ' Combine WHERE clauses with OR
        If whereParts.Count > 0 Then
            If whereParts.Count > 1 Then
                ' Multiple fields: wrap in double parentheses for proper precedence
                queryParts.Add "((" & JoinCollection(whereParts, " OR ") & "))"
            Else
                queryParts.Add whereParts(1)
            End If
        End If
    End If
    
    ' 3. Add Date filters
    If Len(dateFrom) > 0 And Len(dateTo) > 0 Then
        ' Date range
        queryParts.Add "received:" & ConvertDateToAQS(dateFrom) & ".." & ConvertDateToAQS(dateTo)
    ElseIf Len(dateFrom) > 0 Then
        ' Date from only
        queryParts.Add "received:>=" & ConvertDateToAQS(dateFrom)
    ElseIf Len(dateTo) > 0 Then
        ' Date to only
        queryParts.Add "received:<=" & ConvertDateToAQS(dateTo)
    End If
    
    ' 4. Add Attachment filter
    Dim attFilter As String
    attFilter = BuildAttachmentAQS(attachmentType)
    If Len(attFilter) > 0 Then
        queryParts.Add attFilter
    End If
    
    ' 5. Combine all parts with AND
    If queryParts.Count > 1 Then
        aqsQuery = JoinCollection(queryParts, " AND ")
    ElseIf queryParts.Count = 1 Then
        aqsQuery = queryParts(1)
    Else
        ' No criteria - empty search (will show all emails)
        aqsQuery = ""
    End If
    
    ' 6. Determine search scope
    If currentFolderOnly Then
        searchScope = olSearchScopeCurrentFolder
    Else
        searchScope = olSearchScopeAllFolders
    End If
    
    ' 7. Debug output
    Debug.Print ""
    Debug.Print "========== AQS SEARCH =========="
    Debug.Print "Timestamp: " & Now
    Debug.Print "From: " & fromAddr
    Debug.Print "To: " & toAddr
    Debug.Print "CC: " & ccAddr
    Debug.Print "Search Terms: " & searchTerms
    Debug.Print "Search In: Subject=" & searchInSubject & " Body=" & searchInBody
    Debug.Print "Date From: " & dateFrom
    Debug.Print "Date To: " & dateTo
    Debug.Print "Attachment Type: " & attachmentType
    Debug.Print "Current Folder Only: " & currentFolderOnly
    Debug.Print "AQS Query: " & aqsQuery
    Debug.Print "Scope: " & IIf(currentFolderOnly, "Current Folder", "All Folders")
    Debug.Print "================================"
    Debug.Print ""
    
    ' 8. Execute search
    myOlApp.ActiveExplorer.Search aqsQuery, searchScope
    
    Set myOlApp = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "!!! AQS SEARCH ERROR !!!"
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    MsgBox "Search error: " & Err.Description, vbCritical, "Search Error"
End Sub

' ============================================
' HELPER FUNCTIONS
' ============================================

'-----------------------------------------------------------
' Function: ParseSearchTerms
' Purpose: Convert comma to OR, plus to AND
' Parameters:
'   rawTerms - Raw search terms from user
' Returns: Parsed search terms with OR/AND operators
'-----------------------------------------------------------
Private Function ParseSearchTerms(ByVal rawTerms As String) As String
    Dim result As String
    
    result = rawTerms
    
    ' Replace comma with OR
    result = Replace(result, ",", " OR ")
    
    ' Replace plus with AND
    result = Replace(result, "+", " AND ")
    
    ' Clean up multiple spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ParseSearchTerms = Trim(result)
End Function

'-----------------------------------------------------------
' Function: ApplyFieldToTerms (FIXED VERSION)
' Purpose: Apply field prefix to each term/group in search string
' Parameters:
'   fieldName - Field name (subject, body, etc.)
'   terms - Search terms (may contain OR/AND/wildcards/quoted phrases)
' Returns: Field-prefixed query string
' Notes: FIXED to handle quoted phrases correctly
'-----------------------------------------------------------
Private Function ApplyFieldToTerms(ByVal fieldName As String, ByVal terms As String) As String
    Dim result As String
    Dim i As Integer
    Dim char As String
    Dim currentToken As String
    Dim inQuotes As Boolean
    Dim tokens As Collection
    Dim token As Variant
    
    ' Parse terms respecting quoted phrases
    Set tokens = New Collection
    currentToken = ""
    inQuotes = False
    
    ' Character-by-character parsing to respect quoted phrases
    For i = 1 To Len(terms)
        char = Mid(terms, i, 1)
        
        If char = """" Then
            ' Toggle quote state
            inQuotes = Not inQuotes
            currentToken = currentToken & char
        ElseIf char = " " And Not inQuotes Then
            ' Space outside quotes = token boundary
            If Len(Trim(currentToken)) > 0 Then
                tokens.Add Trim(currentToken)
                currentToken = ""
            End If
        Else
            ' Regular character
            currentToken = currentToken & char
        End If
    Next i
    
    ' Add last token
    If Len(Trim(currentToken)) > 0 Then
        tokens.Add Trim(currentToken)
    End If
    
    ' Build result with field prefixes
    result = ""
    Dim hasOR As Boolean
    hasOR = False
    
    ' Check if there are OR operators (need parentheses for AND groups)
    For Each token In tokens
        If UCase(CStr(token)) = "OR" Then
            hasOR = True
            Exit For
        End If
    Next
    
    ' Build query with proper grouping
    If hasOR Then
        ' Has OR: wrap AND groups in parentheses
        Dim andGroup As String
        andGroup = ""
        
        For Each token In tokens
            token = CStr(token)
            
            If Len(token) > 0 Then
                If UCase(token) = "OR" Then
                    ' Close current AND group and add OR
                    If Len(andGroup) > 0 Then
                        result = result & "(" & Trim(andGroup) & ") OR "
                        andGroup = ""
                    End If
                ElseIf UCase(token) = "AND" Then
                    ' Add AND within group
                    andGroup = andGroup & " AND "
                Else
                    ' Add term with field prefix
                    andGroup = andGroup & " " & fieldName & ":" & token
                End If
            End If
        Next
        
        ' Add final AND group
        If Len(andGroup) > 0 Then
            result = result & "(" & Trim(andGroup) & ")"
        End If
    Else
        ' No OR: simple processing (no extra parentheses needed)
        For Each token In tokens
            token = CStr(token)
            
            If Len(token) > 0 Then
                If UCase(token) = "OR" Or UCase(token) = "AND" Then
                    result = result & " " & UCase(token) & " "
                Else
                    result = result & " " & fieldName & ":" & token
                End If
            End If
        Next
    End If
    
    ' Clean up multiple spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ApplyFieldToTerms = Trim(result)
End Function

'-----------------------------------------------------------
' Function: ConvertDateToAQS
' Purpose: Convert dd.mm.yyyy to mm/dd/yyyy (AQS format)
' Parameters:
'   dateStr - Date string in dd.mm.yyyy format
' Returns: Date string in mm/dd/yyyy format
'-----------------------------------------------------------
Private Function ConvertDateToAQS(ByVal dateStr As String) As String
    On Error GoTo ErrorHandler
    
    Dim parts() As String
    Dim day As String, month As String, year As String
    
    ' Split by dot
    parts = Split(dateStr, ".")
    
    If UBound(parts) - LBound(parts) + 1 = 3 Then
        day = parts(0)
        month = parts(1)
        year = parts(2)
        
        ' Return in mm/dd/yyyy format
        ConvertDateToAQS = month & "/" & day & "/" & year
    Else
        ' Invalid format, return as-is
        ConvertDateToAQS = dateStr
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertDateToAQS = dateStr
End Function

'-----------------------------------------------------------
' Function: BuildAttachmentAQS
' Purpose: Build attachment filter AQS string
' Parameters:
'   attachmentType - Attachment type index from dropdown
' Returns: AQS attachment filter string
'-----------------------------------------------------------
Private Function BuildAttachmentAQS(ByVal attachmentType As Integer) As String
    Select Case attachmentType
        Case 0  ' --- (Any)
            BuildAttachmentAQS = ""
        Case 1  ' With Attachments
            BuildAttachmentAQS = "hasattachment:yes"
        Case 2  ' Without Attachments
            BuildAttachmentAQS = "hasattachment:no"
        Case 3  ' PDF Files
            BuildAttachmentAQS = "attachments:*.pdf"
        Case 4  ' Excel Files (.xls, .xlsx)
            BuildAttachmentAQS = "(attachments:*.xls OR attachments:*.xlsx)"
        Case 5  ' Word Files (.doc, .docx)
            BuildAttachmentAQS = "(attachments:*.doc OR attachments:*.docx)"
        Case 6  ' PowerPoint Files (.ppt, .pptx)
            BuildAttachmentAQS = "(attachments:*.ppt OR attachments:*.pptx)"
        Case 7  ' Images (.jpg, .png, .gif)
            BuildAttachmentAQS = "(attachments:*.jpg OR attachments:*.png OR attachments:*.gif)"
        Case 8  ' Text Files (.txt)
            BuildAttachmentAQS = "attachments:*.txt"
        Case 9  ' ZIP/Archives (.zip, .rar)
            BuildAttachmentAQS = "(attachments:*.zip OR attachments:*.rar)"
        Case Else
            BuildAttachmentAQS = ""
    End Select
End Function

'-----------------------------------------------------------
' Function: WrapIfNeeded
' Purpose: Wrap text in quotes if it contains spaces
' Parameters:
'   text - Text to wrap
' Returns: Wrapped or unwrapped text
'-----------------------------------------------------------
Private Function WrapIfNeeded(ByVal text As String) As String
    If InStr(text, " ") > 0 Then
        ' Contains spaces, wrap in quotes
        WrapIfNeeded = """" & text & """"
    Else
        ' No spaces, return as-is
        WrapIfNeeded = text
    End If
End Function

'-----------------------------------------------------------
' Function: JoinCollection
' Purpose: Join collection items with separator
' Parameters:
'   coll - Collection to join
'   separator - Separator string
' Returns: Joined string
'-----------------------------------------------------------
Private Function JoinCollection(ByVal coll As Collection, ByVal separator As String) As String
    Dim result As String
    Dim item As Variant
    Dim isFirst As Boolean
    
    isFirst = True
    For Each item In coll
        If isFirst Then
            result = CStr(item)
            isFirst = False
        Else
            result = result & separator & CStr(item)
        End If
    Next
    
    JoinCollection = result
End Function
