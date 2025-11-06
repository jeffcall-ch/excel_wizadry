'This VBA script for Outlook reads the current email's subject.
'It creates a "cleaned" search phrase and inputs it into the main Outlook search field.
'Now supports PIRS email format removal as well.

 

Public Sub CleanedSubjectSearch()

  Dim Folder As Outlook.MAPIFolder
  Set Folder = Application.ActiveExplorer.CurrentFolder
  Set obj = Application.ActiveExplorer.Selection(1)

 

  If Not TypeOf obj Is Outlook.MailItem Then
    'this sample is for emails
    Exit Sub
  End If

 

  Set Mail = obj
  Set Folder = Mail.Parent

 

  'Define the cleaned search string - getting the subject of the current email
  Dim SearchWords As Variant
  Dim CleanSubject As String
  CleanSubject = Mail.Subject
  
  'Step 1: Remove PIRS prefix pattern if present
  'PIRS pattern: <Project>, <Code1>/<Code2>/<5-digit-number>: <Original Subject>
  'Example: "Abu Dhabi, KVI AG/INOV/00071: Re: Original Subject"
  CleanSubject = RemovePIRSPrefix(CleanSubject)
  
  'Step 2: Convert to lowercase for case-insensitive replacement
  CleanSubject = LCase$(CleanSubject)
  
  'Step 3: Remove common email prefixes and tags
  SearchWords = Array(" [EXTERNAL] ", " [EXT] ", " RE: ", " FW: ", " AW: ", " WG: ", " TR: ", " RV: ", " SV: ", " FS: ", " VS: ", " VL: ", " ODP: ", " PD: ", " R: ", " I: ")

 

    For Each word In SearchWords
      CleanSubject = " " & CleanSubject 'add a space at front to be able to search for words starting and ending with space
      CleanSubject = Replace(CleanSubject, LCase$(word), " ", vbTextCompare)
    Next word
    CleanSubject = Trim(CleanSubject)

 

  'Debug.Print "Final cleaned subject: " & CleanSubject

 

  'Add search string to top search bar as input
  Dim myOlApp As New Outlook.Application
  quoteChar = """" 'needed to quote the search string, otherwise too many irrelevant results
  txtSearch = quoteChar & CleanSubject & quoteChar
  myOlApp.ActiveExplorer.Search txtSearch, olSearchScopeAllFolders
  Set myOlApp = Nothing

End Sub

 

'Function to remove PIRS prefix from email subject
'PIRS adds: <Project>, <Code1>/<Code2>/<5-digit-number>: before the original subject
'Handles cases with FW:/RE:/[EXT] prefixes before PIRS pattern
'Recursively removes multiple PIRS patterns (nested PIRS references)
'Also removes PIRS references within the subject (e.g., "Re: CODE/CODE/00000 - Subject")
Function RemovePIRSPrefix(ByVal Subject As String) As String
    Dim regEx As Object
    Dim regExRef As Object
    Dim matches As Object
    Dim cleanedSubject As String
    Dim previousSubject As String
    Dim maxIterations As Integer
    Dim iteration As Integer
    
    cleanedSubject = Subject
    maxIterations = 10  'Safety limit to prevent infinite loops
    iteration = 0
    
    'Create regex object for full PIRS prefix
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    'PIRS pattern: optional FW:/RE:/[EXT], then <text>, <text>/<text>/\d{5}: <original>
    'Pattern explanation:
    '  ^               - start of string
    '  (?:(?:FW|RE):\s*(?:\[EXT\]\s*)?)?  - optional FW: or RE: with optional [EXT] after
    '  [^,]+           - project name (anything before comma)
    '  ,\s+            - comma and space
    '  [^/]+/[^/]+/    - Code1/Code2/
    '  \d{5}           - exactly 5 digits
    '  :\s+            - colon and space
    '  (.*)            - capture the rest (original subject)
    
    regEx.Pattern = "^(?:(?:FW|RE):\s*(?:\[EXT\]\s*)?)?[^,]+,\s+[^/]+/[^/]+/\d{5}:\s+(.*)"
    
    'Keep removing PIRS patterns until none are found (handles nested PIRS)
    Do While regEx.Test(cleanedSubject) And iteration < maxIterations
        previousSubject = cleanedSubject
        Set matches = regEx.Execute(cleanedSubject)
        If matches.Count > 0 Then
            'Extract the captured group (original subject after PIRS prefix)
            cleanedSubject = matches(0).SubMatches(0)
            'Debug.Print "PIRS removed (iteration " & iteration & "): " & cleanedSubject
        End If
        iteration = iteration + 1
        
        'Safety check: if nothing changed, break
        If cleanedSubject = previousSubject Then Exit Do
    Loop
    
    'Now remove PIRS references within the subject (e.g., "Re: CODE/CODE/00000 - Subject")
    'Pattern: CODE/CODE/00000 followed by optional " - " or ": "
    Set regExRef = CreateObject("VBScript.RegExp")
    regExRef.Global = True
    regExRef.IgnoreCase = True
    regExRef.Pattern = "[A-Z][A-Z0-9\s\-]+/[A-Z][A-Z0-9\s]+/\d{5}\s*[-:]\s*"
    
    If regExRef.Test(cleanedSubject) Then
        cleanedSubject = regExRef.Replace(cleanedSubject, "")
        'Debug.Print "PIRS reference removed from subject: " & cleanedSubject
    End If
    
    RemovePIRSPrefix = cleanedSubject
    Set regEx = Nothing
    Set regExRef = Nothing
End Function
