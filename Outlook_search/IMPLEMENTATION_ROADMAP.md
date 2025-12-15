# Implementation Roadmap for Missing Features

This document outlines how to implement the remaining features that are currently placeholders or not yet implemented.

## Priority 1: JSON Preset Storage (Estimated: 2 hours)

### Current Status
- ⚠️ UI is ready (Save/Load buttons work)
- ⚠️ Functions are placeholders
- ⚠️ No file storage

### Implementation Steps

#### 1. Add JSON Helper Functions to ThisOutlookSession

```vba
' Add to top of ThisOutlookSession module
Private Const PRESET_FILE_PATH As String = "\AppData\Roaming\OutlookEmailSearch\presets.json"

Private Function GetPresetFilePath() As String
    Dim userProfile As String
    userProfile = Environ("USERPROFILE")
    GetPresetFilePath = userProfile & PRESET_FILE_PATH
End Function

Private Sub EnsurePresetDirectoryExists()
    Dim fso As Object
    Dim dirPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    dirPath = Environ("USERPROFILE") & "\AppData\Roaming\OutlookEmailSearch"
    
    If Not fso.FolderExists(dirPath) Then
        fso.CreateFolder dirPath
    End If
End Sub

Private Function SerializeCriteria(criteria As SearchCriteria) As String
    ' Create JSON manually (VBA doesn't have native JSON support)
    Dim json As String
    
    json = "{"
    json = json & """FromAddress"": """ & EscapeJSON(criteria.FromAddress) & """, "
    json = json & """SearchTerms"": """ & EscapeJSON(criteria.SearchTerms) & """, "
    json = json & """SearchSubject"": " & LCase(CStr(criteria.SearchSubject)) & ", "
    json = json & """SearchBody"": " & LCase(CStr(criteria.SearchBody)) & ", "
    json = json & """DateFrom"": """ & criteria.DateFrom & """, "
    json = json & """DateTo"": """ & criteria.DateTo & """, "
    json = json & """FolderName"": """ & EscapeJSON(criteria.FolderName) & """, "
    json = json & """CaseSensitive"": " & LCase(CStr(criteria.CaseSensitive)) & ", "
    json = json & """ToAddress"": """ & EscapeJSON(criteria.ToAddress) & """, "
    json = json & """CcAddress"": """ & EscapeJSON(criteria.CcAddress) & """, "
    json = json & """AttachmentFilter"": """ & EscapeJSON(criteria.AttachmentFilter) & """, "
    json = json & """SizeFilter"": """ & EscapeJSON(criteria.SizeFilter) & """, "
    json = json & """UnreadOnly"": " & LCase(CStr(criteria.UnreadOnly)) & ", "
    json = json & """ImportantOnly"": " & LCase(CStr(criteria.ImportantOnly))
    json = json & "}"
    
    SerializeCriteria = json
End Function

Private Function DeserializeCriteria(json As String) As SearchCriteria
    Dim criteria As SearchCriteria
    
    ' Simple JSON parsing (replace with proper parser if needed)
    criteria.FromAddress = GetJSONValue(json, "FromAddress")
    criteria.SearchTerms = GetJSONValue(json, "SearchTerms")
    criteria.SearchSubject = CBool(GetJSONValue(json, "SearchSubject"))
    criteria.SearchBody = CBool(GetJSONValue(json, "SearchBody"))
    criteria.DateFrom = GetJSONValue(json, "DateFrom")
    criteria.DateTo = GetJSONValue(json, "DateTo")
    criteria.FolderName = GetJSONValue(json, "FolderName")
    criteria.CaseSensitive = CBool(GetJSONValue(json, "CaseSensitive"))
    criteria.ToAddress = GetJSONValue(json, "ToAddress")
    criteria.CcAddress = GetJSONValue(json, "CcAddress")
    criteria.AttachmentFilter = GetJSONValue(json, "AttachmentFilter")
    criteria.SizeFilter = GetJSONValue(json, "SizeFilter")
    criteria.UnreadOnly = CBool(GetJSONValue(json, "UnreadOnly"))
    criteria.ImportantOnly = CBool(GetJSONValue(json, "ImportantOnly"))
    
    DeserializeCriteria = criteria
End Function

Private Function GetJSONValue(json As String, key As String) As String
    ' Extract value for a key from JSON string
    Dim startPos As Long
    Dim endPos As Long
    Dim searchStr As String
    Dim value As String
    
    searchStr = """" & key & """: """
    startPos = InStr(json, searchStr)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        endPos = InStr(startPos, json, """")
        value = Mid(json, startPos, endPos - startPos)
    Else
        ' Try boolean/number format
        searchStr = """" & key & """: "
        startPos = InStr(json, searchStr)
        If startPos > 0 Then
            startPos = startPos + Len(searchStr)
            endPos = InStr(startPos, json, ",")
            If endPos = 0 Then endPos = InStr(startPos, json, "}")
            value = Trim(Mid(json, startPos, endPos - startPos))
        End If
    End If
    
    GetJSONValue = value
End Function

Private Function EscapeJSON(text As String) As String
    ' Escape special characters for JSON
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJSON = result
End Function
```

#### 2. Replace SaveSearchPreset Function

```vba
Public Sub SaveSearchPreset(presetName As String, criteria As SearchCriteria)
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim file As Object
    Dim filePath As String
    Dim json As String
    Dim allPresets As Object
    Dim existingContent As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ensure directory exists
    EnsurePresetDirectoryExists
    
    ' Get file path
    filePath = GetPresetFilePath()
    
    ' Load existing presets if file exists
    Set allPresets = CreateObject("Scripting.Dictionary")
    
    If fso.FileExists(filePath) Then
        Set file = fso.OpenTextFile(filePath, 1, False, -1) ' 1 = ForReading, -1 = Unicode
        existingContent = file.ReadAll
        file.Close
        
        ' Parse existing JSON (simplified - assumes format)
        ' In production, use a proper JSON parser
        ' For now, we'll just append
    End If
    
    ' Serialize criteria
    json = SerializeCriteria(criteria)
    
    ' Create preset entry
    Dim presetEntry As String
    presetEntry = """" & presetName & """: " & json
    
    ' Write to file (simplified - overwrites for now)
    Dim fullJSON As String
    If existingContent <> "" And existingContent <> "{}" Then
        ' Remove closing brace, add comma and new preset
        fullJSON = Left(existingContent, Len(existingContent) - 1)
        fullJSON = fullJSON & ", " & presetEntry & "}"
    Else
        ' First preset
        fullJSON = "{" & presetEntry & "}"
    End If
    
    ' Write to file
    Set file = fso.CreateTextFile(filePath, True, True) ' True = overwrite, True = Unicode
    file.Write fullJSON
    file.Close
    
    MsgBox "Preset '" & presetName & "' saved successfully.", vbInformation, "Preset Saved"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error saving preset: " & Err.Description, vbCritical, "Save Error"
End Sub
```

#### 3. Replace LoadSearchPreset Function

```vba
Public Function LoadSearchPreset(presetName As String, ByRef criteria As SearchCriteria) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim file As Object
    Dim filePath As String
    Dim fileContent As String
    Dim presetJSON As String
    Dim startPos As Long
    Dim endPos As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetPresetFilePath()
    
    ' Check if file exists
    If Not fso.FileExists(filePath) Then
        MsgBox "No presets found. Please save a preset first.", vbInformation, "No Presets"
        LoadSearchPreset = False
        Exit Function
    End If
    
    ' Read file
    Set file = fso.OpenTextFile(filePath, 1, False, -1)
    fileContent = file.ReadAll
    file.Close
    
    ' Find preset in JSON
    Dim searchStr As String
    searchStr = """" & presetName & """: {"
    startPos = InStr(fileContent, searchStr)
    
    If startPos = 0 Then
        MsgBox "Preset '" & presetName & "' not found.", vbExclamation, "Preset Not Found"
        LoadSearchPreset = False
        Exit Function
    End If
    
    ' Extract preset JSON (find matching closing brace)
    startPos = InStr(startPos, fileContent, "{")
    endPos = FindMatchingBrace(fileContent, startPos)
    
    If endPos = 0 Then
        MsgBox "Error parsing preset file.", vbCritical, "Parse Error"
        LoadSearchPreset = False
        Exit Function
    End If
    
    presetJSON = Mid(fileContent, startPos, endPos - startPos + 1)
    
    ' Deserialize
    criteria = DeserializeCriteria(presetJSON)
    
    LoadSearchPreset = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error loading preset: " & Err.Description, vbCritical, "Load Error"
    LoadSearchPreset = False
End Function

Private Function FindMatchingBrace(text As String, startPos As Long) As Long
    Dim depth As Long
    Dim i As Long
    Dim char As String
    
    depth = 0
    
    For i = startPos To Len(text)
        char = Mid(text, i, 1)
        
        If char = "{" Then
            depth = depth + 1
        ElseIf char = "}" Then
            depth = depth - 1
            If depth = 0 Then
                FindMatchingBrace = i
                Exit Function
            End If
        End If
    Next i
    
    FindMatchingBrace = 0 ' Not found
End Function
```

#### 4. Add Preset Management UI

Add a button to the form to list all presets:

```vba
Private Sub btnManagePresets_Click()
    Dim fso As Object
    Dim file As Object
    Dim filePath As String
    Dim fileContent As String
    Dim presetList As String
    Dim presetName As String
    Dim startPos As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetPresetFilePath()
    
    If Not fso.FileExists(filePath) Then
        MsgBox "No presets saved yet.", vbInformation, "Manage Presets"
        Exit Sub
    End If
    
    ' Read file
    Set file = fso.OpenTextFile(filePath, 1, False, -1)
    fileContent = file.ReadAll
    file.Close
    
    ' Extract preset names (simple parsing)
    presetList = "Saved Presets:" & vbCrLf & vbCrLf
    
    startPos = 1
    Do
        startPos = InStr(startPos, fileContent, """")
        If startPos = 0 Then Exit Do
        
        startPos = startPos + 1
        Dim endQuote As Long
        endQuote = InStr(startPos, fileContent, """")
        
        If endQuote > 0 Then
            presetName = Mid(fileContent, startPos, endQuote - startPos)
            presetList = presetList & "• " & presetName & vbCrLf
            startPos = endQuote + 1
        End If
    Loop
    
    MsgBox presetList, vbInformation, "Manage Presets"
End Sub
```

---

## Priority 2: Persistent Search History (Estimated: 1 hour)

### Implementation Steps

#### 1. Add History File Functions

```vba
Private Const HISTORY_FILE_PATH As String = "\AppData\Roaming\OutlookEmailSearch\history.txt"
Private Const MAX_HISTORY_ITEMS As Integer = 10

Private Function GetHistoryFilePath() As String
    Dim userProfile As String
    userProfile = Environ("USERPROFILE")
    GetHistoryFilePath = userProfile & HISTORY_FILE_PATH
End Function

Private Sub SaveHistoryToFile()
    On Error Resume Next
    
    Dim fso As Object
    Dim file As Object
    Dim filePath As String
    Dim i As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    EnsurePresetDirectoryExists
    filePath = GetHistoryFilePath()
    
    ' Write history to file
    Set file = fso.CreateTextFile(filePath, True, True)
    
    For i = 1 To Me.cboSearchHistory.ListCount - 1 ' Skip first item (header)
        If i <= MAX_HISTORY_ITEMS Then
            file.WriteLine Me.cboSearchHistory.List(i)
        End If
    Next i
    
    file.Close
End Sub

Private Sub LoadHistoryFromFile()
    On Error Resume Next
    
    Dim fso As Object
    Dim file As Object
    Dim filePath As String
    Dim line As String
    Dim count As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetHistoryFilePath()
    
    If Not fso.FileExists(filePath) Then Exit Sub
    
    ' Read history from file
    Set file = fso.OpenTextFile(filePath, 1, False, -1)
    
    Me.cboSearchHistory.Clear
    Me.cboSearchHistory.AddItem "--- Select from recent searches ---"
    
    count = 0
    Do While Not file.AtEndOfStream And count < MAX_HISTORY_ITEMS
        line = file.ReadLine
        If Trim(line) <> "" Then
            Me.cboSearchHistory.AddItem line
            count = count + 1
        End If
    Loop
    
    file.Close
    
    Me.cboSearchHistory.ListIndex = 0
End Sub
```

#### 2. Update UserForm_Initialize

```vba
Private Sub UserForm_Initialize()
    ' ... existing code ...
    
    ' Load persistent search history
    LoadHistoryFromFile
End Sub
```

#### 3. Update SaveToSearchHistory

```vba
Private Sub SaveToSearchHistory(criteria As SearchCriteria)
    ' ... existing code to build historyText ...
    
    ' Add to dropdown
    If Me.cboSearchHistory.ListCount >= MAX_HISTORY_ITEMS + 1 Then
        Me.cboSearchHistory.RemoveItem Me.cboSearchHistory.ListCount - 1
    End If
    
    Me.cboSearchHistory.AddItem historyText, 1
    
    ' Save to file
    SaveHistoryToFile
End Sub
```

---

## Priority 3: Advanced Boolean Expression Parser (Estimated: 4 hours)

### Current Limitation
The current implementation handles simple "A AND B OR C" but doesn't properly parse nested parentheses like "(A OR B) AND (C OR D)".

### Implementation Approach

#### 1. Create Token Structure

```vba
Private Type Token
    TokenType As String ' "TERM", "AND", "OR", "NOT", "LPAREN", "RPAREN"
    Value As String
End Type
```

#### 2. Create Tokenizer Function

```vba
Private Function TokenizeQuery(query As String) As Collection
    Dim tokens As New Collection
    Dim i As Long
    Dim char As String
    Dim currentToken As String
    Dim inQuotes As Boolean
    Dim token As Token
    
    currentToken = ""
    inQuotes = False
    
    For i = 1 To Len(query)
        char = Mid(query, i, 1)
        
        If char = Chr(34) Then ' Quote
            inQuotes = Not inQuotes
            currentToken = currentToken & char
        
        ElseIf char = "(" And Not inQuotes Then
            If currentToken <> "" Then
                token.TokenType = "TERM"
                token.Value = Trim(currentToken)
                tokens.Add token
                currentToken = ""
            End If
            token.TokenType = "LPAREN"
            token.Value = "("
            tokens.Add token
        
        ElseIf char = ")" And Not inQuotes Then
            If currentToken <> "" Then
                token.TokenType = "TERM"
                token.Value = Trim(currentToken)
                tokens.Add token
                currentToken = ""
            End If
            token.TokenType = "RPAREN"
            token.Value = ")"
            tokens.Add token
        
        ElseIf char = " " And Not inQuotes Then
            If currentToken <> "" Then
                Dim upperToken As String
                upperToken = UCase(Trim(currentToken))
                
                If upperToken = "AND" Then
                    token.TokenType = "AND"
                    token.Value = "AND"
                    tokens.Add token
                ElseIf upperToken = "OR" Then
                    token.TokenType = "OR"
                    token.Value = "OR"
                    tokens.Add token
                ElseIf upperToken = "NOT" Then
                    token.TokenType = "NOT"
                    token.Value = "NOT"
                    tokens.Add token
                Else
                    token.TokenType = "TERM"
                    token.Value = currentToken
                    tokens.Add token
                End If
                
                currentToken = ""
            End If
        
        Else
            currentToken = currentToken & char
        End If
    Next i
    
    ' Add final token
    If currentToken <> "" Then
        token.TokenType = "TERM"
        token.Value = currentToken
        tokens.Add token
    End If
    
    Set TokenizeQuery = tokens
End Function
```

#### 3. Create Recursive Parser

```vba
Private Function ParseExpression(tokens As Collection, ByRef index As Long) As String
    Dim left As String
    Dim right As String
    Dim operator As String
    Dim token As Token
    
    ' Parse left side
    left = ParseTerm(tokens, index)
    
    ' Check for binary operator
    If index <= tokens.Count Then
        Set token = tokens(index)
        
        If token.TokenType = "AND" Or token.TokenType = "OR" Then
            operator = token.Value
            index = index + 1
            
            ' Parse right side
            right = ParseExpression(tokens, index)
            
            ' Combine with operator
            ParseExpression = "(" & left & " " & operator & " " & right & ")"
        Else
            ParseExpression = left
        End If
    Else
        ParseExpression = left
    End If
End Function

Private Function ParseTerm(tokens As Collection, ByRef index As Long) As String
    Dim token As Token
    Dim result As String
    
    If index > tokens.Count Then
        ParseTerm = ""
        Exit Function
    End If
    
    Set token = tokens(index)
    index = index + 1
    
    Select Case token.TokenType
        Case "NOT"
            ' Handle NOT operator
            result = "NOT " & ParseTerm(tokens, index)
            ParseTerm = result
        
        Case "LPAREN"
            ' Handle parenthesized expression
            result = ParseExpression(tokens, index)
            
            ' Expect closing paren
            If index <= tokens.Count Then
                Set token = tokens(index)
                If token.TokenType = "RPAREN" Then
                    index = index + 1
                End If
            End If
            
            ParseTerm = result
        
        Case "TERM"
            ' Convert term to DASL filter
            ParseTerm = TermToDASLFilter(token.Value)
        
        Case Else
            ParseTerm = ""
    End Select
End Function

Private Function TermToDASLFilter(term As String) As String
    ' Convert a single term to DASL filter
    ' This would call the existing logic for subject/body search
    
    ' Simplified version:
    Dim escaped As String
    escaped = EscapeSQL(term)
    
    TermToDASLFilter = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%" & escaped & "%'"
End Function
```

---

## Priority 4: Settings Dialog (Estimated: 3 hours)

### Implementation Steps

#### 1. Create New UserForm (frmSettings)

Controls needed:
- Label + ComboBox for Default Folder
- Label + TextBox for Result Limit (number)
- Label + ComboBox for Date Format (YYYY-MM-DD, MM/DD/YYYY, DD/MM/YYYY)
- Label + CheckBox for Auto-save History
- Button: Save Settings
- Button: Cancel

#### 2. Add Settings Storage Functions

```vba
' In ThisOutlookSession module
Private Type AppSettings
    DefaultFolder As String
    ResultLimit As Long
    DateFormat As String
    AutoSaveHistory As Boolean
End Type

Private Function LoadSettings() As AppSettings
    Dim settings As AppSettings
    Dim fso As Object
    Dim file As Object
    Dim filePath As String
    Dim fileContent As String
    
    ' Default values
    settings.DefaultFolder = "All Folders (including subfolders)"
    settings.ResultLimit = 500
    settings.DateFormat = "YYYY-MM-DD"
    settings.AutoSaveHistory = True
    
    ' Load from file if exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = Environ("USERPROFILE") & "\AppData\Roaming\OutlookEmailSearch\settings.txt"
    
    If fso.FileExists(filePath) Then
        Set file = fso.OpenTextFile(filePath, 1)
        
        Do While Not file.AtEndOfStream
            Dim line As String
            line = file.ReadLine
            
            If InStr(line, "=") > 0 Then
                Dim parts() As String
                parts = Split(line, "=")
                
                Select Case Trim(parts(0))
                    Case "DefaultFolder"
                        settings.DefaultFolder = Trim(parts(1))
                    Case "ResultLimit"
                        settings.ResultLimit = CLng(Trim(parts(1)))
                    Case "DateFormat"
                        settings.DateFormat = Trim(parts(1))
                    Case "AutoSaveHistory"
                        settings.AutoSaveHistory = CBool(Trim(parts(1)))
                End Select
            End If
        Loop
        
        file.Close
    End If
    
    LoadSettings = settings
End Function

Private Sub SaveSettings(settings As AppSettings)
    Dim fso As Object
    Dim file As Object
    Dim filePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    EnsurePresetDirectoryExists
    filePath = Environ("USERPROFILE") & "\AppData\Roaming\OutlookEmailSearch\settings.txt"
    
    Set file = fso.CreateTextFile(filePath, True)
    
    file.WriteLine "DefaultFolder=" & settings.DefaultFolder
    file.WriteLine "ResultLimit=" & settings.ResultLimit
    file.WriteLine "DateFormat=" & settings.DateFormat
    file.WriteLine "AutoSaveHistory=" & settings.AutoSaveHistory
    
    file.Close
End Sub
```

---

## Priority 5: Logging System (Estimated: 1.5 hours)

### Implementation Steps

#### 1. Add Logging Functions

```vba
' In ThisOutlookSession module
Private Sub WriteLog(message As String, Optional level As String = "INFO")
    On Error Resume Next
    
    Dim fso As Object
    Dim file As Object
    Dim filePath As String
    Dim timestamp As String
    Dim logEntry As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    EnsurePresetDirectoryExists
    filePath = Environ("USERPROFILE") & "\AppData\Roaming\OutlookEmailSearch\app.log"
    
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
    logEntry = timestamp & " [" & level & "] " & message
    
    ' Append to log file
    If fso.FileExists(filePath) Then
        Set file = fso.OpenTextFile(filePath, 8, False, -1) ' 8 = ForAppending
    Else
        Set file = fso.CreateTextFile(filePath, True, True)
    End If
    
    file.WriteLine logEntry
    file.Close
End Sub

Private Sub TrimLogFile()
    ' Keep only last 1000 lines to prevent huge log files
    On Error Resume Next
    
    Dim fso As Object
    Dim file As Object
    Dim filePath As String
    Dim lines As Collection
    Dim line As String
    Dim maxLines As Long
    
    maxLines = 1000
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = Environ("USERPROFILE") & "\AppData\Roaming\OutlookEmailSearch\app.log"
    
    If Not fso.FileExists(filePath) Then Exit Sub
    
    Set file = fso.OpenTextFile(filePath, 1, False, -1)
    Set lines = New Collection
    
    ' Read all lines
    Do While Not file.AtEndOfStream
        lines.Add file.ReadLine
    Loop
    file.Close
    
    ' Keep only last maxLines
    If lines.Count > maxLines Then
        Set file = fso.CreateTextFile(filePath, True, True)
        
        Dim i As Long
        For i = lines.Count - maxLines + 1 To lines.Count
            file.WriteLine lines(i)
        Next i
        
        file.Close
    End If
End Sub
```

#### 2. Add Log Calls to Key Functions

```vba
Public Sub ExecuteEmailSearch(criteria As SearchCriteria)
    WriteLog "Starting search - From: " & criteria.FromAddress & ", Terms: " & criteria.SearchTerms
    
    ' ... existing code ...
    
    WriteLog "Search completed in " & Format(searchTime, "0.00") & "s - Results: " & resultCount
End Sub

' Add error logging
On Error GoTo ErrorHandler

ErrorHandler:
    WriteLog "ERROR in ExecuteEmailSearch: " & Err.Description, "ERROR"
    MsgBox "Error during search: " & Err.Description, vbCritical
```

---

## Testing Checklist for New Features

### JSON Presets
- [ ] Save a simple preset
- [ ] Load the saved preset - verify all fields populate correctly
- [ ] Save a preset with special characters in name
- [ ] Save multiple presets
- [ ] Load non-existent preset - should show error
- [ ] Overwrite existing preset
- [ ] Manage presets - list should show all saved presets

### Persistent History
- [ ] Perform 5 searches
- [ ] Close and reopen Outlook
- [ ] Verify history dropdown shows previous searches
- [ ] Perform 15 searches (more than max)
- [ ] Verify only last 10 are kept

### Advanced Boolean Parser
- [ ] Search: `(urgent OR important) AND meeting`
- [ ] Search: `project AND (budget OR funding) NOT cancelled`
- [ ] Search: `((A OR B) AND C) OR (D AND E)`
- [ ] Search: `NOT spam`
- [ ] Verify results match expected logic

### Settings Dialog
- [ ] Open settings, change default folder
- [ ] Save and reopen search form - verify default folder changed
- [ ] Change result limit to 100
- [ ] Perform search with > 100 results - verify only 100 shown
- [ ] Toggle auto-save history off
- [ ] Perform search - verify history not saved

### Logging
- [ ] Perform successful search - check log file exists
- [ ] Perform search with error - check error logged
- [ ] Perform 100 searches - verify log trimmed to 1000 lines
- [ ] Check log file location in AppData\Roaming
- [ ] Verify timestamps are correct
- [ ] Verify log levels (INFO, ERROR) are correct

---

## Estimated Total Implementation Time

| Feature | Priority | Estimated Time |
|---------|----------|----------------|
| JSON Preset Storage | High | 2 hours |
| Persistent History | Medium | 1 hour |
| Advanced Boolean Parser | Low | 4 hours |
| Settings Dialog | Low | 3 hours |
| Logging System | Low | 1.5 hours |
| **Total** | | **11.5 hours** |

## Recommended Implementation Order

1. **Week 1**: JSON Preset Storage (most requested feature)
2. **Week 2**: Persistent History (quick win, high value)
3. **Week 3**: Logging System (helps debugging future features)
4. **Week 4**: Settings Dialog (nice to have)
5. **Week 5**: Advanced Boolean Parser (complex, low usage)

## Alternative: Use JSON Parser Library

Instead of manual JSON parsing, consider using a VBA JSON library:

- **VBA-JSON** by Tim Hall: https://github.com/VBA-tools/VBA-JSON
- Install as additional module
- Simplifies JSON serialization/deserialization
- Reduces implementation time by ~50%

Example with VBA-JSON library:

```vba
' After adding JsonConverter module

Private Function SerializeCriteria(criteria As SearchCriteria) As String
    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
    json.Add "FromAddress", criteria.FromAddress
    json.Add "SearchTerms", criteria.SearchTerms
    ' ... etc ...
    
    SerializeCriteria = JsonConverter.ConvertToJson(json)
End Function

Private Function DeserializeCriteria(json As String) As SearchCriteria
    Dim obj As Object
    Set obj = JsonConverter.ParseJson(json)
    
    Dim criteria As SearchCriteria
    criteria.FromAddress = obj("FromAddress")
    criteria.SearchTerms = obj("SearchTerms")
    ' ... etc ...
    
    DeserializeCriteria = criteria
End Function
```

This would reduce JSON implementation time from 2 hours to ~30 minutes.
