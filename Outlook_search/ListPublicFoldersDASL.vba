' ============================================
' Module: ListPublicFoldersDASL
' Purpose: List all public folders/directories accessible via Outlook using DASL
' Output: Prints folder paths to Immediate Window (Debug.Print)
' ============================================

Option Explicit

Private entryCount As Integer

Public Sub ListPublicFoldersWithP5014()
    Dim ns As Outlook.NameSpace
    Dim publicRoot As Outlook.Folder

    entryCount = 0
    Set ns = Application.GetNamespace("MAPI")

    ' Try to get the root public folder
    On Error Resume Next
    Set publicRoot = ns.GetDefaultFolder(olPublicFoldersAllPublicFolders)
    On Error GoTo 0

    If publicRoot Is Nothing Then
        Debug.Print "No public folders found or no access."
        Exit Sub
    End If

    Debug.Print "=== PUBLIC FOLDERS WITH 'P-5014' ==="
    Debug.Print "Root: " & publicRoot.FolderPath

    ' Recursively list filtered public folders
    ListFoldersRecursiveP5014 publicRoot, ""

    Debug.Print "=== END OF LIST ==="
End Sub

Private Sub ListFoldersRecursiveP5014(parentFolder As Outlook.Folder, indent As String)
    Dim subFolder As Outlook.Folder
    For Each subFolder In parentFolder.Folders
        If InStr(1, subFolder.Name, "P-5014", vbTextCompare) > 0 Then
            Debug.Print indent & subFolder.FolderPath & " | DASL: " & subFolder.EntryID
            entryCount = entryCount + 1
            If entryCount >= 10 Then Exit Sub
        End If
        If entryCount < 10 Then
            ListFoldersRecursiveP5014 subFolder, indent & "    "
        Else
            Exit Sub
        End If
    Next
End Sub
