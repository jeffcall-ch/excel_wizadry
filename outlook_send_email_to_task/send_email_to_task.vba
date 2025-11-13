Sub CreateTaskFromEmailWithLink()
    Dim objMail As Outlook.MailItem
    Dim objTask As Outlook.TaskItem
    Dim taskDueDate As Date
    Dim priorityChoice As Integer
    Dim folderName As String
    Dim objNS As Outlook.NameSpace
    Dim objTasksFolder As Outlook.Folder
    Dim objTargetFolder As Outlook.Folder
    
    On Error GoTo ErrorHandler
    
    ' Check if an email is selected
    If Application.ActiveExplorer.Selection.Count > 0 Then
        If TypeName(Application.ActiveExplorer.Selection.Item(1)) = "MailItem" Then
            Set objMail = Application.ActiveExplorer.Selection.Item(1)
            
            ' Get the folder name where the email is located (immediate parent only)
            folderName = objMail.Parent.Name
            
            ' Get the Tasks folder
            Set objNS = Application.GetNamespace("MAPI")
            Set objTasksFolder = objNS.GetDefaultFolder(olFolderTasks)
            
            ' Check if a subfolder with the email folder name already exists
            Set objTargetFolder = Nothing
            On Error Resume Next
            Set objTargetFolder = objTasksFolder.Folders(folderName)
            On Error GoTo ErrorHandler
            
            ' Create the folder if it doesn't exist
            If objTargetFolder Is Nothing Then
                On Error Resume Next
                Set objTargetFolder = objTasksFolder.Folders.Add(folderName, olFolderTasks)
                On Error GoTo ErrorHandler
            End If
            
            ' Verify we have a target folder
            If objTargetFolder Is Nothing Then
                MsgBox "Failed to create or access task folder: " & folderName, vbCritical
                Exit Sub
            End If
            
            ' Show date picker form
            DatePickerForm.Show
            
            ' Check if user cancelled
            If DatePickerForm.Cancelled Then
                Unload DatePickerForm
                Exit Sub
            End If
            
            taskDueDate = DatePickerForm.SelectedDate
            priorityChoice = DatePickerForm.Priority
            Unload DatePickerForm
            
            ' Create a new task in the target folder
            Set objTask = objTargetFolder.Items.Add(olTaskItem)
            
            If objTask Is Nothing Then
                MsgBox "Failed to create task object", vbCritical
                Exit Sub
            End If
            
            objTask.Subject = objMail.Subject
            
            ' Add email metadata and body to task (plain text)
            objTask.Body = "From: " & objMail.SenderName & vbCrLf & _
                          "To: " & objMail.To & vbCrLf & _
                          IIf(objMail.CC <> "", "CC: " & objMail.CC & vbCrLf, "") & _
                          "Sent: " & Format(objMail.ReceivedTime, "dd/mm/yyyy hh:nn") & vbCrLf & _
                          "Subject: " & objMail.Subject & vbCrLf & _
                          String(50, "-") & vbCrLf & vbCrLf & _
                          objMail.Body
            
            objTask.DueDate = taskDueDate
            
            ' Set priority
            If priorityChoice = 2 Then
                objTask.Importance = olImportanceHigh
            Else
                objTask.Importance = olImportanceNormal
            End If
            
            ' Save first, then attach email
            objTask.Save
            
            ' Attach the original email to the task as an attachment
            Dim objAttachment As Outlook.Attachment
            On Error Resume Next
            Set objAttachment = objTask.Attachments.Add(objMail, olByValue)
            If Err.Number = 0 Then
                objTask.Save
            End If
            On Error GoTo ErrorHandler
            
        Else
            MsgBox "Please select an email.", vbExclamation
        End If
    Else
        MsgBox "Please select an email.", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating task: " & Err.Description & vbCrLf & "Error number: " & Err.Number, vbCritical
End Sub
