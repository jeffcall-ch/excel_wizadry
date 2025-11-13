# Outlook Email to Task Converter

Convert selected Outlook emails into tasks with automatic organization, date selection, and priority settings.

## Features

✅ **Automatic Task Organization** - Tasks are created in folders matching the email's folder name  
✅ **Date Picker** - User-friendly dropdown calendar for selecting due dates  
✅ **Priority Toggle** - Switch between Normal and High priority with a button  
✅ **Email Attachment** - Original email is attached to the task for easy reference  
✅ **Complete Email Metadata** - Preserves From, To, CC, Sent date, and full email body  
✅ **Error Handling** - Robust error checking with informative messages  

## Installation Instructions

### Step 1: Open Outlook VBA Editor
1. Open Microsoft Outlook
2. Press `Alt + F11` to open the VBA Editor
3. In the Project Explorer, find your VBA project (usually "Project1" or "VbaProject.OTM")

### Step 2: Import the Main Macro
1. In VBA Editor, go to `File` → `Import File...`
2. Navigate to: `c:\Users\szil\Repos\excel_wizadry\outlook_send_email_to_task\`
3. Select `send_email_to_task.vba`
4. Click `Open`

### Step 3: Create the Date Picker UserForm

#### A. Create the Form
1. In VBA Editor, right-click on your project → `Insert` → `UserForm`
2. Press `F4` to open the Properties window
3. Set these properties:
   - **Name**: `DatePickerForm`
   - **Caption**: `Select Due Date`
   - **Height**: `240`
   - **Width**: `280`

#### B. Add Controls

Make sure the **Toolbox** is visible (`View` → `Toolbox`). Add these controls:

**Labels (4 total):**
- Name: `lblMonth`, Caption: `Month:`, Position: Left=12, Top=12
- Name: `lblDay`, Caption: `Day:`, Position: Left=12, Top=42
- Name: `lblYear`, Caption: `Year:`, Position: Left=12, Top=72
- Name: `lblPreview`, Caption: `Selected:`, Position: Left=12, Top=108, Width: 240, Font: Bold

**ComboBoxes (3 total):**
- Name: `cboMonth`, Position: Left=60, Top=12, Width: 180
- Name: `cboDay`, Position: Left=60, Top=42, Width: 180
- Name: `cboYear`, Position: Left=60, Top=72, Width: 180

**ToggleButton (1 total):**
- Name: `tglPriority`, Caption: `Normal Priority`, Position: Left=12, Top=138, Width: 228, Height: 24

**CommandButtons (3 total):**
- Name: `btnToday`, Caption: `Today`, Position: Left=12, Top=174, Width: 60
- Name: `btnOK`, Caption: `OK`, Position: Left=84, Top=174, Width: 60
- Name: `btnCancel`, Caption: `Cancel`, Position: Left=156, Top=174, Width: 84

#### C. Add Code to UserForm
1. Double-click the UserForm to open its code window
2. Open `DatePickerForm_Code.bas` in a text editor
3. Copy **all** the code from the file
4. Paste it into the UserForm's code window (replace any existing code)

### Step 4: Enable Macros
1. In Outlook, go to `File` → `Options` → `Trust Center` → `Trust Center Settings`
2. Click `Macro Settings`
3. Select **"Notifications for digitally unsigned macros"** (recommended)
4. Click `OK` and restart Outlook

### Step 5: Add to Quick Access Toolbar (Optional but Recommended)

**Method 1: Quick Access Toolbar**
1. Click the dropdown arrow on the Quick Access Toolbar (top-left corner)
2. Select `More Commands...`
3. In "Choose commands from" dropdown, select `Macros`
4. Find `Project1.CreateTaskFromEmailWithLink`
5. Click `Add >>`
6. Click `Modify...` to choose an icon (suggest a task/flag icon)
7. Click `OK`

**Method 2: Custom Ribbon Button**
1. Right-click the Ribbon → `Customize the Ribbon...`
2. Create a new tab or group
3. Add the macro from the left list (choose Macros)
4. Customize icon and name
5. Click `OK`

## Usage

1. **Select an email** in Outlook
2. **Click your Quick Access button** (or run from Developer → Macros)
3. **Select the due date** using the dropdown menus
   - Or click "Today" button for today's date
4. **Toggle priority** if you want High priority (red flag)
5. **Click OK**
6. Task is created automatically!

## How It Works

### Task Organization
- Email in **"Inbox"** → Task created in **"Inbox"** list
- Email in **"Projects"** → Task created in **"Projects"** list
- Email in **"01 Projects\02 Medworth"** → Task created in **"02 Medworth"** list

The script creates task folders (lists) automatically if they don't exist.

### Task Content

Each task includes:
```
From: Sender Name
To: Recipient1; Recipient2
CC: CC Recipient (if any)
Sent: 13/11/2025 09:30
Subject: Email Subject
--------------------------------------------------

[Full email body text...]
```

**Plus:** The original email is attached as a `.msg` file that you can click to open.

### Priority Levels
- **Normal Priority** - Default, no special marking
- **High Priority** - Red exclamation mark (!) in task list

## File Structure

```
outlook_send_email_to_task/
├── send_email_to_task.vba          # Main macro code
├── DatePickerForm_Code.bas         # UserForm code (copy into UserForm)
├── DatePickerForm.frm              # Form definition (reference only)
├── DatePickerForm.frx              # Form binary (reference only)
└── README.md                       # This file
```

## Troubleshooting

### "Compile error: User-defined type not defined"
- Make sure you're running this in **Outlook VBA**, not Excel VBA
- The code requires the Outlook object library

### "Can't find DatePickerForm"
- Ensure the UserForm is named exactly `DatePickerForm` (check Name property, not Caption)
- Verify the UserForm code is properly pasted

### "Object variable or With block variable not set"
- Make sure you've selected an email before running the macro
- Try restarting Outlook and reimporting the code

### Date picker shows wrong date format
- The format is `dd/mm/yyyy hh:nn` (European format)
- Edit line 71 in the code to change format if needed

### Task created but no email attachment
- Check if the task was saved successfully
- Open the task and look for a paperclip icon at the top
- The attachment should be named with the email subject

### Macro doesn't appear in the list
- Make sure macros are enabled (see Step 4)
- Check that the macro name is exactly `CreateTaskFromEmailWithLink`
- Restart Outlook after importing

## Customization

### Change Date Format
Edit line 71 in `send_email_to_task.vba`:
```vb
"Sent: " & Format(objMail.ReceivedTime, "dd/mm/yyyy hh:nn") & vbCrLf & _
```
Change `"dd/mm/yyyy hh:nn"` to your preferred format (e.g., `"mm/dd/yyyy hh:mm"` for US format)

### Change Default Priority
Edit line 14 in `DatePickerForm_Code.bas`:
```vb
Priority = 1 ' Default to Normal priority
```
Change to `Priority = 2` for High priority as default

### Modify Task Body Format
Edit lines 68-75 in `send_email_to_task.vba` to customize what information appears in the task body.

## System Requirements

- Microsoft Outlook 2010 or newer
- Windows operating system
- VBA support enabled in Outlook

## Notes

- Tasks are created in your default Tasks folder structure
- The original email must remain in your mailbox for the attachment to stay accessible
- Task folders are created under the main Tasks folder, not as separate lists
- The script only works with `MailItem` objects (not meeting requests or other item types)

## Version History

**v1.0 (November 2025)**
- Initial release
- Date picker with dropdown selection
- Priority toggle button
- Automatic task folder organization
- Email attachment support
- Complete error handling

---

**Created by:** Laszlo Sziva  
**Date:** November 13, 2025  
**Repository:** excel_wizadry/outlook_send_email_to_task
