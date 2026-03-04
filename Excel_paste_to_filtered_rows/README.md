# Paste To Filtered Rows (Excel VBA)

This macro lets you paste copied values only into the **visible cells** of a filtered selection, then restores the filter.

## Files

- `excel_paste_to_filtered_rows.vba` – source code of the macro.

## Install so it is available in **all Excel files**

Use the **Personal Macro Workbook** (`PERSONAL.XLSB`).

1. Open Excel.
2. If the **Developer** tab is hidden:
   - `File` → `Options` → `Customize Ribbon`
   - Enable `Developer`.
3. Open VBA editor: `Developer` → `Visual Basic` (or `Alt+F11`).
4. In Project Explorer, find `VBAProject (PERSONAL.XLSB)`.
   - If `PERSONAL.XLSB` does not exist yet:
     - `View` tab → `Macros` → `Record Macro...`
     - In **Store macro in**, choose **Personal Macro Workbook**.
     - Start and immediately stop recording.
     - Re-open `Alt+F11`.
5. In `PERSONAL.XLSB`, right-click `Modules` → `Insert` → `Module`.
6. Open `excel_paste_to_filtered_rows.vba`, copy all code, paste into the new module.
7. Save VBA project:
   - In VBA editor press `Ctrl+S`, or close Excel and choose **Save** when prompted for `PERSONAL.XLSB`.
8. Close and reopen Excel.

Now the macro is globally available in any workbook.

## Add button to the top Ribbon

1. `File` → `Options` → `Customize Ribbon`.
2. On the right side, create or select a custom group (for example under `Home` or `Data`).
3. On the left, set dropdown to `Macros`.
4. Select `PERSONAL.XLSB!PasteAndRestoreFilter` and click `Add >>`.
5. (Optional) Click `Rename...` to set a friendly label/icon.
6. Click `OK`.

You now have a Ribbon button at the top in every Excel session.

## Optional: assign a keyboard shortcut

1. `View` → `Macros` → `View Macros`.
2. Select `PERSONAL.XLSB!PasteAndRestoreFilter`.
3. Click `Options...` and assign a shortcut.

## Macro security notes

If macros are blocked:

- `File` → `Options` → `Trust Center` → `Trust Center Settings` → `Macro Settings`.
- Use your organization’s allowed setting (commonly **Disable with notification**).
- If needed, place files in a trusted location under `Trusted Locations`.

## Usage

1. Copy source cells (`Ctrl+C`).
2. Select target range (can be filtered; visible cells only are used).
3. Run `PasteAndRestoreFilter` from Ribbon.

The macro validates copied count vs visible target count and warns if they differ.
