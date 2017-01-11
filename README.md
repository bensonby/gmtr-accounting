# GMT Research Accounting Screen Excel Spreadsheet

## Initial Setup (for dev)

- In the Excel file, press `Alt+F11`to open the VBA Editor
- Paste the content of `ThisWorkbook.bas` into the `ThisWorkbbook` item
- Go to `File -> Import File...`, import `CheckExpiry.bas`. (Or just drag the file in)

## Modifying the spreadsheet (for user)

- Important: Before saving the spreadsheet to upload to S3, make the main worksheet (Outputs) hidden; and leave an arbitrary worksheet there.
  - When a client downloads the spreadsheet from the Internet (S3), Excel will, by default, disallow the execution of any vba code for security reason. Therefore the client can see the spreadsheet as is.
  - It means that the code can only hide the main worksheet when the Excel has been saved at the local harddisk, or when the client selected "Enable Editing" after they download and open the spreadsheet.
- To manipulate sheets (add sheets / unhide data sheets), you need to unprotect the Workbook: Review -> Protect Workbook (Keyboard shortcut: Alt + T + P + W)
  - You need to re-protect the workbook structure after unprotecting it
  - Also please note that the password for protecting the workbook is also written in the VBA code because it needs to toggle the protection for hiding/unhiding worksheets.
  - If you want to change the password for protecting the workbook, you need to change the password in the VBA code as well
- You can move the cells of the code freely on the "Output" worksheet freely by Cut and Paste. Just make sure the cell name CODE_INPUT is preserved. The vba code will read this name.
- VBA Codes are written in both "Microsoft Excel Objects" -> "ThisWorkbook" and "Modules" -> "CheckExpiry".
