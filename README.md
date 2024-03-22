# excel-copy

# Define paths to the source and template Excel files
$sourceFilePath = "C:\Path\To\Source\File.xlsx"
$templateFilePath = "C:\Path\To\Template\File.xlsx"

# Create an instance of Excel
$excel = New-Object -ComObject Excel.Application

# Open the source Excel file
$sourceWorkbook = $excel.Workbooks.Open($sourceFilePath)
$sourceWorksheet = $sourceWorkbook.Sheets.Item(1) # Assuming the data is on the first sheet

# Select and copy the data from the source workbook
$sourceRange = $sourceWorksheet.UsedRange
$sourceRange.Copy() | Out-Null

# Open the template Excel file
$templateWorkbook = $excel.Workbooks.Open($templateFilePath)
$templateWorksheet = $templateWorkbook.Sheets.Item(1) # Assuming you want to paste the data into the first sheet

# Select the destination cell in the template workbook and paste the data
$templateRange = $templateWorksheet.Cells.Item(1, 1) # Change the row and column as needed
$templateRange.PasteSpecial(-4163) # Paste values only, adjust the value as needed

# Save and close the template workbook
$templateWorkbook.Save()
$templateWorkbook.Close()

# Close the source workbook without saving changes
$sourceWorkbook.Close()

# Quit Excel
$excel.Quit()

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null



Make sure to adjust the paths ($sourceFilePath and $templateFilePath) to your source and template Excel files, respectively. Additionally, you may need to adjust the sheet numbers and cell references as needed based on your specific scenario.

Save this script with a .ps1 extension, and then you can run it using PowerShell. This script will open Excel, copy data from the source file, paste it into the template file, save the changes, and then close Excel.
