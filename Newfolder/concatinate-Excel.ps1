# Step 3: Concatenate and trim CVE Description
# Create a custom function to concatenate and trim a list of values
# function Remove-TrimUnique($values) 
# {
#     $uniqueValues = $values | Get-Unique
#     $concatenatedValues = $uniqueValues -join "|"
    
#     if ($concatenatedValues.Length -gt 200) {
#         $concatenatedValues = $concatenatedValues.Substring(0, 200)
#     }
    
#     return $concatenatedValues
# }

# # Step 3: Concatenate and trim "CVE Description" for records with the same "CVE ID" and "Hostname"
# $data | Group-Object -Property "Hostname" | ForEach-Object {
#     $uniqueCVE = $_.Group | Select-Object -ExpandProperty "Concatinated"
#     $concatenatedCVE = Remove-TrimUnique $uniqueCVE
#     $_.Group | ForEach-Object {
#         $_."CVE Description" = $concatenatedCVE
#     }
# }

# Step 4: Concatenate and trim "Remediation Details" for records with the same "CVE ID" and "Hostname"
# $data | Group-Object -Property "CVE ID" | ForEach-Object {
#     $uniqueRemediation = $_.Group | Select-Object -ExpandProperty "Remediation Details"
#     $concatenatedRemediation = Remove-TrimUnique $uniqueRemediation
#     $_.Group | ForEach-Object {
#         $_."Remediation Details" = $concatenatedRemediation
#     }
# }
# Define the path to the master Excel file
$excelPath = "C:\Users\c-tgudise\OneDrive - APX\APX\Alex\Book1.xlsx"
# Create a new Excel application
$excel = New-Object -ComObject Excel.Application

# Open the master Excel file
$workbook = $excel.Workbooks.Open($excelPath)
# Create a new sheet in the same workbook
$newSheet = $workbook.Sheets.Add()
$newSheet.Name = "NewSheet"

# Select the "Geek911-managed" sheet from the master Excel file
$sourceSheet = $workbook.Sheets.Item("Geek911-managed")

# Define the row from which you want to start copying data
$startRow = 5  # For example, starting from row 2

# Copy data from "Geek911-managed" sheet (CVE ID, Hostname, and Comments columns only)
$row = 2
foreach ($entry in $sourceSheet.UsedRange.Rows) {
    $newSheet.Cells.Item($row, 1).Value2 = $sourceSheet.Cells.Item($startRow, 3).Value2 
    $newSheet.Cells.Item($row, 2).Value2 = $sourceSheet.Cells.Item($startRow, 5).Value2
    $newSheet.Cells.Item($row, 3).Value2 = $sourceSheet.Cells.Item($startRow, 5).Value2 + " | " + $sourceSheet.Cells.Item($startRow, 3).Value2
    $newSheet.Cells.Item($row, 4).Value2 = $sourceSheet.Cells.Item($startRow, 16).Value2
    $row++
    $startRow++
}

# Insert a new column with the name "Concatenated" in between Hostname and Comments columns
# $newSheet.Cells.Item(1, 1).EntireColumn.Insert()
$newSheet.Cells.Item(1, 1).Value2 = "CVE ID"
$newSheet.Cells.Item(1, 2).Value2 = "HostName"
$newSheet.Cells.Item(1, 3).Value2 = "Concatenated"
$newSheet.Cells.Item(1, 4).Value2 = "Comments"

# Save the changes to the master Excel file
$workbook.Save()

# Close the workbook and Excel application
$workbook.Close()
$excel.Quit()

# Release Excel COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel

Write-Host "Data copied, concatenated, and inserted into the new sheet in the master Excel file."

