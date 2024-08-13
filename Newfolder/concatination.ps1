# Define the path to the master Excel file
$excelPath = "C:\Users\c-tgudise\OneDrive - APX\Testfiles\Book123.xlsx"
$csvFilePath = "C:\Users\c-tgudise\OneDrive - APX\CVE_output.csv"

# Load the CSV data
$data = Import-Csv -Path $csvFilePath

# Step 1: Update MachineDomain for specific Hostname
$data | ForEach-Object {
    if ($_.Hostname -eq "apxdev-usw-linuxbld") 
    {
        $_.MachineDomain = "apxna.com"
    }
}

# Step 2: Delete records with MachineDomain not equal to "apxna.com"
$data = $data | Where-Object { 
    $_.MachineDomain -eq "apxna.com" 
}


# Step 5: Sort by Severity and Hostname
$data = $data | Sort-Object -Property "HostnameDomain","Severity" | Select-Object -Unique

# Create a new Excel application
$excel = New-Object -ComObject Excel.Application
 
# Open the master Excel file
$workbook = $excel.Workbooks.Open($excelPath)
 
# Select the "RawData" sheet from the master Excel file
$rawDataSheet = $workbook.Sheets.Item("Raw Data")
 
# Get the headers (column names) from the CSV file
$headers = $data[0].PSObject.Properties.Name
# Create a new sheet in the same workbook
$newSheet = $workbook.Sheets.Add()
$newSheet.Name = "NewShaet"
 
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
# Determine the next available row for pasting data in the "RawData" sheet
$nextAvailableRow = $rawDataSheet.UsedRange.Rows.Count + 1
 
# Loop through the CSV data and paste it into the "RawData" sheet
foreach ($rowdata in $data ) {
    # Calculate the range of columns from Column A to AQ
    $columnsToCopy = $headers[0..42]
   
    # Loop through the columns and copy data from the CSV to the RawData sheet
    foreach ($columnName in $columnsToCopy) {
        # Find the corresponding column index in the RawData sheet
        $columnIndex = $rawDataSheet.UsedRange.Rows.Item(1).Find($columnName).Column
        $rawDataSheet.Cells.Item($nextAvailableRow, $columnIndex).Value2 = $rowdata.$columnName
    }
   
    #Increment the row counter
    $nextAvailableRow++
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
Get-Date