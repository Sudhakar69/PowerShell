# $path = "Leave Tracker and On Call.xlsx"
[Net.ServicePointManager]::SecurityProtocol = "tls12"
function Get-EmployeeAvailability {
    $EmployeeAvailability = New-Object -TypeName PSObject
    $EmployeeAvailability| Add-Member -MemberType NoteProperty -Name Name -Value $null
    $EmployeeAvailability| Add-Member -MemberType NoteProperty -Name Type -Value $null
    $EmployeeAvailability| Add-Member -MemberType NoteProperty -Name Date -Value $null
    $EmployeeAvailability| Add-Member -MemberType NoteProperty -Name Hours -Value $null
    return $EmployeeAvailability
    
}

$RP_FILEPATH = ""
$sheet ="Time Off"
$excel = new-object -com excel.application
# $excel.Visible = $true
$rsWorkbook = $excel.workbooks.open($RP_FILEPATH)
$resourcesSheet = $rsWorkbook.Worksheets.Item($sheet)
$Availability = @()
for ($i = 2; $i -lt $resourcesSheet.UsedRange.Rows.Count; $i++) {
    $name = $resourcesSheet.Range("a$i").text
    $Type= $resourcesSheet.Range("b$i").text
    $date = $resourcesSheet.Range("c$i").text
    $Hours =$resourcesSheet.Range("d$i").text
    # Write-Host $name "::" $Type "::" $date "::" $Hours
        $EmployeeAvailability = Get-EmployeeAvailability
        $EmployeeAvailability.Name = $name
        $EmployeeAvailability.Type = $Type
        $EmployeeAvailability.Date = $date
        $EmployeeAvailability.Hours = $Hours
        $Availability += $EmployeeAvailability
    
}
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel)
$Availability |Format-Table
