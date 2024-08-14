$SiteURL = ""
# $CSVPath = "Folders.csv"
$SASUri =""
$regex = [System.Text.RegularExpressions.Regex]::Match($SASUri, '(?i)\/+(?<StorageAccountName>.*?)\..*\/(?<Container>.*)\?(?<SASToken>.*)')
$storageAccountName = $regex.Groups['StorageAccountName'].Value
$container = $regex.Groups['Container'].Value
$sasToken = $regex.Groups['SASToken'].Value
$storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -SasToken $sasToken
$date = Get-Date
$day = $date.Day -1
$month = $date.Month
$year = $date.Year
$yy =$date.ToString("yy")
$yyyy =$date.ToString("yyyy")
$mm =$date.ToString("MM")
$dd = $mm =$date.ToString("dd")
$ListName = "Shared%20Documents/"
$SFEfolder = "SFE Confo "+$day+"-"+$month+"-"+$year+".xlsm"
$CsvSFEfolder = "SFE Confo_"+$yyyy+"-"+$mm+"-"+$dd+".CSV"
$OTCFolder = "OTC Confo "+$month+"."+$yy+".xlsm"
$CSVOTCFile = "OTC Confo_"+$yyyy+"-"+$mm+"-"+$dd+".CSV"
$GreenFolder ="Renewable Review/Green Trade List.xlsx"
$CsvGreenFile = "Green Trades_"+$yyyy+"-"+$mm+"-"+$dd+".CSV"
Connect-PnPOnline -ClientId "" -CertificatePath "XSignals.pfx" -CertificatePassword (ConvertTo-SecureString -AsPlainText "" -Force) -Url $SiteURL -Tenant "hvbrokers.onmicrosoft.com"
#Get the List
$List = Get-PnPList -Identity $ListName
$Folders = Get-PnPListItem -List $List -PageSize 500
#Iterate through all folders in the list

$sfe=$Folders.FieldValues|Where-Object {$_.FileLeafRef -match $SFEfolder}
Get-PnPFile -ServerRelativeUrl $sfe.FileRef -Path $env:temp -FileName $SFEfolder -AsFile -Force
$SFEExcelPath = Join-Path $env:temp $sfe.FileLeafRef
$SFEsheet = "Input - SFE"
function Get-ExtractedSFEData {
    $Items= @("Time","Date","Code","Period","Term","Description","B/S (buySell)","Qty","Price Premium","Option Strike","Option Type","SFE Flat","SFE Monthly","SFE Peak","SFE Option","SFE Green/Carbon","Total Bro")
    $ExtractedSFEData = New-Object -TypeName PSObject
    foreach($Item in $Items){
        $ExtractedSFEData| Add-Member -MemberType NoteProperty -Name $Item -Value $null
    }
    return $ExtractedSFEData
    
}
function Get-ExtractedOTCData {
    $Items= @("Deal Number","Type","Date","Time","Node","Contract","Terms","Trading Interval","Price","Quantity","Strike Price","Option Type","Payment Date","Contract MWh","Total Contract MWh","Contract Value Notes","SWAP","CAP","SWAPTION","Seller Brokerage","Total Brokerage")
    $ExtractedOTCData = New-Object -TypeName PSObject
    foreach($Item in $Items){
        $ExtractedOTCData| Add-Member -MemberType NoteProperty -Name $Item -Value $null
    }
    return $ExtractedOTCData
    
}
Get-Process -ProcessName *Excel* |Stop-Process
Start-Sleep -s 10
[Net.ServicePointManager]::SecurityProtocol = "tls12"

$SFEpath
$excel =  new-object -comobject Excel.Application
# New-Object -Com "Excel.Application"
# $excel.Visible = $true
# $SFEWorkbook = $excel.workbooks.open($SFEpath)
$SFEWorkbook = $excel.workbooks.open($SFEExcelPath)
$SFEresourcesSheet = $SFEWorkbook.Worksheets.Item($SFEsheet)
$SFEresourcesSheet.UsedRange.Rows.Count
# $EmployeeSheet = $rsWorkbook.Worksheets.Item($sheet1)
$SFEAvailability = @()
for ($i = 6; $i -lt $SFEresourcesSheet.UsedRange.Rows.Count; $i++) {
    $ExtractedSFEData = Get-ExtractedSFEData
    $ExtractedSFEData.Time = $SFEresourcesSheet.Range("b$i").text
    $ExtractedSFEData.Date = $SFEresourcesSheet.Range("c$i").text
    $ExtractedSFEData.Code = $SFEresourcesSheet.Range("d$i").text
    $ExtractedSFEData.Period = $SFEresourcesSheet.Range("e$i").text
    $ExtractedSFEData.Term = $SFEresourcesSheet.Range("f$i").text
    $ExtractedSFEData.Description = $SFEresourcesSheet.Range("g$i").text
    $ExtractedSFEData.Qty = $SFEresourcesSheet.Range("j$i").text
    try {
        $ExtractedSFEData.'B/S (buySell)' = $SFEresourcesSheet.Range("i$i").text 
        $ExtractedSFEData.'Price Premium' = $SFEresourcesSheet.Range("k$i").text
        $ExtractedSFEData.'Option Strike' = $SFEresourcesSheet.Range("l$i").text
        $ExtractedSFEData.'Option Type' = $SFEresourcesSheet.Range("m$i").text
        $ExtractedSFEData.'SFE Flat' = $SFEresourcesSheet.Range("o$i").text
        $ExtractedSFEData.'SFE Monthly' = $SFEresourcesSheet.Range("p$i").text
        $ExtractedSFEData.'SFE Peak' = $SFEresourcesSheet.Range("q$i").text
        $ExtractedSFEData.'SFE Option' = $SFEresourcesSheet.Range("r$i").text
        $ExtractedSFEData.'SFE Green/Carbon' = $SFEresourcesSheet.Range("s$i").text
        $ExtractedSFEData.'Total Bro' = $SFEresourcesSheet.Range("t$i").text
        Write-Host $i":: Row Data collected for "$SFEfolder
    }
    catch {
        $_.Exception.Message
    }
    $SFEAvailability += $ExtractedSFEData
    
}
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel)
try {
    $excel.Quit()
}
catch {
    $_.Exception.Message
}
Write-Host "Coverting extracted data to CSV file"
$SFECsvpath = Join-Path $env:TEMP $CsvSFEfolder
Get-Process -ProcessName *Excel* |Stop-Process -Force
$SFEAvailability|Where-Object{$_.term -ne "#N/A"}|Export-Csv $SFECsvpath
Write-Host "CSV conversion for SFE file is completed" -BackgroundColor Yellow
Start-Sleep -s 10
[Net.ServicePointManager]::SecurityProtocol = "tls12"

$OTC=$Folders.FieldValues|Where-Object {$_.FileLeafRef -match $OTCFolder}
Get-PnPFile -ServerRelativeUrl $OTC.FileRef -Path $env:temp -FileName $OTCFolder -AsFile -Force
$OTCExcelPath = Join-Path $env:temp $OTC.FileLeafRef
$OTCsheet = "Input OTC"
$SFEpath
$OTCPath
$excel1 =  new-object -comobject Excel.Application
$OTCWorkbook = $excel1.workbooks.open($OTCExcelPath)
$OTCresourcesSheet = $OTCWorkbook.Worksheets.Item($OTCsheet)

$OTCAvailability = @()
# Columns needed: "Deal Number","Type","Date","Time","Node","Contract","Terms","Trading Interval","Price","Quantity","Strike Price","Option Type","Payment Date","Contract MWh","Total Contract MWh","Contract Value Notes","SWAP","CAP","SWAPTION","Seller Brokerage","Total Brokerage"
# Deal Number, Type, Date, Time, Buyer,Trader,Seller,Trader,Node, Contract, Terms, Trading Interval, Price, Quantity, Strike Price, Option Type, Payment Date, Contract MWh, Total Contract MWh, Contract Value Notes, SWAP, 
# CAP, SWAPTION, Seller Brokerage, Total Brokerage
# Columns to ignore/drop: Buyer, Trader (Column F), Seller, Trader (Column H), Column AE, Column AF
$OTCresourcesSheet.UsedRange.Rows.Count
for ($i = 6; $i -lt $OTCresourcesSheet.UsedRange.Rows.Count; $i++) {
    $ExtractedOTCData = Get-ExtractedOTCData
    $ExtractedOTCData.Type = $OTCresourcesSheet.Range("b$i").text
    $ExtractedOTCData.Date = $OTCresourcesSheet.Range("c$i").text
    $ExtractedOTCData.Time = $OTCresourcesSheet.Range("d$i").text
    $ExtractedOTCData.Node = $OTCresourcesSheet.Range("i$i").text
    $ExtractedOTCData.Contract = $OTCresourcesSheet.Range("j$i").text
    $ExtractedOTCData.Terms = $OTCresourcesSheet.Range("k$i").text
    $ExtractedOTCData.Price = $OTCresourcesSheet.Range("m$i").text
    $ExtractedOTCData.Quantity = $OTCresourcesSheet.Range("n$i").text
    $ExtractedOTCData.SWAP = $OTCresourcesSheet.Range("u$i").text
    $ExtractedOTCData.CAP = $OTCresourcesSheet.Range("v$i").text
    $ExtractedOTCData.SWAPTION = $OTCresourcesSheet.Range("w$i").text
    try {
        $ExtractedOTCData.'Deal Number' = $OTCresourcesSheet.Range("a$i").text 
        $ExtractedOTCData.'Trading Interval' = $OTCresourcesSheet.Range("l$i").text
        $ExtractedOTCData.'Strike Price' = $OTCresourcesSheet.Range("o$i").text
        $ExtractedOTCData.'Option Type' = $OTCresourcesSheet.Range("p$i").text
        $ExtractedOTCData.'Payment Date' = $OTCresourcesSheet.Range("q$i").text
        $ExtractedOTCData.'Contract MWh' = $OTCresourcesSheet.Range("r$i").text
        $ExtractedOTCData.'Total Contract MWh' = $OTCresourcesSheet.Range("s$i").text
        $ExtractedOTCData.'Contract Value Notes' = $OTCresourcesSheet.Range("t$i").text
        $ExtractedOTCData.'Seller Brokerage' = $OTCresourcesSheet.Range("ac$i").text
        $ExtractedOTCData.'Total Brokerage' = $OTCresourcesSheet.Range("ad$i").text
    Write-Host $i":: Row Data collected for "$OTCfolder
    }
    catch {
        $_.Exception.Message
    }
    $OTCAvailability += $ExtractedOTCData
    
}

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel1)
try {
    $excel1.Quit()
}
catch {
    $_.Exception.Message
}

Get-Process -ProcessName *Excel* |Stop-Process -Force
$OTCCsvpath = Join-Path $env:TEMP $CSVOTCFile
$OTCAvailability|Where-Object{$_.Type -ne "" -or $_.Type -ne $null}|Export-Csv $OTCCsvpath
Write-Host "CSV conversion for OTC file is completed" -BackgroundColor Yellow
$OTCsheet = "Spot and Forward"
$SFEpath
$OTCPath
<#
$excel1 =  new-object -comobject Excel.Application
$OTCWorkbook = $excel1.workbooks.open($OTCparentpath)
$OTCresourcesSheet = $OTCWorkbook.Worksheets.Item($OTCsheet)

$OTCAvailability = @()
# Columns needed: "Deal Number","Type","Date","Time","Node","Contract","Terms","Trading Interval","Price","Quantity","Strike Price","Option Type","Payment Date","Contract MWh","Total Contract MWh","Contract Value Notes","SWAP","CAP","SWAPTION","Seller Brokerage","Total Brokerage"
# Deal Number, Type, Date, Time, Buyer,Trader,Seller,Trader,Node, Contract, Terms, Trading Interval, Price, Quantity, Strike Price, Option Type, Payment Date, Contract MWh, Total Contract MWh, Contract Value Notes, SWAP, 
# CAP, SWAPTION, Seller Brokerage, Total Brokerage
# Columns to ignore/drop: Buyer, Trader (Column F), Seller, Trader (Column H), Column AE, Column AF
$OTCresourcesSheet.UsedRange.Rows.Count
for ($i = 6; $i -lt $OTCresourcesSheet.UsedRange.Rows.Count; $i++) {
    $ExtractedOTCData = Get-ExtractedOTCData
    $ExtractedOTCData.Type = $OTCresourcesSheet.Range("b$i").text
    $ExtractedOTCData.Date = $OTCresourcesSheet.Range("c$i").text
    $ExtractedOTCData.Time = $OTCresourcesSheet.Range("d$i").text
    $ExtractedOTCData.Node = $OTCresourcesSheet.Range("i$i").text
    $ExtractedOTCData.Contract = $OTCresourcesSheet.Range("j$i").text
    $ExtractedOTCData.Terms = $OTCresourcesSheet.Range("k$i").text
    $ExtractedOTCData.Price = $OTCresourcesSheet.Range("m$i").text
    $ExtractedOTCData.Quantity = $OTCresourcesSheet.Range("n$i").text
    $ExtractedOTCData.SWAP = $OTCresourcesSheet.Range("u$i").text
    $ExtractedOTCData.CAP = $OTCresourcesSheet.Range("v$i").text
    $ExtractedOTCData.SWAPTION = $OTCresourcesSheet.Range("w$i").text
    try {
        $ExtractedOTCData.'Deal Number' = $OTCresourcesSheet.Range("a$i").text 
        $ExtractedOTCData.'Trading Interval' = $OTCresourcesSheet.Range("l$i").text
        $ExtractedOTCData.'Strike Price' = $OTCresourcesSheet.Range("o$i").text
        $ExtractedOTCData.'Option Type' = $OTCresourcesSheet.Range("p$i").text
        $ExtractedOTCData.'Payment Date' = $OTCresourcesSheet.Range("q$i").text
        $ExtractedOTCData.'Contract MWh' = $OTCresourcesSheet.Range("r$i").text
        $ExtractedOTCData.'Total Contract MWh' = $OTCresourcesSheet.Range("s$i").text
        $ExtractedOTCData.'Contract Value Notes' = $OTCresourcesSheet.Range("t$i").text
        $ExtractedOTCData.'Seller Brokerage' = $OTCresourcesSheet.Range("ac$i").text
        $ExtractedOTCData.'Total Brokerage' = $OTCresourcesSheet.Range("ad$i").text
    Write-Host $i":: Row Data collected for "$OTCfolder
    }
    catch {
        $_.Exception.Message
    }
    $OTCAvailability += $ExtractedOTCData
    
}

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel1)
try {
    $excel1.Quit()
}
catch {
    $_.Exception.Message
}

Get-Process -ProcessName *Excel* |Stop-Process -Force
$OTCAvailability|Where-Object{$_.Type -ne "" -or $_.Type -ne $null}|Export-Csv "C:\Temp\OTC_output.csv"
Write-Host "CSV conversion for OTC file is completed" -BackgroundColor Yellow
#>