$SiteURL = ""
Install-Module -Name ImportExcel
Import-Module ImportExcel
function Get-DateType {
  param (
    [string]$dateparam,
    [string]$formate
  )
  if ([double]::TryParse($dateparam,[ref]"")) {
    [datetime]::FromOADate($dateparam).ToString($formate)
  }
  elseif ($dateparam -eq $null -or $dateparam -eq "") {
    $dateparam = "N/A"
  } 
  else {
    try {
      ([datetime]$dateparam).ToString($formate)
    }
    catch {
      $_.Exception.Message
      Write-Host $dateparam.GetType().Name " for the date is "$dateparam
    }
    
  }
  return
}
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
$SFEsheet = "Input - SFE"
#Iterate through all folders in the list
try {
  $sfe=$Folders.FieldValues|Where-Object {$_.FileLeafRef -match $SFEfolder}
  Get-PnPFile -ServerRelativeUrl $sfe.FileRef -Path $env:temp -FileName $SFEfolder -AsFile -Force
  $SFEExcelPath = Join-Path $env:temp $sfe.FileLeafRef
  $SFEParsedData = Import-Excel -Path $SFEExcelPath -WorksheetName $SFEsheet -NoHeader
}
catch { 
  Write-Output '** ERROR**'
  Write-Output  "Folder Name: $SFEFolder"
  Write-Output $_.Exception.Message
}

$Trades = @()

$FilteredSFEData = $SFEParsedData | Where-Object -FilterScript { $_.P2 -ne $null } | Where-Object -FilterScript { $_.P2 -notmatch "Time" }
$FilteredSFEData | ForEach-Object {
  $Trade = [PSCustomObject]@{
      time = Get-DateType -dateparam $_.P2 -formate "HH:mm:ss"
      date = Get-DateType -dateparam $_.P3 -formate "yyyy-MM-dd"
      code = $_.P4
      period = $_.P5
      term = $_.P6
      description = $_.P7
     
  }
  $Trades += $Trade
}
if ($Trades.Count -gt 0) {
  $SFECsvpath = Join-Path $env:TEMP $CsvSFEfolder
  Write-Host "SFE File is available in "$SFECsvpath
  $Blob1HT = @{
    File             = $SFECsvpath
    Container        = $container
    Blob             = "Black/$CsvSFEfolder"
    Context          = $storageContext
    StandardBlobTier = 'Hot'
    BlobType         = 'Block'
  }
  Set-AzStorageBlobContent @Blob1HT
}
else {
  Write-Host $SFECsvpath" does not exist"
}

$Trades |Export-Csv $SFECsvpath
try {
  $OTC=$Folders.FieldValues|Where-Object {$_.FileLeafRef -match $OTCFolder}
  Get-PnPFile -ServerRelativeUrl $OTC.FileRef -Path $env:temp -FileName $OTCFolder -AsFile -Force
  $OTCExcelPath = Join-Path $env:temp $OTC.FileLeafRef
  # $OTCsheet = "Input OTC"
  $OTCParsedData = Import-Excel -Path $OTCExcelPath -NoHeader
}
catch { 
  Write-Output '** ERROR**'
  Write-Output  "Folder Name: $OTCFolder"
  Write-Output $_.Exception.Message
}

$FilteredOTCData = $OTCParsedData | Where-Object -FilterScript { $_.P3 -ne $null } | Where-Object -FilterScript { $_.P4 -notmatch "Time" }
$OTCs = @()
$FilteredOTCData | ForEach-Object {
    $OTC = [PSCustomObject]@{
        # [datetime]$OTCdate = $_.P3
        'Deal Number'= $_.P1
        Type = $_.P2
        # date = $OTCdate.ToString("yyyy-MM-dd")
        date = Get-DateType -dateparam $_.P3 -formate "yyyy-MM-dd"
        Time = Get-DateType -dateparam $_.P4 -formate "HH:mm:ss"     
        
        
    }
    $OTCs += $OTC
}
if ($OTCs.Count -gt 0) {
  $OTCCsvpath = Join-Path $env:TEMP $CSVOTCFile
  $OTCs |Export-Csv $OTCCsvpath
  # upload a file to the default account (inferred) access tier
  $Blob1HT = @{
    File             = $OTCCsvpath
    Container        = $container
    Blob             = "OTC/$CSVOTCFile"
    Context          = $storageContext
    StandardBlobTier = 'Hot'
    BlobType         = 'Block'
  }
  Set-AzStorageBlobContent @Blob1HT

}
else {
  Write-Host $OTCCsvpath" does not exist"
}

try {
  $Green=$Folders.FieldValues|Where-Object {$_.FileRef -match $GreenFolder}
  Get-PnPFile -ServerRelativeUrl $Green.FileRef -Path $env:temp -FileName $Green.FileLeafRef -AsFile -Force
  $GreenExcelPath = Join-Path $env:temp $Green.FileLeafRef
  $GreenParsedData = Import-Excel -Path $GreenExcelPath -NoHeader
}
catch {
  Write-Output '** ERROR**'
  Write-Output  "Folder Name: $GreenFolder"
  Write-Output $_.Exception.Message
}
Write-Host "Checking Green Trades file" -BackgroundColor Blue -ForegroundColor Yellow
$FilteredGreenData = $GreenParsedData | Where-Object -FilterScript { $_.P3 -ne $null } | Where-Object -FilterScript { $_.P3 -notmatch "Time" }
$GreenTrades = @()
$FilteredGreenData | ForEach-Object {
    $GreenTrade = [PSCustomObject]@{
      # P1,P2,P3,P4,P5,P6,P7,P9,P10,P11,P12
        Number = $_.P1
        date = Get-DateType -dateparam $_.P2 -formate "yyyy-MM-dd"
        Time = Get-DateType -dateparam $_.P3 -formate "HH:mm:ss"        
        
        
    }
    $GreenTrades += $GreenTrade
}
if ($GreenTrades.Count -gt 0) {
  $GreenCsvpath = Join-Path $env:TEMP $CsvGreenFile
  $GreenTrades |Export-Csv $GreenCsvpath
    # upload a file to the default account (inferred) access tier
  $Blob1HT = @{
    File             = $OTCCsvpath
    Container        = $container
    Blob             = "Green/$CsvGreenFile"
    Context          = $storageContext
    StandardBlobTier = 'Hot'
    BlobType         = 'Block'
  }
  Set-AzStorageBlobContent @Blob1HT
}
else {
  Write-Host $GreenCsvpath" does not exist" 
}

Write-Host "Green Trade CSV file is available in "$GreenCsvpath -BackgroundColor Green









