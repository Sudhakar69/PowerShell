# JB: Test of loading xlsm and serializing to a CSV
# Requires ImportExcel module
# Install-Module -Name ImportExcel
# Import-Module ImportExcel
$parentpath = ""
$date = Get-Date
$day = $date.Day
$month = $date.Month
$year = $date.Year
$SFEfolder = "SFE%20Confo%209-"+$month+"-"+$year+".xlsm"
$OTCFolder = "OTC Confo "+$month+"."+$date.ToString("yy")+".xlsm"
$SFEpath = $parentpath+$SFEfolder
$OTCPath =Join-Path $parentpath -ChildPath $OTCFolder
$SFEpath
$OTCPath
$SFEData = Import-Excel -Path $SFEpath -NoHeader # -WorksheetName - defaults to first sheet
$OTCData = Import-Excel -Path $OTCPath -NoHeader

$FilteredSFEData = $SFEData | Where-Object -FilterScript { $_.P2 -ne $null } | Where-Object -FilterScript { $_.P2 -notmatch "Time" }
Write-Host "we are in Filtered "$SFEfolder
$FilteredOTCData = $OTCData | Where-Object -FilterScript { $_.P2 -ne $null } | Where-Object -FilterScript { $_.P3 -notmatch "Time" }
$Trades = @()
Write-Host "we are in "$SFEfolder
$FilteredSFEData | ForEach-Object {
    $Trade = [PSCustomObject]@{
        time = [DateTime]::FromOADate($_.P2).ToString('HH:mm:ss')
        date = [DateTime]::FromOADate($_.P3).ToString('yyyy-MM-dd')
        code = $_.P4
        period = $_.P5
        term = $_.P6
        description = $_.P7
        #client = $_.P8
        buySell = $_.P9
        qty = $_.P10
        pricePremium = $_.P11
        optionStrike = $_.P12
        optionType = $_.P13
        #trader = $_.P14
        sfeFlat = $_.P15
        sfeMonthly = $_.P16
        sfePeak = $_.P17
        sfeOption = $_.P18
        sfeGreenCarbon = $_.P19
        totalBro = $_.P20
    }
    $Trades += $Trade
}
$Trades | Export-Csv -Path "Trades.csv"
$OTCs = @()
# Columns needed: Deal Number, Type, Date, Time, Node, Contract, Terms, Trading Interval, Price, Quantity, Strike Price, Option Type, Payment Date, Contract MWh, Total Contract MWh, Contract Value Notes, SWAP, CAP, SWAPTION, Seller Brokerage, Total Brokerage
# Deal Number, Type, Date, Time, Buyer,Trader,Seller,Trader,Node, Contract, Terms, Trading Interval, Price, Quantity, Strike Price, Option Type, Payment Date, Contract MWh, Total Contract MWh, Contract Value Notes, SWAP, 
# CAP, SWAPTION, Seller Brokerage, Total Brokerage
# Columns to ignore/drop: Buyer, Trader (Column F), Seller, Trader (Column H), Column AE, Column AF
Write-Host "we are in "$OTCFolder
$FilteredOTCData | ForEach-Object {
    $OTC = [PSCustomObject]@{
        'Deal Number'= $_.P1
        Type = $_.P2
        Time = [DateTime]::FromOADate($_.P3).ToString('HH:mm:ss')
        date = [DateTime]::FromOADate($_.P4).ToString('yyyy-MM-dd')
        # Buyer = $_.P5
        # Trader = $_.P6
        # Seller = $_.P7
        #client = $_.P8
        # Trader = $_.P9
        Node = $_.P10
        Contract = $_.P11
        Terms = $_.P12
        'Trading Interval' = $_.P13
        Price = $_.P14
        Quantity = $_.P15
        'Strike Price' = $_.P16
        'Option Type' = $_.P17
        'Payment Date' = $_.P18
        'Contract MWh' = $_.P19
        'Total Contract MWh' = $_.P20
        'Contract Value Notes' = $_.P21
        'SWAP' = $_.P22
        CAP = $_.P23
        SWAPTION = $_.P24
        'Seller Brokerage' = $_.P25
        'Total Brokerage' = $_.P26
    }
    $OTCs += $OTC
}
$OTCs | Export-Csv -Path "OTCs.csv"