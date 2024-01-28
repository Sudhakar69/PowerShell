$excel=New-Object -comobject excel.application
$excel.Visible=$true
$ExcelWorkBook = $excel.Workbooks.Open("C:\Users\gudisesu\Documents\IPAddress.xlsx")
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Tabelle1")
$ExcelWorkSheet.Columns.Item(1).Rows.Item(1)="ServerName"
 $ExcelWorkSheet.Columns.Item(2).Rows.Item(1)= "IPAddress"
$con=Get-Content "C:\Users\gudisesu\Documents\ServersList.txt" 
$i=2
if($i -le $con.count){
foreach($ser in $con){
      $entry = [System.Net.Dns]::GetHostEntry($Ser)
             $entry.AddressList                  
                     
                         
                         [array]$x = Test-Connection -Delay 15 -ComputerName $ser -Count 1 -ErrorAction SilentlyContinue
                         
                         $ExcelWorkSheet.Columns.Item(1).Rows.Item($i) = $ser
                         $ExcelWorkSheet.Columns.Item(2).Rows.Item($i) = $x[0].IPV4Address.IPAddressToString
                         
                     }
        
     $i++    
     }