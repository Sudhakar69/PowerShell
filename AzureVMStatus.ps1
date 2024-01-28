$excel=New-Object -comobject excel.application
$excel.Visible=$true
$ExcelWorkBook = $excel.Workbooks.Open("C:\Users\gudisesu\Documents\WTW-CRBIAAS-PROD.xlsx")
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("VMStatus")
$ExcelWorkSheet.Columns.Item(1).Rows.Item(1)="ServerName"
$ExcelWorkSheet.Columns.Item(2).Rows.Item(1)= "ResourceGroupName"
$ExcelWorkSheet.Columns.Item(3).Rows.Item(1)= "ComputerName"
$ExcelWorkSheet.Columns.Item(4).Rows.Item(1)= "VM Running Status"
$ExcelWorkSheet.Columns.Item(5).Rows.Item(1)= "OS Details"
$con=Get-Content "C:\Users\gudisesu\Documents\ServersList.txt" 
$i =2
if($i -le $con.count)
{
	foreach($ser in $con)
	{
		
		$vm=Get-AzVM -Status -Name $ser
		$vmstatus= Get-AzVM -Status -Name $ser -ResourceGroupName $vm.ResourceGroupName
		$status= $vmstatus.Statuses[1].DisplayStatus
		if($status -contains "running)
		{
 			$ExcelWorkSheet.Columns.Item(1).Rows.Item($i)=$vmstatus.Name
			$ExcelWorkSheet.Columns.Item(2).Rows.Item($i)= $vmstatus.ResourceGroupName
			$ExcelWorkSheet.Columns.Item(3).Rows.Item($i)= $vmstatus.ComputerName
			$ExcelWorkSheet.Columns.Item(4).Rows.Item($i)= "Running"
			$ExcelWorkSheet.Columns.Item(5).Rows.Item($i)= $vmstatus.OSName
		}
		else
		{
 			$ExcelWorkSheet.Columns.Item(1).Rows.Item($i)=$vmstatus.Name
			$ExcelWorkSheet.Columns.Item(2).Rows.Item($i)= $vmstatus.ResourceGroupName
			$ExcelWorkSheet.Columns.Item(3).Rows.Item($i)= $vmstatus.ComputerName
			$ExcelWorkSheet.Columns.Item(4).Rows.Item($i)= $vmstatus.Statuses[1].DisplayStatus
			$ExcelWorkSheet.Columns.Item(5).Rows.Item($i)= "NA"
			
           
                }
        
	$i++    
     }
}