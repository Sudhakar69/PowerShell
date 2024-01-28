$RGs=Get-content "C:\Users\gudisesu\OneDrive - Willis Towers Watson\Scripts\ResourceGroup.txt"
foreach($RG in $RGs)
{
	$pip=Get-AzPublicIpAddress -ResourceGroupName $RG
	$i=0
	if($i -lt $pip.count)
	{
		$ipconfig=$pip[$i].IpConfiguration
		if($ipconfig.PublicIpAddress -eq $null)
		{
			$bastionid=Get-content "C:\Users\gudisesu\OneDrive - Willis Towers Watson\Scripts\bastionid.txt"
			$bastion = Get-AzBastion -ResourceGroupName $RG
			$j=0
			if($j -lt $bastion.count)
			{			
				foreach($bid in $bastionid)
				{
					if($bid -eq $bastion[$j].id)
					{
				
						Remove-AzBastion -InputObject $bastion[$j] -force
						write-host "Deleting '$bastion[$j].Name'"
						Remove-AzPublicIpAddress -Name $pip[$i].Name -ResourceGroupName $RG -force
						write-host "Deleting '$pip[$i].Name'."
					}
				}
			$j++
			}
		} 
	$i++
	} 
}            