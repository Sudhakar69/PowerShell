$subs=Get-content "C:\Users\gudisesu\OneDrive - Willis Towers Watson\Scripts\Subscription.txt"
foreach($sub in $subs)
{
	Select-AzSubscription -Subscription $sub
	write-host "Selected Subscription '$sub'"
	$ro=Get-AzResourceGroup |Select-Object ResourceGroupName
	$RGs=Get-content "C:\Users\gudisesu\OneDrive - Willis Towers Watson\Scripts\groups.txt"
	foreach($RG in $RGs)
	{
		foreach($r in $ro)
		{
			if($r -eq $RG)
			{
		
				$pip=Get-AzPublicIpAddress -ResourceGroupName $RG
				$ipconfig=$pip.IpConfiguration
				if($ipconfig.PublicIpAddress -eq $null)
				{
					$bastion = Get-AzBastion -ResourceGroupName $RG
					$bastionid=Get-content "C:\Users\gudisesu\OneDrive - Willis Towers Watson\Scripts\bastiongroup.txt"
					foreach($bid in $bastionid)
					{
						if($bid -eq $bastion.id)
						{
				
							Remove-AzBastion -InputObject $bastion -force -asjob
							$name=$bastion.Name
							write-host "Deleting '$name'"
							Remove-AzPublicIpAddress -Name $pip.Name -ResourceGroupName $RG -force -asjob
							$pname=$pip.Name
							write-host "Deleting '$pname'"
						}
					}
				}
			}
		}
	}
}