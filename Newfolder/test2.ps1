$SMTP_Password ="Welcome123456"


$SMTP_SERVER = "smtp.office365.com"
$SMTP_PORT = 587
$SMTP_USERNAME = "c-tgudise@xpansiv.com"
$UseSsl = $true

$SMTP_Client = New-Object System.Net.Mail.SmtpClient
$SMTP_Client.Host = $SMTP_SERVER
$SMTP_Client.Port = $SMTP_PORT
$SMTP_Client.EnableSsl = $UseSsl

$SMTP_Client.Credentials = New-Object System.Net.NetworkCredential($SMTP_USERNAME,$SMTP_Password)
$matrixHTML = "<tbody><table border='1' cellspacing='1' style='font-size:9px'><tr>`n"
$matrixHTML += "<td valign='top'><b>Branch Name</b></td>`n"
$matrixHTML += "<td valign='top' bgcolor='#ff0000' nowrap><b>Behind</b></td>`n"
$matrixHTML += "<td valign='top' bgcolor='#90ee90'><b>Ahead</b></td>`n"
$matrixHTML += "</tr>`n"
Write-Host ("Sending to " + $SMTP_USERNAME)
$queryId = "4shokbf63trv4xr3wth635rgothfv57ii5n3odhh7eq2s5tccycq"

$projectReportingQueryUri = "https://dev.azure.com/apxinc/apx/_apis/git/repositories?api-version=6.0"
	$headers = @{
		Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryId)"))
	}
$repositoriesResponse = Invoke-RestMethod -Uri $projectReportingQueryUri -Headers $headers -Method Get

	# Loop through each repository
foreach ($repository in $repositoriesResponse.value) 
{
	if ($repository.defaultBranch -eq "refs/heads/main") 
	{
		$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=main&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get	
	}
	elseif ($repository.defaultBranch -eq "refs/heads/master") 
	{
		$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=master&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get		
	}
	else
	{
		try {
			$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?api-version=7.1-preview.1"
			$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get		
		}
		catch {
			$commitsResponse = $_.Exception.message
		}
	}
	$body = "<h2> $SMTP_USERNAME </h2>"
    $body += "<h3>YBranch Status Summary</h3>"
    $body += $matrixHTML
	foreach($commit in $commitsResponse)
	{
		$body += "<tr><td align='right'>" + $commit.name + "</td>
                        <td align='right'>" + $commit.behindCount + "</td>
                        <td align='right'>" + $commit.aheadCount + "</td></tr>"
            
	}
	$body += "</tbody></table>`n"
}

$Message = New-Object System.Net.Mail.MailMessage
$Message.From = $SMTP_USERNAME
$Message.to.add("c-tgudise@xpansiv.com")
$Message.subject = "Branch status summary"
$Message.body = $body
$Message.IsBodyHtml = $true

	$SmtpClient.Send($Message)
	$SmtpClient.Dispose()
	$Message.Dispose()
