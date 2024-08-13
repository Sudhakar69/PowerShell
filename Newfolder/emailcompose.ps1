$queryId = "4shokbf63trv4xr3wth635rgothfv57ii5n3odhh7eq2s5tccycq"


# projectReportingQueryUri for Azure DevOps REST API

$projectReportingQueryUri = "https://dev.azure.com/apxinc/apx/_apis/git/repositories?api-version=6.0"
	$headers = @{
		Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryId)"))
	}

#$projectReportingQueryResult = Invoke-RestMethod -Uri $projectReportingQueryUri -Method Get -ContentType "application/json" -Headers $headers
# Make the API request to get a list of all repositories in the project
$repositoriesResponse = Invoke-RestMethod -Uri $projectReportingQueryUri -Headers $headers -Method Get

# Loop through each repository
foreach ($repository in $repositoriesResponse.value) 
{
	if ($repository.defaultBranch -eq "refs/heads/main") {
		$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=main&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get	
		Write-Host $repository.name " has MAIN as default branch" -BackgroundColor Yellow -ForegroundColor Black
	}
	elseif ($repository.defaultBranch -eq "refs/heads/master") {
		$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=master&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get	
		Write-Host $repository.name " has Master as Default branch" -BackgroundColor Green -ForegroundColor Black
		
	}
	else {
		$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get	
		Write-Host $repository.defaultBranch " is the default branch in " $repository.name -BackgroundColor Red -ForegroundColor Black
	}
    Write-Host "-------------------------------------------------------------------------------------------"
Write-Host "Sending out emails"

$email = "c-tgudise@xpansiv.com"
            Write-Host ("Sending to " + $email)
	foreach($commit in $commitsResponse.value){
		# Creating a custom object with properties
		#$workItemInfo = [PSCustomObject]@{
		#	RepositoryName = $repository.name
		#	BranchName = $commit.name
		#	Behindcount = $commit.behindcount
		#	Aheadcount = $commit.aheadcount
		#}
		#$workItemInfo

# compose and send out email messages to individuals who have booked to at least one task in the sprint

        
            # Display header
            $body = "<h2>$email</h2>"
            $body += "<h3><b>Branch Status Summary</b></h3>"

            $matrixHTML = "<table border='1' cellspacing='1' style='font-size:9px'><tr>`n"
            $matrixHTML += "<td valign='top'><b>Branch Name</b></td>`n"
            $matrixHTML += "<td valign='top' bgcolor='#ff0000' nowrap><b>Behind</b></td>`n"
            $matrixHTML += "<td valign='top' bgcolor='#90ee90'><b>Ahead</b></td>`n"
            $matrixHTML += "</tr>`n"
            $matrixHTML += "</table>`n"
            # Display summary table
            
            $body += "<tr><td align='right'>" + $commit.name + "</td>
                            <td align='right'>" + $commit.behindCount + "</td>
                            <td align='right'><b>" + $commit.aheadCount + "</b></td></tr>"
            $body += "</tbody></table>"

    
            $body += "<br><br><br><h3>Team Time Booking</h3>"
            $body += $matrixHTML
            $body += "<br><br>"
            

            $cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
        	    #Send-MailMessage -SmtpServer $SMTP_SERVER -Port 2525 -Credential $cred -To "agurvits@apx.com" -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml

            #}

        }
        $emailSubject = "Branch Status Summary "
            
            #$emailSubject += " (" + $email + ")"
        
            #if ($email.IndexOf("gurvits") -ge 0)
            #{
                Start-Sleep -s 15
		
                $SMTP_SERVER = "smtp.socketlabs.com"
                $SMTP_PORT = 25
                $SMTP_USERNAME = "c-tgudise@xpnsiv.com"
                $PWORD = ConvertTo-SecureString -String "Welcome123456" -AsPlainText -Force

                #$body | Out-file -Filepath "R:\vsts-scripts\projectreporting\temp\$($email).html"
                #$body | Out-file -Filepath "C:\_tmp\sprint-time-booking-emails.html" -Append
                #Write-Host ">>>> sending email to: " $email
                #$email = "agurvits@apx.com"
                
                #$bccEmailList = @("agurvits@apx.com","kliang@apx.com")
                
                $cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
        	    Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -From "c-tgudise@xpansiv.com" -Subject $emailSubject -Body $body -BodyAsHtml

                #
    }
       
Write-Host "Finished sending emails"
