param ([string]$info="PLEASE PROVIDE PASSWORD AS AN ARGUMENT TO THIS SCRIPT")


Add-Type -AssemblyName System.Web

$scriptDirectory = ($pwd.path | Split-Path -Parent)
$commonLibPath = (Join-Path (Join-Path $scriptDirectory -ChildPath "common") -ChildPath "common.ps1")
. $commonLibPath

$infoLibPath = (Join-Path (Join-Path $scriptDirectory -ChildPath "common") -ChildPath "info.ps1")
. $infoLibPath

$infoLibPath = (Join-Path $PSScriptRoot -ChildPath "project_reporting_common.ps1")
. $infoLibPath

$SMTP_SERVER = "smtp.socketlabs.com"
$SMTP_PORT = 25
$SMTP_USERNAME = "server4507"
$PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force
$email = "c-tgudise@xpansiv.com"
$emailSubject ="Sample mail"

#$SMTP_Client.Credentials = New-Object System.Net.NetworkCredential($SMTP_USERNAME,$SMTP_Password)
# Render header row
#$epicIds = ($epicsByEpicId.Keys | Sort-Object)
$matrixHTML = "<tbody><table border='1' cellspacing='1' style='font-size:9px'><tr>`n"
$matrixHTML += "<td valign='top'><b>Branch Name</b></td>`n"
$matrixHTML += "<td valign='top' bgcolor='#ff0000' nowrap><b>Behind</b></td>`n"
$matrixHTML += "<td valign='top' bgcolor='#90ee90'><b>Ahead</b></td>`n"
$matrixHTML += "</tr>`n"
#$matrixHTML += "</table>`n"
# add resources who have not booked any time in VSTS
$email = "c-tgudise@xpansiv.com"

# Define email settings

Write-Host ("Sending to " + $SMTP_USERNAME)
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
		#Write-Host $repository.name " has MAIN as default branch" -BackgroundColor Yellow -ForegroundColor Black
		}
		elseif ($repository.defaultBranch -eq "refs/heads/master") {
			$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=master&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
			$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get	
		#Write-Host $repository.name " has Master as Default branch" -BackgroundColor Green -ForegroundColor Black
		
		}
		else {
			try{
				$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?api-version=7.1-preview.1"
				$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get	
				#Write-Host $repository.defaultBranch " is the default branch in " $repository.name -BackgroundColor Red -ForegroundColor Black
				return $commitsResponse
			}
			catch{
				$body += "<td>" + $($_.Exception.message) +"</td>"
				return $commitsResponse
			}
		}
		

	
            # Display header
            $body = "<h2> $SMTP_USERNAME </h2>"
            $body += "<h3>YBranch Status Summary</h3>"
            $body += $matrixHTML
            
	foreach($commit in $commitsResponse.value){
		# Creating a custom object with properties
		$body += "<tr><td align='right'>" + $commit.name + "</td>
                        <td align='right'>" + $commit.behindCount + "</td>
                        <td align='right'>" + $commit.aheadCount + "</td></tr>"
							}
	$body += "</tbody></table>"
}
	
Start-Sleep -s 15	
    
$cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml

        
            

            # Display table of projects
            #$body += "<h3>Breakdown by Project</h3>"
            