# param ([string]$info="PLEASE PROVIDE PASSWORD AS AN ARGUMENT TO THIS SCRIPT")

<#

Project Reporting script for VSTS

This script sends emails to all developers, testers, and product managers notifying them
of pending branches for merging and send a gentle reminder to the Application owner and respective leads to say hey these are Ahead and Behind Count for these particular branches


Author: Taraka Rama Gudise <c-tgudise@xpansiv.com>
Reviewer: Alex Gurvits <agurvits@apx.com>

#>
# Add-Type -AssemblyName System.Web
# $scriptDirectory = ($PSScriptRoot | Split-Path -Parent)
# $commonLibPath = (Join-Path (Join-Path $scriptDirectory -ChildPath "common") -ChildPath "common.ps1")
# . $commonLibPath
# $infoLibPath = (Join-Path $PSScriptRoot -ChildPath "project_reporting_common.ps1")
# . $infoLibPath



# $matrixHTML = "<tbody><table border='1' cellspacing='1' style='font-size:9px;border:1px solid black;border-collapse: collapse'><tr>`n"
# $matrixHTML += "<td valign='top'><b>Branch Name</b></td>`n"
# $matrixHTML += "<td valign='top' bgcolor='#ff0000' nowrap><b>Behind Count</b></td>`n"
# $matrixHTML += "<td valign='top' bgcolor='#90ee90'><b>Ahead Count</b></td>`n"
# $matrixHTML += "</tr>`n"

# $BranchSummary = Get-BranchSummary

$queryId = "4shokbf63trv4xr3wth635rgothfv57ii5n3odhh7eq2s5tccycq"
$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryId)"))
}
$projectReportingQueryUri = "https://dev.azure.com/apxinc/apx/_apis/git/repositories?api-version=6.0"
$repositoriesResponse = Invoke-RestMethod -Uri $projectReportingQueryUri -Headers $headers -Method Get

$repositories= @("account-service-model-service",
"acr",
"act",
"alerts",
"api-management",
"apx-jwt-auth-server",
"art",
"billing",
"billing-presentment",
"car",
"climateforward",
"connector-host",
"contract",
"edb",
"event-detector",
"external-asm",
"file-registry",
"grid",
"issuance",
"kafka-snowflake-bridge",
"legalentity",
"location",
"marketsuite",
"meter-data-telemetry-service",
"mirecs",
"multitenant",
"nar",
"ncrets",
"nepoolgis",
"notification",
"notification-callback-receiver",
"nygats",
"oo-edw-dbt",
"optimal-dataroom",
"optimal-event-pusher",
"optimal-external-api-gateway",
"optimal-gateway",
"optimal-ui",
"ownership",
"pm-dispatch",
"pm-dsm-aggregator-api",
"pm-dsm-curtailment",
"pm-dsm-notification-service",
"pm-file-api",
"rebalancer",
"reporting",
"resource",
"scada-ingester",
"servicebus-snowflake-bridge",
"tigrs",
"transfer-position-service",
"verra"
)
$selectedRepositories = $repositoriesResponse.value | Where-Object { $_.name -in $repositories}
# $body = "<h2> $email </h2>"
#             $body += "<h3>Branch Status Summary</h3>"
#             $body += $matrixHTML
# Loop through each repository

foreach ($repository in $selectedRepositories) 
{ 
    # $body += "<tr style='border:1px solid black;border-collapse: collapse'><b>Repo: </b>" + $repository.name + "<b>;   Default Branch: </b>" + $repository.defaultBranch + "</tr>`n"   
    
    #Main Branch is the default branch in the repository  
    if ($repository.defaultBranch -eq "refs/heads/main") {
        $branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.name)/stats/branches?baseVersionDescriptor.version=main&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
    }
    #Master Branch is the default branch in the repository 
    elseif ($repository.defaultBranch -eq "refs/heads/master") {
        $branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=master&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
    }
    #neither master nor main branches are default branch
    else {
        try{
            $branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?api-version=7.1-preview.1"
			$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
        }
        catch{
               $_.Exception.message
        }
    }

    # branch comparison details, commits ahead, and pending changes not approved
    $workItemInfo = New-Object -TypeName psobject
		$names = @("RepoName","BranchName", "AheadCount", "BehindCount")
		foreach($name in $names)
		{
				$workItemInfo | Add-Member -MemberType NoteProperty -Name $name -Value $null 
		}
			

    		#Write-Host "Repos: "$repository.name "; Default Branch: "$repository.defaultBranch
		foreach($commit in $commitsResponse.value)
		{
		# Creating a custom object with properties
			$workItemInfo.RepoName = $repository.name
			$workItemInfo.BranchName = $commit.name
			$workItemInfo.AheadCount = $commit.aheadCount
			$workItemInfo.BehindCount = $commit.behindCount
			# $workItemInfo.RepoID = $repository.id
			# $workItemInfo.URI = $branchComparisonApiUrl
			$workItemInfo
		}

    #Checking Tags details on Repos
    $tagsuri = "https://dev.azure.com/apxinc/Apx/_apis/git/repositories/$($repository.name)/refs?api-version=6.0-preview.1"
	$tagsresponse = Invoke-RestMethod -Uri $tagsuri -Method Get -Headers $headers
	$tags= $tagsresponse.value|Where-Object{$_.name -cmatch "refs/tags"}
    
    $latestVersion = [System.Version]::new("0.0.0")

    #Getting Latest version of the Tag
    foreach ($tag in $tags) 
    {
       $version = $tag.name
       
       #trimming refs/tags from the Version name 
       $versionNumber = $version.TrimStart('refs/tags/V')
	   try {
       		$parsedVersion = [System.Version]::new($versionNumber)      
            #Comparing all available versions in the Tag
            if ($parsedVersion -gt $latestVersion) 
            {
                $latestVersion = $parsedVersion
				$laterversion = [System.Version]::new("0.0.0")
	 			if($latestVersion -eq $laterversion){
					$workItemInfo.RepoName = $repository.name
					$workItemInfo.BranchName = "No Latest Tags"
					$workItemInfo.AheadCount = "NA"
					$workItemInfo.BehindCount = "NA"
					$workItemInfo
	 			}
				else {
					<# Action when all if and elseif conditions are false #>
	
					$vartag = "V"+ $latestVersion
	 				$uri1= "https://dev.azure.com/apxinc/Apx/_apis/git/repositories/$($repository.name)/diffs/commits?diffCommonCommit=false"
	 				$uri2 = "&baseVersion=$($vartag)&"
	 				$uri3 = "baseVersionOptions=none&"
	 				$uri4="baseVersionType=tag&"
	 				$uri5="targetVersion=main&"
	 				$uri6="targetVersionOptions=none&"
	 				$uri7="targetVersionType=branch&"
	 				$uri8="api-version=7.1-preview.1"
	 				$Versionuri = $uri1+$uri2+$uri3+$uri4+$uri5+$uri6+$uri7+$uri8
					$test= Invoke-RestMethod -Uri $Versionuri -Method Get -Headers $headers
					$workItemInfo.AheadCount = $test.aheadCount
					$workItemInfo.BehindCount = $test.behindCount
	            }            
        	}
        	        
     	}
		 catch{
			$_.Exception.Message
	   } 
	 
	 
             
           
        
        } 
		 # $body += "<tr><td align='right' style='border:1px solid black;border-collapse: collapse'>"+ $test.name + "<b> in Latest Version: V</b>" + $latestVersion +   "</td>
            #             <td align='right' style='border:1px solid black;border-collapse: collapse'>" + $test.behindCount + "</td>
            #             <td align='right' style='border:1px solid black;border-collapse: collapse'>" + $test.aheadCount + "</td></tr>"
			$workItemInfo.RepoName = $repository.name
			$workItemInfo.BranchName = "Latest Version V"+ $latestVersion
			
			# $workItemInfo.RepoID = $repository.id
			# $workItemInfo.URI = $Versionuri
			$workItemInfo
        
}

# $body += "</table></tbody>"	

# # compose and send out email messages to individuals who have booked to at least one task in the sprint
# Write-Host "-------------------------------------------------------------------------------------------"
# Write-Host "Sending out emails"


# Start-Sleep -s 15
# $SMTP_SERVER = "smtp.socketlabs.com"
# $emailSubject = "Branches Summary"
# $SMTP_PORT = 25
# $SMTP_USERNAME = "server4507"
# $PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force

# $bccEmailList = @("agurvits@apx.com","kliang@apx.com")

# $cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
# Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml   