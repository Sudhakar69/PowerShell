param ([string]$info="PLEASE PROVIDE PASSWORD AS AN ARGUMENT TO THIS SCRIPT")

<#

Project Reporting script for VSTS

This script sends emails to all developers, testers, and product managers notifying them
of pending branches for merging and send a gentle reminder to the Application owner and respective leads to say hey these are Ahead and Behind Count for these particular branches



#>
Add-Type -AssemblyName System.Web
$scriptDirectory = ($PSScriptRoot | Split-Path -Parent)
$commonLibPath = (Join-Path (Join-Path $scriptDirectory -ChildPath "common") -ChildPath "common.ps1")
. $commonLibPath
$infoLibPath = (Join-Path $PSScriptRoot -ChildPath "project_reporting_common.ps1")
. $infoLibPath



$matrixHTML = "<tbody><table border='1' cellspacing='1' style='font-size:9px;border:1px solid black;border-collapse: collapse'><tr>`n"
$matrixHTML += "<td valign='top'><b>Branch Name</b></td>`n"
$matrixHTML += "<td valign='top' bgcolor='#ff0000' nowrap><b>Behind Count</b></td>`n"
$matrixHTML += "<td valign='top' bgcolor='#90ee90'><b>Ahead Count</b></td>`n"
$matrixHTML += "</tr>`n"

$BranchSummary = Get-BranchSummary

$queryId = "Query ID"
$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryId)"))
}
$projectReportingQueryUri = "https://dev.azure.com/$($Organization)/$($project)/_apis/git/repositories?api-version=6.0"
$repositoriesResponse = Invoke-RestMethod -Uri $projectReportingQueryUri -Headers $headers -Method Get

$body = "<h2> $email </h2>"
            $body += "<h3>Branch Status Summary</h3>"
            $body += $matrixHTML
# Loop through each repository
foreach ($repository in $repositoriesResponse.value) 
{ 
    $body += "<tr style='border:1px solid black;border-collapse: collapse'><b>Repo: </b>" + $repository.name + "<b>;   Default Branch: </b>" + $repository.defaultBranch + "</tr>`n"   
    
    #Main Branch is the default branch in the repository  
    if ($repository.defaultBranch -eq "refs/heads/main") {
        $branchComparisonApiUrl = "https://dev.azure.com/$($Organization)/$($project)/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=main&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
    }
    #Master Branch is the default branch in the repository 
    elseif ($repository.defaultBranch -eq "refs/heads/master") {
        $branchComparisonApiUrl = "https://dev.azure.com/$($Organization)/$($project)/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=master&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
    }
    #neither master nor main branches are default branch
    else {
        try{
            $branchComparisonApiUrl = "https://dev.azure.com/$($Organization)/$($project)/_apis/git/repositories/$($repository.Id)/stats/branches?api-version=7.1-preview.1"
			$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
        }
        catch{
               $_.Exception.message
        }
    }

    # branch comparison details, commits ahead, and pending changes not approved
    foreach($commit in $commitsResponse.value){
           
        $body += "<tr><td align='right' style='border:1px solid black;border-collapse: collapse'>" + $commit.name + "</td>
                        <td align='right' style='border:1px solid black;border-collapse: collapse'>" + $commit.behindCount + "</td>
                        <td align='right' style='border:1px solid black;border-collapse: collapse'>" + $commit.aheadCount + "</td></tr>"
            
            
    }

    #Checking Tags details on Repos
    $tagsuri = "https://dev.azure.com/$($Organization)/$($project)/_apis/git/repositories/$($repository.Id)/refs?api-version=6.0-preview.1"
	$tagsresponse = Invoke-RestMethod -Uri $tagsuri -Method Get -Headers $headers
	$tags= $tagsresponse.value|Where-Object{$_.name -cmatch "refs/tags"}
    
    $latestVersion = [System.Version]::new("0.0.0")

    #Getting Latest version of the Tag
    foreach ($tag in $tags) 
    {
       $version = $tag.name
       
       #trimming refs/tags from the Version name 
       $versionNumber = $version.TrimStart('refs/tags/V')
       $parsedVersion = [System.Version]::new($versionNumber)

       try {
            #Comparing all available versions in the Tag
            if ($parsedVersion -gt $latestVersion) 
            {
                $latestVersion = $parsedVersion
            }
            $Versionuri = "https://dev.azure.com/$($Organization)/$($project)/_apis/git/repositories/$($repository.id)/stats/branches?baseVersionDescriptor.version=V$($latestVersion)&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=tag&api-version=7.1-preview.1"
            $tests= Invoke-RestMethod -Uri $Versionuri -Method Get -Headers $headers
                #$aheadcount =0
                #$behindcount = 0
             foreach($test in $tests.value){
                    $body += "<tr><td align='right' style='border:1px solid black;border-collapse: collapse'>"+ $test.name + "<b> in Latest Version: V</b>" + $latestVersion +   "</td>
                        <td align='right' style='border:1px solid black;border-collapse: collapse'>" + $test.behindCount + "</td>
                        <td align='right' style='border:1px solid black;border-collapse: collapse'>" + $test.aheadCount + "</td></tr>"
             }
                
            
        }
        catch{
             $_.Exception.Message
        }
        
            
     }
        
        
        
}

$body += "</table></tbody>"	

# compose and send out email messages to individuals who have booked to at least one task in the sprint
Write-Host "-------------------------------------------------------------------------------------------"
Write-Host "Sending out emails"


Start-Sleep -s 15
$SMTP_SERVER = "smtp.socketlabs.com"
$emailSubject = "Branches Summary"
$SMTP_PORT = 25
$SMTP_USERNAME = "server4507"
$PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force

$bccEmailList = @("")

$cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From tfsbuild@apx.com -Subject $emailSubject -Body $body -BodyAsHtml -Cc $