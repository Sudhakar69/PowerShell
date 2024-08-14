param ([string]$info="Please enter password")
$email = @("")

<#

Project Reporting script for VSTS

This script sends emails to all developers, testers, and product managers notifying them
of pending branches for merging and send a gentle reminder to the Application owner and respective leads to say hey these are Ahead and Behind Count for these particular branches



#>
function New-RepoInfo {
    $RepoInfo = New-Object -TypeName PSObject
    $RepoInfo | Add-Member -MemberType NoteProperty -Name RepoName -Value $null
    $RepoInfo | Add-Member -MemberType NoteProperty -Name BranchesBehindMain -Value $null
    $RepoInfo | Add-Member -MemberType NoteProperty -Name AverageCommits -Value $null
    $RepoInfo | Add-Member -MemberType NoteProperty -Name MaxCommitsBehind -Value $null
    $RepoInfo | Add-Member -MemberType NoteProperty -Name MaxCommitBranchesBehind -Value $null
    $RepoInfo | Add-Member -MemberType NoteProperty -Name RowColor -Value $null

    return $RepoInfo
    
}
function New-BranchInfo {
    $BranchInfo = New-Object -TypeName PSObject
    $BranchInfo| Add-Member -MemberType NoteProperty -Name BranchName -Value $null
    $BranchInfo| Add-Member -MemberType NoteProperty -Name RepoBranchBehindMain -Value $null
    $BranchInfo| Add-Member -MemberType NoteProperty -Name AverageCommits -Value $null
    $BranchInfo| Add-Member -MemberType NoteProperty -Name MaxCommits -Value $null
    $BranchInfo| Add-Member -MemberType NoteProperty -Name MaxCommitRepos -Value $null
    return $BranchInfo
    
}
function New-BranchReport {
    $BranchReport = New-Object -TypeName psobject
    $BranchReport|Add-Member -MemberType NoteProperty -Name BranchName -Value $null
    $BranchReport|Add-Member -MemberType NoteProperty -Name BranchBehindCount -Value $null
    $BranchReport|Add-Member -MemberType NoteProperty -Name BranchAheadCount -Value $null
    return $BranchReport
    
}
function New-TagReport {
    $TagReport = New-Object -TypeName psobject
    $TagReport | Add-Member -MemberType NoteProperty -Name TagVersion -Value $null
    $TagReport | Add-Member -MemberType NoteProperty -Name TagBehindCount -Value $null
    $TagReport | Add-Member -MemberType NoteProperty -Name TagAheadCount -Value $null
    return $TagReport
    
}

$queryId = "4shokbf63trv4xr3wth635rgothfv57ii5n3odhh7eq2s5tccycq"
$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryId)"))
}
$projectReportingQueryUri = "https://dev.azure.com/$($organization)/$($project)/_apis/git/repositories?api-version=6.0"
$repositoriesResponse = Invoke-RestMethod -Uri $projectReportingQueryUri -Headers $headers -Method Get
$repositories= @("")
$selectedRepositories = $repositoriesResponse.value | Where-Object { $_.name -in $repositories}

$BranchReportHtml = "<h1>Branch Status Summary</h1>"
    
$BranchReportHtml += "<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
    <tr style='font-size:12px;'>
        <th valign='top' nowrap><b>Branch Name</b></th>
        <th valign='top' bgcolor='#ff0000' nowrap><b>Behind Count(vs 'Main')</b></th>
        <th valign='top' bgcolor='#90ee90'><b>Ahead Count(vs 'Main')</b></th>
    </tr>"
$RepoSummaryHtml = "<h1>Summary by Repo</h1>"

$RepoSummaryHtml += "<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
    <tr style='font-size:12px;'>
        <th valign='top' nowrap><b>Repo Name</b></th>
        <th valign='top' nowrap><b>Of Branches <br>Behind 'Main'/'Master'</b></th>
        <th valign='top' ><b>Average of Commits <br>Behind (rounded up)</b></th>
        <th valign='top' ><b>Max of Commits Behind</b></th>
        <th valign='top' ><b>Branch(es) with Max of Commits Behind</b></th>
    </tr>
"
# ==============================================
$BranchSummaryHtml = "<h1>Summary by Branch</h1>"
$BranchSummaryHtml += "<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
    <tr style='font-size:12px;'>
        <th valign='top' nowrap><b>Branch Name</b></th>
        <th valign='top' nowrap><b>Of Repos in which Branch <br>is Behind 'Main'/'Master'</b></th>
        <th valign='top' ><b>Average of Commits<br> Behind (rounded up)</b></th>
        <th valign='top' ><b>Max of Commits<br> Behind</b></th>
        <th valign='top' ><b>Repo(s) with Max of Commits Behind</b></th>   
    </tr>
"
$ColorcodingHtml = "<h2>Color Coding</h2>"
$ColorcodingHtml += "<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
    <tr style='font-size:12px;'>
        <th valign='top' nowrap><b>Color</b></th>
        <th valign='top' nowrap><b>Of Branch is Behind<br> 'Main'/'Master'</b></th>
        <th valign='top' ><b>Or/And</b></th>
        <th valign='top' ><b>Average of <br>Commits Behind is</b></th>    
    </tr>"

# =========================================
# $body = "<h2> $email </h2>"
#             $body += "<h3>Branch Status Summary</h3>"
#             $body += $matrixHTML
# Loop through each repository
$RepoSummaryData = @()
$listofBranches = @()
$sortedRepositories = $selectedRepositories | Sort-Object -Property name
foreach ($repository in $sortedRepositories) { 
    $BranchReportHtml += "<tr>
            <td valign='top' bgcolor='#D3D3D3' style='font-size:13.5px;' colspan=3 nowrap><b style='font-size:12px;'>Repo:&nbsp; </b><a href='https://$($organization).visualstudio.com/$($Project)/_git/$($repository.name)/branches'>$($repository.name)</a></td>
        </tr>"   
    $tagsuri = "https://dev.azure.com/$($organization)/$($project)/_apis/git/repositories/$($repository.Id)/refs?api-version=6.0-preview.1"
    $tagsresponse = Invoke-RestMethod -Uri $tagsuri -Method Get -Headers $headers
    #Main Branch is the default branch in the repository  
    if ($repository.defaultBranch -eq "refs/heads/main") {
        $branchComparisonApiUrl = "https://dev.azure.com/$($organization)/$($project)/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=main&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
    }
    #Master Branch is the default branch in the repository 
    elseif ($repository.defaultBranch -eq "refs/heads/master") {
        $branchComparisonApiUrl = "https://dev.azure.com/$($organization)/$($project)/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=master&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
    }
    #neither master nor main branches are default branch
    else {
        $branchComparisonApiUrl = "https://dev.azure.com/$($organization)/$($project)/_apis/git/repositories/$($repository.Id)/stats/branches?api-version=7.1-preview.1"
		$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
      
    }
    $commitbehindCount = 0
    $maxcommits = 0
    $maxcommitBranches = @()
    $commitBranches =@()
    foreach($commit in $commitsResponse.value){
        foreach($Branch in $tagsresponse.value){
            #Getting only unlocked branch details
            if ($null -eq $Branch.PSObject.Properties['IsLocked'] -or $Branch.IsLocked -eq $false) 
            {   
                #Getting only Branches from commits and Branches objects         
                if ($Branch.name -cmatch $commit.name -and $Branch.objectId -eq $commit.commit.commitId ) 
                {
                    $commitbehindCount += $commit.behindCount
                    $behindcountofcommits = $commit.behindCount
                    $commitName = $commit.name
                    #Color coding for Behind commits of Branch Report
                    if ($behindcountofcommits -gt 0) {
                        $BehindCountRowColor = '#ff0000'
                    }
                    else {
                        $BehindCountRowColor = '#ffffff'
                    }
                    # $BranchUri = $BranchbasicUri + $commitName
                    if($behindcountofcommits -gt $maxcommits)
                    {
                        $maxcommits = $behindcountofcommits
                        $maxcommitBranches +=  $commitName -join ","
                    }
                    # else
                    # {        
                    #     $commit.name =  "N/A"
                    # } 
                    if ($commit.name -ne "main" -and $commit.aheadCount -eq 0) {
                        $commitcolor = '#FFA500'                
                    }
                    else {
                        $commitcolor = '#ffffff' 
                    }
                    $commitBranches += $commitName
                    $BranchReportHtml += " 
                        <tr style='font-size:12px;' nowrap>
                            <td nowrap><a href='https://$($organization).visualstudio.com/$($Project)/_git/$($repository.name)?version=GB$($commit.name)'>$($commit.name)</a></td>
                            <td bgcolor='$($BehindCountRowColor)'>$($commit.behindCount)</td>
                            <td bgcolor='$($commitcolor)'>$($commit.aheadCount)</td>
                        </tr>"                                  
                }
                # $maxcommitBranches = $maxcommitBranches.Trim()
                $RepoInfo = New-RepoInfo
                $RepoInfo.RepoName = $repository.name
                $RepoInfo.BranchesBehindMain = $commitsResponse.value.Count
                $RepoInfo.MaxCommitsBehind = $maxcommits
                $RepoInfo.MaxCommitBranchesBehind = $maxcommitBranches
                $AverageCommits = [math]::Ceiling($commitbehindCount / $RepoInfo.BranchesBehindMain)
                $RepoInfo.AverageCommits = $AverageCommits
                #Row Coloring using commitbehind count and average count
                if ($commitbehindCount -ge 5 -or $AverageCommits -ge 10) 
                {
                    $Rowcolor='#ff0000'
                }
                else
                {
                    if ($commitbehindCount -ge 3 -and $commitbehindCount -lt 5 -or $AverageCommits -ge 7) 
                    {
                        $Rowcolor='#FFA500'
                    }
                    elseif ($commitbehindCount -ge 1 -and $commitbehindCount -lt 3 -or $AverageCommits -ge 1) 
                    {
                        $Rowcolor='#FFFF00' 
                    }
                    else
                    {
                        if ($commitbehindCount -eq 0 -and $AverageCommits -eq 0) 
                        {
                            $Rowcolor='#90EE90' 
                        }         
                    }

                } 
            }
            else {
                Write-Host $Branch.name "is Locked"
            }
            $RepoInfo.RowColor = $Rowcolor   
        }
    }
    $RepoSummaryData += $RepoInfo
    $listofBranches += $commitBranches
   
    # branch comparison details, commits ahead, and pending changes not approved
    #Checking Tags details on Repos
    $tags= $tagsresponse.value|Where-Object{$_.name -cmatch "refs/tags"}    
    $latestVersion = [System.Version]::new("0.0.0")
    $problematicVersions = @()
    #Getting Latest version of the Tag
    foreach ($tag in $tags) 
    {
       $version = $tag.name
       
       #trimming refs/tags from the Version name 
       $versionNumber = $version -replace "refs/tags/V",""
       
        #Try catch block to overcome issues like version names are not correct formate or doesn't found versions, etc
       try {
            $parsedVersion = [System.Version]::new($versionNumber)
            #Comparing all available versions in the Tag
            if ($parsedVersion -gt $latestVersion) 
            {
                $latestVersion = $parsedVersion
            }           
        }
        catch{
            $problematicVersions += $version
        }
           
     }
     if ($problematicVersions.Count -gt 0) {
        Write-Host "Problematic version strings:"
        $problematicVersions | ForEach-Object { Write-Host $_  " in "$repository.name}
    }
     $laterVersion = [System.Version]::new("0.0.0")
        if ($latestVersion -eq $laterVersion) {
            $test.behindCount = 0
            $test.aheadCount = 0
        }
        else {
            #To Overcome issues like main or master branches are not default branches or null value exceptions and etc
            try {
                if ($repository.defaultBranch -eq "refs/heads/main") {
                    $Versionuri = "https://dev.azure.com/$($organization)/$($project)/_apis/git/repositories/$($repository.Id)/diffs/commits?diffCommonCommit=false&baseVersion=V$($latestVersion)&baseVersionOptions=none&baseVersionType=tag&targetVersion=main&targetVersionOptions=none&targetVersionType=branch&api-version=7.1-preview.1"
                    $test= Invoke-RestMethod -Uri $Versionuri -Method Get -Headers $headers    
                }
                else {
                    $Versionuri = "https://dev.azure.com/$($organization)/$($project)/_apis/git/repositories/$($repository.Id)/diffs/commits?diffCommonCommit=false&baseVersion=V$($latestVersion)&baseVersionOptions=none&baseVersionType=tag&targetVersion=master&targetVersionOptions=none&targetVersionType=branch&api-version=7.1-preview.1"
                    $test= Invoke-RestMethod -Uri $Versionuri -Method Get -Headers $headers    
                }
            }
            catch {
                $_.Exception.Message
            }
        }
        $TagReport = New-TagReport
        $TagReport.TagVersion = $latestVersion
        $TagReport.TagBehindCount = $test.aheadCount
        $TagReport.TagAheadCount = $test.behindCount
        #Row Coloring using commitbehind count and average count
        if ($test.behindCount -gt 0) 
        {
            $Tagbgcolor='#ff0000'           
        }
        else
        {
            $Tagbgcolor ='#90EE90'

        }      
        $BranchReportHtml += " 
                    <tr style='font-size:12px;'nowrap>
                        <td nowrap><b>Latest Production Tag:&nbsp;</b> V$($TagReport.TagVersion)</td>
                        <td>$($TagReport.TagBehindCount)</td>
                        <td bgcolor='$($Tagbgcolor)' >$($TagReport.TagAheadCount)</td>
                    </tr>"  
}
$sortedRepoSummaryData = $RepoSummaryData  |Sort-Object {
    switch ($_.RowColor) {
    '#ff0000' { 1 }  # Red
    '#FFA500' { 2 }  # Orange
    '#FFFF00' { 3 }  # Yellow
    '#90EE90' { 4 }  # Light Green
    default { 5 }    # Default for other colors
} 
},RepoName
$colors= ("Red","Orange","Yellow","Green")
foreach($color in $colors){
    if ($color -eq "Red") {
        $CellColor= '#ff0000'
        $BehindCode = ">=5"
        $ColorCondition= "OR"
        $AheadCode = ">=10"
    }
    else {
        if ($color -eq "Orange") {
            $CellColor= '#FFA500'
            $BehindCode = ">=3 and <5"
            $ColorCondition= "OR"
            $AheadCode = ">=7"
        }
        elseif ($color -eq "Yellow") {
            $CellColor= '#FFFF00'
            $BehindCode = ">=1 and <3"
            $ColorCondition= "OR"
            $AheadCode = ">=7"
        }
        else {
            $CellColor= '#90EE90'
            $BehindCode = "=0"
            $ColorCondition= "AND"
            $AheadCode = "=0"
        }
    }
    $ColorcodingHtml += "
        <tr style='font-size:12px;' nowrap>
            <td  bgcolor='$($CellColor)'nowrap><b>$($color)</b></td>
            <td>$($BehindCode)</td>
            <td border='0'>$($ColorCondition)</td>
            <td >$($AheadCode)</td>
        </tr>
"
}

foreach ($Repodata in $sortedRepoSummaryData) {
    $RepoBasicUri = "https://$($organization).visualstudio.com/$($Project)/_git/$($Repodata.RepoName)?version=GB"
    $RepoSummaryHtml += "<tr style='font-size:12px;' bgcolor='$($Repodata.RowColor)' nowrap>
            <td nowrap><a href='https://$($organization).visualstudio.com/$($Project)/_git/$($Repodata.RepoName)/branches'>$($Repodata.RepoName)</a></td>
            <td nowrap>$($Repodata.BranchesBehindMain)</td>
            <td nowrap>$($Repodata.AverageCommits)</td>
            <td snowrap>$($Repodata.MaxCommitsBehind)</td>
            <td nowrap>"

                # Create an array of Repo names with links
                $RepoUris = $Repodata.MaxCommitBranchesBehind | ForEach-Object {
                    $RepoUri = $RepobasicUri + $_
                        "<a href='$RepoUri'>$_</a><br>"
                }

                # Join the branch names with commas
                # $RepoNamesString = $RepoUris -join ', '

                $RepoSummaryHtml += $RepoUris
                $RepoSummaryHtml += "
                        </td>
                    </tr>
"
}
#filterin with unique branches and sorting in alphabetical order
$uniqueBranchNames = $listofBranches|Sort-Object |Select-Object -Unique
foreach($uniqueBranchName in $uniqueBranchNames){
    $RepoBranchCommitsBehind =0
    $RepoBranchMaxCommits =0
    $Repos = @()
    $Reposmaxcommit = @()
    foreach($repository in $sortedRepositories){
        $RepobranchComparisonApiUrl = "https://dev.azure.com/$($organization)/$($project)/_apis/git/repositories/$($repository.name)/stats/branches?api-version=7.1-preview.1"
		$RepoBranchcommitsResponse = Invoke-RestMethod -Uri $RepobranchComparisonApiUrl -Headers $headers -Method Get
        foreach($RepoBranch in $RepoBranchcommitsResponse.value){
            $RepoBranchName = $RepoBranch.name
            #Checking unique branches which are present in Branches in repository
            if ($RepoBranchName -eq $uniqueBranchName) {
                $RepoBranchCommitsBehind += $RepoBranch.behindCount
                $RepoBranchCommits =$RepoBranch.behindCount
                #Getting Highest commits behind Repositories by Branches wise
                if ($RepoBranchCommits -gt $RepoBranchMaxCommits) {
                    $RepoBranchMaxCommits = $RepoBranchCommitsBehind
                    $Reposmaxcommit += $repository.name
                }
                $Repos = $repository.name 
            }    
        }
    }
    # $RepoBranchMaxCommits=$RepoBranchMaxCommits.TrimEnd(',')
    $BranchInfo = New-BranchInfo
    $BranchInfo.BranchName = $uniqueBranchName
    $BranchInfo.RepoBranchBehindMain = $Repos.Count    
    $BranchInfo.MaxCommits = $RepoBranchMaxCommits
    $BranchInfo.MaxCommitRepos =$Reposmaxcommit
    
    $RepoBranchAverageCommits = [math]::Ceiling($RepoBranchCommitsBehind / $($BranchInfo.RepoBranchBehindMain))
    $BranchInfo.AverageCommits = $RepoBranchAverageCommits
    $RepoBaseUri = "https://$($organization).visualstudio.com/$($Project)/_git/"
    $BranchSummaryHtml += "
        <tr style='font-size:12px;' nowrap>
            <td nowrap><a href='https://$($organization).visualstudio.com/$($Project)/_git/$($BranchInfo.MaxCommitRepos[-1])?version=GB$($BranchInfo.BranchName)' >$($BranchInfo.BranchName)</a></td>
            <td nowrap>$($BranchInfo.RepoBranchBehindMain)</td>
            <td nowrap>$($BranchInfo.AverageCommits)</td>
            <td nowrap>$($BranchInfo.MaxCommits)</td>
            <td nowrap>
"
                # Create an array of Repo names with links
                $BranchUris = $Reposmaxcommit | ForEach-Object {
                    $BranchUri = $RepoBaseUri + $_ + "/branches?_a=all"
                        "<a href='$BranchUri'>$_</a><br>"
                }
                $BranchSummaryHtml += $BranchUris
                $BranchSummaryHtml += "
                        </td>
                    </tr>
"   
}              
# Close the HTML table and body
$BranchSummaryHtml +="
</table></tbody>
"
$ColorcodingHtml +="
</table></tbody>
"
$RepoSummaryHtml += "
    </table></tbody>
"
$BranchReportHtml += "
    </table></tbody>
"
# $body += $BranchSummaryHtml
$body += $RepoSummaryHtml
$body += $ColorcodingHtml
$body += $BranchSummaryHtml
$body  += $BranchReportHtml


# # compose and send out email messages to individuals who have booked to at least one task in the sprint
Write-Host "-------------------------------------------------------------------------------------------"
Write-Host "Sending out emails"
$emailSubject = "Branch Summary status report"
Start-Sleep -s 15
		
$SMTP_SERVER = "smtp.socketlabs.com"
$SMTP_PORT = 25
$SMTP_USERNAME = "server4507"
$PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force
$bccEmailList = ("agurvits@$($Project).com","kliang@$($Project).com","c-tgudise@xpansiv.com")
$cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From tfsbuild@$($Project).com -Subject $emailSubject -Body $body -BodyAsHtml
