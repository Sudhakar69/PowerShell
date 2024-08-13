# $xl = New-Object -ComObject Excel.Application

# # Make Excel visible (optional)
# $xl.Visible = $true

# # Add a new workbook
# $workbook = $xl.Workbooks.Add()

# # Add a new worksheet to the workbook
# $worksheet = $workbook.Worksheets.Add()
# $worksheet.Cells.Item(1,1) = "Branch Name"
# $worksheet.Cells.Item(1,2) = "Ahead Count"
# $worksheet.Cells.Item(1,3) = "Behind Count"
function New-RepoInfo {
    $RepoInfo = New-Object -TypeName PSObject
    $RepoInfo | Add-Member -MemberType NoteProperty -Name RepoName -Value $null
    $RepoInfo | Add-Member -MemberType NoteProperty -Name BehindMain -Value $null
    $RepoInfo | Add-Member -MemberType NoteProperty -Name AheadMain -Value $null
    $RepoInfo | Add-Member -MemberType NoteProperty -Name MaxCommits -Value $null
    $RepoInfo | Add-Member -MemberType NoteProperty -Name MaxCommitBranches -Value $null
    
}
function New-BranchInfo {
    $BranchInfo = New-Object -TypeName PSObject
    $BranchInfo| Add-Member -MemberType NoteProperty -Name BranchName -Value $null
    $BranchInfo| Add-Member -MemberType NoteProperty -Name BehindMain -Value $null
    $BranchInfo| Add-Member -MemberType NoteProperty -Name AheadMain -Value $null
    $BranchInfo| Add-Member -MemberType NoteProperty -Name MaxCommits -Value $null
    $BranchInfo| Add-Member -MemberType NoteProperty -Name MaxCommitBranches -Value $null
    $BranchInfo| Add-Member -MemberType NoteProperty -Name IsLocked -Value $null
    $BranchInfo| Add-Member -MemberType NoteProperty -Name IsTag -Value $null
    $BranchInfo| Add-Member -MemberType NoteProperty -Name CreatedDate -Value $null
    
}
$queryId = "4shokbf63trv4xr3wth635rgothfv57ii5n3odhh7eq2s5tccycq"
$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryId)"))
}
$projectReportingQueryUri = "https://dev.azure.com/apxinc/apx/_apis/git/repositories?api-version=6.0"
$repositoriesResponse = Invoke-RestMethod -Uri $projectReportingQueryUri -Headers $headers -Method Get# For ($i=2; $i -le 10000; $i++){
# Loop through each repository
foreach ($repository in $repositoriesResponse.value) 
{
    $workItemInfo = New-Object -TypeName psobject
	$names = @("Repo","BranchName", "AheadCount", "BehindCount")
    foreach($name in $names)
	{
		$workItemInfo | Add-Member -MemberType NoteProperty -Name $name -Value $null 
	}
        # $worksheet.Cells.Item($i,1)= 
        Write-Host "Repo: " + $repository.name + ";   Default Branch: " + $repository.defaultBranch
        if ($repository.defaultBranch -eq "refs/heads/main") {
            $branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=main&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
			$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
        }
        elseif ($repository.defaultBranch -eq "refs/heads/master") {
            $branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=master&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
			$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
        }
        else {
            try{
                $branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?api-version=7.1-preview.1"
				$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
            }
            catch{
                Write-Host $_.Exception.message
            }
        }
        foreach($commit in $commitsResponse.value){
           
            $workItemInfo.Repo = $repository.name
            $workItemInfo.BranchName = $commit.name
	        $workItemInfo.AheadCount = $commit.aheadCount
	        $workItemInfo.BehindCount = $commit.behindCount
            $workItemInfo
        }
        $turi = "https://dev.azure.com/apxinc/Apx/_apis/git/repositories/$($repository.Id)/refs?api-version=6.0-preview.1"
	    $tres = Invoke-RestMethod -Uri $turi -Method Get -Headers $headers
	    $tags= $tres.value|Where-Object{$_.name -cmatch "refs/tags"}
        $latestVersion = [System.Version]::new("0.0.0")
        foreach ($tag in $tags) 
        {
            $version = $tag.name
        
            $versionNumber = $version.TrimStart('refs/tags/V')
            $parsedVersion = [System.Version]::new($versionNumber)
            try {
                if ($parsedVersion -gt $latestVersion) 
                {
                    $latestVersion = $parsedVersion
                }
                $suri = "https://dev.azure.com/apxinc/Apx/_apis/git/repositories/$($repository.id)/stats/branches?baseVersionDescriptor.version=V$latestVersion&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=tag&api-version=7.1-preview.1"
                $tests= Invoke-RestMethod -Uri $suri -Method Get -Headers $headers
                $aheadcount =0
                $behindcount = 0
                foreach($test in $tests.value){
                    $aheadcount += $test.aheadcount
                    $behindcount += $test.behindcount
                }
                
                $workItemInfo.Repo = $repository.name
                $workItemInfo.BranchName = "Latest Version: V" + $latestVersion
	            $workItemInfo.AheadCount = $aheadcount
	            $workItemInfo.BehindCount = $behindcount
            
            }
            catch{
                Write-Host $_.Exception.Message
            }
        }
        
        
     }
# }

   