$queryId = "4shokbf63trv4xr3wth635rgothfv57ii5n3odhh7eq2s5tccycq"


$projectReportingQueryUri = "https://dev.azure.com/apxinc/apx/_apis/git/repositories?api-version=6.0"
	$headers = @{		Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryId)"))}

$repositoriesResponse = Invoke-RestMethod -Uri $projectReportingQueryUri -Headers $headers -Method Get



foreach ($repository in $repositoriesResponse.value) 
	{
		#Write-Host "Repo: " $repository.name "; Default Branch: " $repository.defaultBranch 
		if ($repository.defaultBranch -eq "refs/heads/main") 
		{
			$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=main&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
			$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get	
			Write-Host $repository.name " has MAIN as default branch" -BackgroundColor Yellow -ForegroundColor Black
		}
		elseif ($repository.defaultBranch -eq "refs/heads/master") 
		{
			$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=master&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
			$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get	
			Write-Host $repository.name " has Master as Default branch" -BackgroundColor Green -ForegroundColor Black
		
		}
		else 
		{
			try
			{
				$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Id)/stats/branches?api-version=7.1-preview.1"
				$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get	
				Write-Host $repository.defaultBranch " is the default branch in " $repository.name -BackgroundColor Red -ForegroundColor Black
				return $commitsResponse
			}
			catch
			{
				$commit= $_.Exception.message
				return $commitsResponse
			}
		}

        $workItemInfo = New-Object -TypeName psobject
		$names = @("BranchName", "AheadCount", "BehindCount")
		foreach($name in $names)
		{
				$workItemInfo | Add-Member -MemberType NoteProperty -Name $name -Value $null 
		}
			

    		#Write-Host "Repos: "$repository.name "; Default Branch: "$repository.defaultBranch
		foreach($commit in $commitsResponse.value)
		{
		# Creating a custom object with properties
			$workItemInfo.BranchName = $commit.name
			$workItemInfo.AheadCount = $commit.aheadCount
			$workItemInfo.BehindCount = $commit.behindCount
			$workItemInfo
		}
	}


