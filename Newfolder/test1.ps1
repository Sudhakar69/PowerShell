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
	# Make the API request to get a list of all branches in the repository
	write-host "we are in $($repository.name) repository" -BackgroundColor Yellow -ForegroundColor Black
    $branchesApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.id)/refs?filter=heads&api-version=6.0"
    $branchesResponse = Invoke-RestMethod -Uri $branchesApiUrl -Headers $headers -Method Get
    foreach ($branch in $branchesResponse.value) 
	{
		#$branchName = $branch.name
		$SourceBranch = $branch.name
		$targetBranch = "refs/heads/main"

		# Construct the API URL for listing branches in the repository
		$branchesApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.name)/refs?filter=heads&api-version=6.0"
        $headers = @{
			Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryId)"))
		}
		# Make the API request to get the list of branches in the repository
		$branchesResponse = Invoke-RestMethod -Uri $branchesApiUrl -Headers $headers -Method Get

		# Find the source and target branch objects
		$sourceBranchObject = $branchesResponse.value | Where-Object { $_.name -cmatch $sourceBranch }
		$targetBranchObject = $branchesResponse.value | Where-Object { $_.name -cmatch $targetBranch }

		if ($sourceBranchObject -ne $null -and $targetBranchObject -ne $null) {
			# Extract branch SHAs
			$sourceBranchSha = $sourceBranchObject.objectId
			$targetBranchSha = $targetBranchObject.objectId

			# Compare branch SHAs to determine if there are pending changes
			if ($sourceBranchSha -ne $targetBranchSha) {
				# Construct the API URL for comparing branches
				$branchComparisonApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.name)/commits?api-version=6.0&baseVersion=$targetBranchSha&targetVersion=$sourceBranchSha"

				# Make the API request to get the list of commits between branches
				$commitsResponse = Invoke-RestMethod -Uri $branchComparisonApiUrl -Headers $headers -Method Get
                # Construct the API URL for comparing branches
                $compareBranchesApiUrl = "https://dev.azure.com/apxinc/apx/_apis/policy/configurations?api-version=6.0"
				# Count the number of pending changes (commits)
				$pendingChangesCount = $commitsResponse.count
				# ...
                $branchComparisonResponse = Invoke-RestMethod -Uri $compareBranchesApiUrl -Headers $headers -Method Get
				$branchComparisonConfiguration = $null

                # Iterate through the branch comparison configurations to find the desired one
				foreach ($config in $branchComparisonResponse.value) {
    				if ($config.type -eq "VstsGitPullRequest" -and $config.settings.repositoryId -eq $repository.Name) {
       					$branchComparisonConfiguration += $config
        				break  # Exit the loop once a matching configuration is found
    				}
				}

				if ($branchComparisonConfiguration -ne $null) {
    

					$policyId = $branchComparisonConfiguration.id

				    # Construct the API URL for pending changes (files) for merging
				    $pendingChangesApiUrl = "https://dev.azure.com/apxinc/apx/_apis/policy/configurations/$policyId/settings?api-version=6.0"
		
				    # Make the API request to get pending changes (files) for merging
				    $pendingChangesResponse = Invoke-RestMethod -Uri $pendingChangesApiUrl -Headers $headers -Method Get
		
				    # Extract the file paths of pending changes
				    $pendingChangesPaths = $pendingChangesResponse.value.filePathFilters
		
				    # Construct the API URL for comparing branch commits
				    $compareCommitsApiUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Name)/compareCommits?api-version=6.0&$sourceBranchSha..$targetBranchSha"
		
				    # Make the API request to get commit comparison details
				    $commitComparisonResponse = Invoke-RestMethod -Uri $compareCommitsApiUrl -Headers $headers -Method Get
		
				    # Extract the count of commits ahead (changes in the target branch not in the source branch)
				    $commitsAheadCount = $commitComparisonResponse.aheadCount
		
				    # Calculate the count of pending changes not yet approved
				    $pendingChangesNotApprovedCount = ($pendingChangesPaths | Where-Object { $_ -notin $commitComparisonResponse.commitPaths }).Count
		
				    # Output branch comparison details, commits ahead, and pending changes not approved
				    Write-Host "Source Branch: $sourceBranch ($sourceBranchSha)"
				    Write-Host "Target Branch: $targetBranch ($targetBranchSha)"
				    Write-Host "Commits Ahead: $commitsAheadCount , $pendingChangesCount"
				    Write-Host "Pending Changes Not Approved: $pendingChangesNotApprovedCount" -BackgroundColor Red -ForegroundColor Black
			    }
			} 
			else {
				Write-Host "Branch Comparison Details:"
				Write-Host "Source Branch: $sourceBranch ($sourceBranchSha)"
				Write-Host "Target Branch: $targetBranch ($targetBranchSha)"
				Write-Host "There are no pending changes (commits) between the branches."
			}
		}
    }
	write-host "=========== "
}