$queryId = "4shokbf63trv4xr3wth635rgothfv57ii5n3odhh7eq2s5tccycq"
$projectReportingQueryUri = "https://dev.azure.com/apxinc/apx/_apis/git/repositories?api-version=6.0"
$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryId)"))
}

$repositoriesResponse = Invoke-RestMethod -Uri $projectReportingQueryUri -Headers $headers -Method Get

foreach ($repository in $repositoriesResponse.value) 
{
    
$branchesUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Name)/refs?filter=heads&api-version=7.1-preview.1"
$branchesResponse = Invoke-RestMethod -Uri $branchesUrl -Method Get -Headers $headers

# Initialize an array to store the differences
$differences = @()

# Loop through the branches and compare commits with the main branch
foreach ($branch in $branchesResponse.value) {
    $branchName = $branch.name -replace "refs/heads/", ","
    
    if ($branchName -ne $mainBranch) {
        # Get the latest commit SHA for the main branch
        $mainBranchCommitsUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Name)/commits?searchCriteria.itemVersion.version=$mainBranch&api-version=7.1-preview.1"
        $mainBranchCommitsResponse = Invoke-RestMethod -Uri $mainBranchCommitsUrl -Method Get -Headers $headers | Select-Object -ExpandProperty value

        if ($mainBranchCommitsResponse.Count -gt 0) {
            $mainBranchLatestCommit = $mainBranchCommitsResponse[0].committer.date
        } else {
            $mainBranchLatestCommit = "No commits on the main branch."
        }

        # Get the latest commit SHA for the current branch
        $branchCommitsUrl = "https://dev.azure.com/apxinc/apx/_apis/git/repositories/$($repository.Name)/commits?searchCriteria.itemVersion.version=$branchName&api-version=7.1-preview.1"
        $branchCommitsResponse = Invoke-RestMethod -Uri $branchCommitsUrl -Method Get -Headers $headers | Select-Object -ExpandProperty value

        if ($branchCommitsResponse.Count -gt 0) {
            $branchLatestCommit = $branchCommitsResponse[0].committer.date
        } else {
            $branchLatestCommit = "No commits on this branch."
        }

        if ($branchLatestCommit -ne $mainBranchLatestCommit) {
            $differences += "Branch: $branchName, Latest Commit Date: $branchLatestCommit, Latest Commit Date on Main Branch: $mainBranchLatestCommit"
        }
    }
}

# Output the differences
if ($differences.Count -gt 0) {
    Write-Host "Differences between the latest commits on branches and the main branch:"
    $differences | ForEach-Object { Write-Host $_ }
} else {
    Write-Host "No differences found between the latest commits on branches and the main branch."
}




}








