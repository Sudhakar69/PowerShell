
Param(
    [Parameter(Mandatory=$true)]
    [string] $BranchName
)
$queryId = "Query ID"
$headers = @{
Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryId)"))
}
$projectReportingQueryUri = "https://dev.azure.com/$($Organization)/$($Project)/_apis/git/repositories?api-version=6.0"
$repositoriesResponse = Invoke-RestMethod -Uri $projectReportingQueryUri -Headers $headers -Method Get

# $selectedRepositories = $repositoriesResponse.value |Select-Object -Property name
$sortedRepositories = $repositoriesResponse.value | Sort-Object -Property name
$Repos = @()
foreach($repository in $sortedRepositories){
    try {
        $RepobranchComparisonApiUrl = "https://dev.azure.com/$($Organization)/$($Project)/_apis/git/repositories/$($repository.name)/stats/branches?api-version=7.1-preview.1"
	    $RepoBranchcommitsResponse = Invoke-RestMethod -Uri $RepobranchComparisonApiUrl -Headers $headers -Method Get
        foreach($RepoBranch in $RepoBranchcommitsResponse.value){
            $RepoBranchName = $RepoBranch.name
            #Checking unique branches which are present in Branches in repository
            if ($RepoBranchName -eq $BranchName) {   
                $Repos += $repository.name
            }   
  
        }
    }
    catch {
        Write-Host $repository.name "Does not exist"
    }

}
Write-Output $Repos

    
    