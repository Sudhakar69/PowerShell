$queryId = "4shokbf63trv4xr3wth635rgothfv57ii5n3odhh7eq2s5tccycq"
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
    $turi = "https://dev.azure.com/apxinc/Apx/_apis/git/repositories/$($repository.Id)/refs?api-version=6.0-preview.1"
	$tres = Invoke-RestMethod -Uri $turi -Method Get -Headers $headers
	$tags= $tres.value|Where-Object{$_.name -cmatch "refs/tags"}
	

foreach ($tag in $tags) {
	$version = $tag.name
    $latestVersion = [System.Version]::new("0.0.0")
    $versionNumber = $version.TrimStart('refs/tags/V')
    $parsedVersion = [System.Version]::new($versionNumber)
    try {
        
        if ($parsedVersion -gt $latestVersion) {
            $latestVersion = $parsedVersion
            $ltd = "V"+$latestVersion
            Write-Host $ltd " in " $repository.name
            # $turi = "https://dev.azure.com/apxinc/Apx/_apis/git/repositories/$($repository.Id)/stats/branches?baseVersionDescriptor.version=$ltd&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=tag&api-version=7.1-preview.1"
            # $tests= Invoke-RestMethod -Uri $suri -Method Get -Headers $headers
            # foreach($test in $tests.value){
            #     Write-Host "Ahead Count: " $test.aheadcount "; Behind Count: " $test.behindcount
            #     # $workItemInfo.BranchName = "V"+$latestVersion
            #     # $workItemInfo.AheadCount = $aheadcount
            #     # $workItemInfo.BehindCount = $behindcount
            #     # $workItemInfo
            # }
        }
    } catch {
        Write-Host "Skipping invalid version: $versionNumber in " $repository.name
    }
}
#Write-Host "Latest version is: V$($latestVersion) in  " $repository.name
}

