Add-Type -AssemblyName System.Web

# Get the script directory
$path= "C:\Users\c-tgudise\common\common.ps1"
# # Load and execute the common library script
$commonLibContent = Get-Content -Path $Path -Raw
Invoke-Expression -Command $commonLibContent

$PATId = "Personal Access Token ID"
$queryId = "343d58e6-67ad-48a1-9317-1403d32e1dab"

$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($PATId)"))
}
$latestSprint =Get-SprintInfo -forDate (Get-Date)
$latestSprintNumber = $latestSprint.sprintNumber
$presentsprint = "$($Project)\Sprint " + $latestSprintNumber
$WorkitemsQueryUri = "https://$($Organization).visualstudio.com/DefaultCollection/_apis/wit/wiql/$($queryId)?api-version=4.1"
$WorkitemsQueryResult = Invoke-RestMethod -Uri $WorkitemsQueryUri -Method Get -Headers $headers 
$workItemIds = $WorkitemsQueryResult.workItemRelations
foreach($workItemId in $workItemIds.target.id)
{
    $workuri = "https://dev.azure.com/$($Organization)/$($Project)/_apis/wit/workItems/101812/revisions?api-version=7.2-preview.3"
    $workresponse= Invoke-RestMethod -Uri $workuri -Method Get -Headers $headers
}