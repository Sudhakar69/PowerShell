<#
Helper function to create blank SprintInfo object
#>
$jiraPATId=""



$pair = "$(""):$($jiraPATId)"
$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($pair)"))
}
function New-SprintDetails
{
    $SprintDetails = New-Object -TypeName PSObject
    $SprintDetails | Add-Member -MemberType NoteProperty -Name ID -Value $null
    $SprintDetails | Add-Member -MemberType NoteProperty -Name Name -Value $null
    $SprintDetails | Add-Member -MemberType NoteProperty -Name sprintNumber -Value $null
    $SprintDetails | Add-Member -MemberType NoteProperty -Name startDate -Value $null
    $SprintDetails | Add-Member -MemberType NoteProperty -Name endDate -Value $null
    $SprintDetails | Add-Member -MemberType NoteProperty -Name status -Value $null
    
    return $SprintDetails
}
# function New-SprintInfo
# {
#     $sprintInfo = New-Object -TypeName PSObject
#     $sprintInfo | Add-Member -MemberType NoteProperty -Name ID -Value $null
#     $sprintInfo | Add-Member -MemberType NoteProperty -Name Name -Value $null
#     $sprintInfo | Add-Member -MemberType NoteProperty -Name startDate -Value $null
#     $sprintInfo | Add-Member -MemberType NoteProperty -Name endDate -Value $null
#     $sprintInfo | Add-Member -MemberType NoteProperty -Name sprintNumber -Value $null
#     $sprintInfo | Add-Member -MemberType NoteProperty -Name status -Value $null
    
#     return $sprintInfo
# }
<#
Returns SprintInfo object for sprint in wihch the date specified as the input parameter resides.
#>
$workupdates = @()
for ($i = 500; $i -lt 1000; $i++) {
    try {
        $sprinturi ="https://$($organization).atlassian.net/rest/agile/1.0/sprint/$($i)"
        $sprintresult = Invoke-RestMethod -Method Get -Uri $sprinturi -Headers $headers
        if ($null -ne $sprintresult.psobject.Properties['startdate'] -and $sprintresult.originBoardId -eq "115") {
            $SprintDetails = New-SprintDetails
            $SprintDetails.Name = $sprintresult.name
            $SprintDetails.ID = $sprintresult.id
            $SprintDetails.startDate = $sprintresult.startDate
            $SprintDetails.endDate = $sprintresult.endDate
            $SprintDetails.status = $sprintresult.state
            $workupdates += $SprintDetails
        }
    }
    catch {
        
    }
    

}
$forDate = Get-Date
$forDateUTC = [DateTime]::SpecifyKind($forDate.ToUniversalTime(), [DateTimeKind]::Utc)
$forDateUTC3 = $forDateUTC.ToString("MM/dd/yyyy hh:mm:ss tt")
$sprint = $workupdates |Where-Object{[DateTime]$_.startDate -lt $forDateUTC3 -and [DateTime]$_.endDate -gt $forDateUTC3}
$sprint
    # $sprintInfo = New-SprintInfo
    # $sprintInfo.Name = $sprint.Name
    # $sprintInfo.ID = $sprint.ID
    # $sprintInfo.startDate = $sprint.startDate
    # $sprintInfo.endDate = $sprint.endDate
    # $sprintInfo.status = $sprint.status
    
    # return $sprintInfo
    # Note, sprint always starts as 00:00:00 UTC. This is the sprint start time in VSTS.
