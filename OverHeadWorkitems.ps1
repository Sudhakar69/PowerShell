param ([string]$info="ugNfDglt2opgaXvaBEpO")
<#

Project Reporting script for VSTS

This script sends emails to all developers, testers, and product managers notifying them of 
pending OverHead workItems to remind them to take neccessary acctions



#>
Add-Type -AssemblyName System.Web
# Get the script directory
$path= "common.ps1"
# # Load and execute the common library script
$commonLibContent = Get-Content -Path $Path -Raw
Invoke-Expression -Command $commonLibContent

$WorkitemSummaryHtml = "<h1>Open OverHead Tasks in Completed Sprints</h1>"
$WorkitemSummaryHtml +="
    <table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
    <tr style='font-size:12px;'>
        <th valign='top' nowrap><b>Person</b></th>
        <th valign='top' nowrap><b>WorkItem Type</b></th>
        <th valign='top' ><b>Title</b></th>        
    </tr>
"
function Get-WorkItemUpdate {
    $WorkItemUpdate = New-Object -TypeName PSObject
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name person -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name state -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name WorkItemType -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name WorkItemId -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name Sprint -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name Title -Value $null
    $WorkItemUpdate| Add-Member -MemberType NoteProperty -Name mailid -Value $null
    return $WorkItemUpdate
    
}
function Get-SprintDetails {
    $SprintDetails = New-Object -TypeName PSObject
    $SprintDetails| Add-Member -MemberType NoteProperty -Name startDate -Value $null
    $SprintDetails| Add-Member -MemberType NoteProperty -Name endDate -Value $null
    $SprintDetails| Add-Member -MemberType NoteProperty -Name sprintNumber -Value $null
    return $SprintDetails
    
}

$PATId = "PAT ID"
$queryId = "Query ID"

$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($PATId)"))
}
# Convert the query to JSON
# $jsonQuery = $query | ConvertTo-Json
$latestSprint =Get-SprintInfo -forDate (Get-Date)
$latestSprintNumber = $latestSprint.sprintNumber
$sprintsCount = $latestSprintNumber/2
$presentsprint = "$($Project)\Sprint " + $latestSprintNumber
$startDate = (Get-Date).AddMonths(-$sprintsCount).Date
$endDate = Get-Date -DisplayHint Date
$sprintcurrentDate = $startDate
$Sprintslists = @()
while ($sprintcurrentDate -le $endDate) {
    $Sprintslist = Get-SprintInfo -forDate ($sprintcurrentDate)
    $sprintcurrentDate = $sprintcurrentDate.AddDays(14)
    $SprintDetails = Get-SprintDetails
    $SprintDetails.startDate = $Sprintslist.startDate
    $SprintDetails.endDate = $Sprintslist.endDate
    $SprintDetails.sprintNumber = $Sprintslist.sprintNumber
    $Sprintslists += $SprintDetails
}
$WorkitemsQueryUri = "https://$($organization).visualstudio.com/DefaultCollection/_apis/wit/wiql/$($queryId)?api-version=4.1"
$WorkitemsQueryResult = Invoke-RestMethod -Uri $WorkitemsQueryUri -Method Get -Headers $headers 
# $workitemrecords =@()
# $workItemRelationsCount = $WorkitemsQueryResult.workItemRelations.Count
$workItemIds = $WorkitemsQueryResult.workItemRelations
$workupdates = @()
foreach($workItemId in $workItemIds.target.id)
{
    # $sourceWorkItemId = $projectReportingQueryResult.workItemRelations[$i].source.id
    # $targetWorkItemId = $WorkitemsQueryResult.workItemRelations[$i].target.id
    
    $workuri = "https://dev.azure.com/$($organization)/$($Project)/_apis/wit/workitems?ids=$($workItemId)&api-version=7.1-preview.3"
    $workresponse= Invoke-RestMethod -Uri $workuri -Method Get -Headers $headers 
    $workItemResponse = $workresponse.value
    for($i=0;$i -lt $workItemResponse.Count;$i++){
        # $workupdates = $workItemResponse[$i].fields #|Where-Object {$_.'System.IterationPath' -eq "$($Project)\Sprint 173"} 
        $WorkItemUpdate = Get-WorkItemUpdate
        if ($null -ne $workItemResponse[$i].fields.'System.AssignedTo'.displayName) {
            $WorkItemUpdate.person = $workItemResponse[$i].fields.'System.AssignedTo'.displayName
            $WorkItemUpdate.mailid = $workItemResponse[$i].fields.'System.AssignedTo'.uniqueName
        }
        else {
            $WorkItemUpdate.person = "Unassigned"
        }
        
        $WorkItemUpdate.state=  $workItemResponse[$i].fields.'System.state'
        $WorkItemUpdate.WorkItemType = $workItemResponse[$i].fields.'System.WorkItemType'
        $WorkItemUpdate.sprint = $workItemResponse[$i].fields.'System.IterationPath'
        $WorkItemUpdate.Title = $workItemResponse[$i].fields.'System.Title'
        $WorkItemUpdate.WorkItemId = $workItemId
        $workupdates += $WorkItemUpdate
    }
}

$openWorkItemslists =  $workupdates |Where-Object{$_.Title -match "OverHead" -or $_.Title -match " OH "}| Where-Object { $_.state -eq "New" -or $_.state -eq "To Do" -or $_.state -eq "In Progress" -or $_.state -eq "Ready for Deployment" -or $_.state -eq "Ready for Testing" }

# $emails=@()
$emails=("c-tgudise@xpansiv.com")
$sortedworkitems= $openWorkItemslists | Sort-Object -Property person,workItemType -Descending
foreach($sortedworkitem in $sortedworkitems){
    $Openworkitemperson = $sortedworkitem.person
    # $emails += $sortedworkitem.mailid
    $workItemtype = $sortedworkitem.workItemType
    $openWorkItemId = $sortedworkitem.workItemId
    $WorkItemtitle = $sortedworkitem.title
    $WorkitemSummaryHtml += " 
                <tr style='font-size:12px;' nowrap>
                    <td nowrap>$($Openworkitemperson)</td>
                    <td nowrap>$($workItemtype)</td>
                    <td nowrap><a href='https://$($organization).visualstudio.com/$($Project)/_workitems/edit/$($openWorkItemId)'>$($openWorkItemId)<b> - </b>$($WorkItemtitle)</a></td>
                </tr>
"
}

$WorkitemSummaryHtml += "</table></tbody>"

$body = $WorkitemSummaryHtml

$emailslist = $emails |Select-Object -Unique
$emailSubject = "Open Overhead Tasks"
Start-Sleep -s 15

$SMTP_SERVER = "smtp.socketlabs.com"
$SMTP_PORT = 587
$SMTP_USERNAME = "server4507"
$PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force
$bccEmailList=("")
# $bccEmailList=("")
# $cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
$SMTPClient = New-Object System.Net.Mail.SmtpClient
$SMTPClient.Host = $SMTP_SERVER
$SMTPClient.Port = $SMTP_PORT
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTP_USERNAME, $PWORD)
$message= New-Object System.Net.Mail.MailMessage
$message.From = " "
foreach($mailid in $emailslist)
{
    $message.To.Add($mailid)
}
$message.Subject = $emailSubject
$message.Body = $body
foreach($bccEmail in $bccEmailList){
    $message.cc.Add($bccEmail)
}

$message.IsBodyHtml = $true
# Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From " " -Subject $emailSubject -Body $body -BodyAsHtml
try {
    # Send-MailMessage -To $email -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -Bcc $bccEmailList -From " " -Subject $emailSubject -Body $body -BodyAsHtml
    $SMTPClient.Send($message)
    $SMTPClient.Dispose()
    $message.Dispose()
    
}
catch {
    $_.Exception.message
}
