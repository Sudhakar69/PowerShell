param ([string]$info="Password")
<#

Project Reporting script for VSTS

This script sends emails to all developers, testers, and product managers notifying them
of Open workitems list and send a gentle reminder to the Workitem owner and respective leads to say hey from these Many days your workitem is pending


#>
Add-Type -AssemblyName System.Web

# Get the script directory
$path= "\common\common.ps1"
# # Load and execute the common library script
$commonLibContent = Get-Content -Path $Path -Raw
Invoke-Expression -Command $commonLibContent

function Get-WorkItems {
    $WorkItems = New-Object -TypeName PSObject
    $WorkItems| Add-Member -MemberType NoteProperty -Name WorkitemId -Value $null
    $WorkItems| Add-Member -MemberType NoteProperty -Name WorkItemType -Value $null
    $WorkItems| Add-Member -MemberType NoteProperty -Name Description -Value $null
    $WorkItems| Add-Member -MemberType NoteProperty -Name Tags -Value $null
    $WorkItems| Add-Member -MemberType NoteProperty -Name State -Value $null
    $WorkItems| Add-Member -MemberType NoteProperty -Name AssignedTo -Value $null
    $WorkItems| Add-Member -MemberType NoteProperty -Name ParentId -Value $null
    return $WorkItems
    
}

$queryId = "queryId"
$PATId = "Access token"


$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($PATId)"))
}
$latestSprint =Get-SprintInfo -forDate (Get-Date)
$latestSprintNumber = $latestSprint.sprintNumber
$presentsprint = "$($project)\Sprint " + $latestSprintNumber
$WorkitemsQueryUri = "https://$($organization).visualstudio.com/DefaultCollection/_apis/wit/wiql/$($queryId)?api-version=4.1"
$WorkitemsQueryResult = Invoke-RestMethod -Uri $WorkitemsQueryUri -Method Get -Headers $headers 

# $workItemRelationsCount = $WorkitemsQueryResult.workItemRelations.Count
$workItemIds = $WorkitemsQueryResult.workItemRelations
$notapprovedWorkitemHtml += "<br>"
$notapprovedWorkitemHtml += "<b>Current Sprint: &nbsp;</b>$presentsprint"
$approvedWorkitemHtml += "<br>"
$approvedWorkitemHtml += "<b>Approved List</b>"
$approvedWorkitemHtml +="<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>Task/PBI ID</b></th>
            <th valign='top' nowrap><b>Work Item Type</b></th>
            <th valign='top' nowrap><b>Task/PBI Description</b></th>
            <th valign='top' nowrap><b>Tags</b></th>
            <th valign='top' nowrap><b>Approval status</b></th>
            <th valign='top' nowrap><b>State</b></th>
            <th valign='top' nowrap><b>Assigned To</b></th>
            <th valign='top' nowrap><b>Parent ID</b></th>
        
        </tr>"
$notapprovedWorkitemHtml += "<br>"
$notapprovedWorkitemHtml += "<b>Not Approved List</b>"
$notapprovedWorkitemHtml +="<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
        <tr style='font-size:12px;'>
            <th valign='top' nowrap><b>Task/PBI ID</b></th>
            <th valign='top' nowrap><b>Work Item Type</b></th>
            <th valign='top' nowrap><b>Task/PBI Description</b></th>
            <th valign='top' nowrap><b>Tags</b></th>
            <th valign='top' nowrap><b>Approval status</b></th>
            <th valign='top' nowrap><b>State</b></th>
            <th valign='top' nowrap><b>Assigned To</b></th>
            <th valign='top' nowrap><b>Parent ID</b></th>
        
        </tr>"
$workItemIdslist = @()
for ($i = 0; $i -lt $workItemIds.Count; $i++) {
    $workItemIdslist += [Int64]$workItemIds[$i].target.id
}
Write-Host "Work Item IDs has been collected and going to next step"
$sortedworkItemIdslist = $workItemIdslist|Select-Object -Unique |Sort-Object -Descending
$WorkItemsData = @()
foreach($workItemId in $sortedworkItemIdslist)
{
    try {
        # $workItemuri="https://dev.azure.com/$($organization)/$($project)/_apis/wit/workitems?ids=$($workItemId)&api-version=7.1-preview.3"
        # $workItemResponse = Invoke-RestMethod -Uri $workItemuri -Method Get -Headers $headers 
        $workuri = "https://dev.azure.com/$($organization)/$($project)/_apis/wit/workItems/$($workItemId)/updates"
        $workresponseresult= Invoke-RestMethod -Uri $workuri -Method Get -Headers $headers 
    }
    catch {
        $_.Exception.Message
        Get-Date
    }
    
    $workItemResponseResult = $workresponseresult.value
    if ($null -ne $workItemResponseResult.fields.'System.parent'.PSObject.Properties["oldvalue"]) {
        $parent = $workItemResponseResult
    }
    else {
        $parent = $workItemResponseResult[0]
    }
    # $parent = $workItemResponseResult[0]
    $child = $workItemResponseResult|Where-Object{$_.relations.added.attributes.name -match "Child"}
    $count = $child.count
    if ($count -gt 0) {
        $childurls = $child.relations.added.url
        $childItemResponseresult = @()
        foreach($childurl in $childurls){
            $childitemId = $childurl.split('/')[-1]
            $childitemurl = "https://dev.azure.com/$($organization)/$($project)/_apis/wit/workItems/$($childitemId)/updates"
            $childItemresult= Invoke-RestMethod -Uri $childitemurl -Method Get -Headers $headers
            #$null -ne $_.fields.PSObject.Properties['System.Tags'] -or 
            $childItemResponse = $childItemresult.value|Where-Object{$_.fields.'System.Tags'.newValue -ne $null}
            $childItemResponseresult += $childItemResponse
        }
    }
    # if ($null -ne $parent.fields.'System.Tags'.PSObject.Properties["oldvalue"]) {
    #     $tags = $parent.fields.'System.Tags'.newValue
    # }
    
    
    if ($parent.fields.'System.Tags'.newValue -ne "" -and $parent.fields.'System.Tags'.newValue -match "Hot Fix" -or $parent.fields.'System.Tags'.newValue -match "Data Fix") {
        if ($parent.fields.'System.Tags'.newValue -match "Change Request" -and $parent.fields.'System.state'.newValue -ne "Removed") {
            $childItemresult=$childItemResponseresult|Where-Object{$_.fields.'System.Tags'.newValue -notmatch "Change Request"-or $_.fields.'System.Tags'.newValue -notmatch "Hot Fix" -or $_.fields.'System.Tags'.newValue -notmatch "Data Fix"}
            if ($parent.fields.'System.WorkItemType'.newValue -match "Product Backlog Item" -or $parent.fields.'System.WorkItemType'.newValue -match "Bug") {
                if ($childItemresult.count -eq 0) {
                    $tags = $parent.fields.'System.Tags'.newValue
                    $state = $parent.fields.'System.state'.newValue
                    $title = $parent.fields.'System.Title'.newValue
                    $Parentdata = $parent.fields.'System.Parent'.newValue
                    $workitemtype = $parent.fields.'System.WorkItemType'.newValue
                    $Assigned = $parent.fields.'System.AssignedTo'.newValue
                    $AssignedTo = $Assigned.displayName
                    $WorkItems = Get-WorkItems
                    $WorkItems.WorkitemId = $workItemId
                    $WorkItems.WorkItemType = $workitemtype
                    $WorkItems.Description = $title
                    $WorkItems.Tags = $tags
                    $WorkItems.State = $state
                    $WorkItems.AssignedTo = $AssignedTo
                    $WorkItems.ParentId = $Parentdata
                    $WorkItemsData += $WorkItems
                    if ($parent.fields.'System.state'.newValue -eq "Done") {
                        $Commentsuri = "https://dev.azure.com/$($organization)/$($project)/_apis/wit/workItems/$($workItemId)/comments?$top=2&api-version=7.1-preview.4"
                        $Commentresponse= Invoke-RestMethod -Uri $Commentsuri -Method Get -Headers $headers 
                        $Commentsresult = $Commentresponse|Where-Object{$_.comments.Text -match "approve"}
                        if ($null -ne $Commentsresult) {
                            $approvedstatus = "Yes/Approved"
                            $approvedWorkitemHtml += "<tr style='font-size:12px;' nowrap>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($workItemId)'>$($workItemId)</a></td>
                                        <td align ='left' nowrap>$($workitemtype)</td> 
                                        <td align ='left' nowrap>$($title)</td> 
                                        <td align ='left' nowrap>$($tags)</td>
                                        <td align ='left' nowrap>$($approvedstatus)</td>
                                        <td align ='left' nowrap>$($state)</td>
                                        <td align ='left' nowrap>$($AssignedTo)</td> 
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($Parentdata)'>$($Parentdata)</a></td> 
                                    </tr>" 
                        }
                        else {
                            $approvedstatus = "No/Not Approved"
                            $notapprovedWorkitemHtml += "<tr style='font-size:12px;' nowrap>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($workItemId)'>$($workItemId)</a></td>
                                        <td align ='left' nowrap>$($workitemtype)</td> 
                                        <td align ='left' nowrap>$($title)</td> 
                                        <td align ='left' nowrap>$($tags)</td>
                                        <td align ='left' nowrap>$($approvedstatus)</td>
                                        <td align ='left' nowrap>$($state)</td>
                                        <td align ='left' nowrap>$($AssignedTo)</td> 
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($Parentdata)'>$($Parentdata)</a></td>
                                    </tr>" 
                        }
                            
                    } 
                    elseif ($parent.fields.'System.state'.newValue -eq "To Do" -or $parent.fields.'System.state'.newValue -eq "In Progress") {
                        if ($parent.fields.'System.WorkItemType'.newValue -ne "Task") {
                            $Commentsuri = "https://dev.azure.com/$($organization)/$($project)/_apis/wit/workItems/$($workItemId)/comments?$top=2&api-version=7.1-preview.4"
                            $Commentresponse= Invoke-RestMethod -Uri $Commentsuri -Method Get -Headers $headers 
                            $Commentsresult = $Commentresponse|Where-Object{$_.comments.Text -match "approve"}
                            if ($null -ne $Commentsresult) {
                                $approvedstatus = "Yes/Approved"
                                $approvedWorkitemHtml += "<tr style='font-size:12px;' nowrap>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($workItemId)'>$($workItemId)</a></td>
                                        <td align ='left' nowrap>$($workitemtype)</td> 
                                        <td align ='left' nowrap>$($title)</td> 
                                        <td align ='left' nowrap>$($tags)</td>
                                        <td align ='left' nowrap>$($approvedstatus)</td>
                                        <td align ='left' nowrap>$($state)</td> 
                                        <td align ='left' nowrap>$($AssignedTo)</td>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($Parentdata)'>$($Parentdata)</a></td>
                                    </tr>" 
                            }
                            else {
                                $approvedstatus = "No/Not Approved"
                                $notapprovedWorkitemHtml += "<tr style='font-size:12px;' nowrap>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($workItemId)'>$($workItemId)</a></td>
                                        <td align ='left' nowrap>$($workitemtype)</td> 
                                        <td align ='left' nowrap>$($title)</td> 
                                        <td align ='left' nowrap>$($tags)</td>
                                        <td align ='left' nowrap>$($approvedstatus)</td>
                                        <td align ='left' nowrap>$($state)</td> 
                                        <td align ='left' nowrap>$($AssignedTo)</td>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($Parentdata)'>$($Parentdata)</a></td>
                                    </tr>" 
                            }
                        }
                        else {
                            $Commentsuri = "https://dev.azure.com/$($organization)/$($project)/_apis/wit/workItems/$($workItemId)/comments?$top=2&api-version=7.1-preview.4"
                            $Commentresponse= Invoke-RestMethod -Uri $Commentsuri -Method Get -Headers $headers 
                            $Commentsresult = $Commentresponse|Where-Object{$_.comments.Text -match "approve"}
                            if ($null -ne $Commentsresult) {
                                $approvedstatus = "Yes/Approved"
                                $approvedWorkitemHtml += "<tr style='font-size:12px;' nowrap>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($workItemId)'>$($workItemId)</a></td>
                                        <td align ='left' nowrap>$($workitemtype)</td> 
                                        <td align ='left' nowrap>$($title)</td> 
                                        <td align ='left' nowrap>$($tags)</td>
                                        <td align ='left' nowrap>$($approvedstatus)</td>
                                        <td align ='left' nowrap>$($state)</td> 
                                        <td align ='left' nowrap>$($AssignedTo)</td>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($Parentdata)'>$($Parentdata)</a></td>
                                    </tr>" 
                            }
                                
                        }
                    } 
                    else {
                        $Commentsuri = "https://dev.azure.com/$($organization)/$($project)/_apis/wit/workItems/$($workItemId)/comments?$top=2&api-version=7.1-preview.4"
                        $Commentresponse= Invoke-RestMethod -Uri $Commentsuri -Method Get -Headers $headers 
                        $Commentsresult = $Commentresponse|Where-Object{$_.comments.Text -match "approve"}
                        if ($null -ne $Commentsresult) {
                            $approvedstatus = "Yes/Approved"
                            $approvedWorkitemHtml += "<tr style='font-size:12px;' nowrap>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($workItemId)'>$($workItemId)</a></td>
                                        <td align ='left' nowrap>$($workitemtype)</td> 
                                        <td align ='left' nowrap>$($title)</td> 
                                        <td align ='left' nowrap>$($tags)</td>
                                        <td align ='left' nowrap>$($approvedstatus)</td>
                                        <td align ='left' nowrap>$($state)</td> 
                                        <td align ='left' nowrap>$($AssignedTo)</td>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($Parentdata)'>$($Parentdata)</a></td>
                                    </tr>" 
                        }
                        else {
                            $approvedstatus = "No/Not yet Approved"
                            $notapprovedWorkitemHtml += "<tr style='font-size:12px;' nowrap>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($workItemId)'>$($workItemId)</a></td>
                                        <td align ='left' nowrap>$($workitemtype)</td> 
                                        <td align ='left' nowrap>$($title)</td> 
                                        <td align ='left' nowrap>$($tags)</td>
                                        <td align ='left' nowrap>$($approvedstatus)</td>
                                        <td align ='left' nowrap>$($state)</td>
                                        <td align ='left' nowrap>$($AssignedTo)</td>
                                        <td nowrap><a href='https://$($organization).visualstudio.com/$($project)/_workitems/edit/$($Parentdata)'>$($Parentdata)</a></td> 
                                    </tr>" 
                        }
                    }
                }
                $childItemresult = $null
            }     
        }
        Write-Host "Data Pulled for " $workItemId   
    }    
}
# $uniqueParentId = $WorkItemsData

$approvedWorkitemHtml += "</table></tbody>"
$notapprovedWorkitemHtml += "</table></tbody>"
    
    
$body += $notapprovedWorkitemHtml
$body += $approvedWorkitemHtml
    

    # # compose and send out email messages to individuals who have booked to at least one task in the sprint
    Write-Host "-------------------------------------------------------------------------------------------"
    Write-Host "Sending out email " 
    # $email=("")
$emailSubject = "Incorrect Change Request Status V2.1"
    Start-Sleep -s 15

$SMTP_SERVER = "smtp.socketlabs.com"
$SMTP_PORT = 587
$SMTP_USERNAME = "server4507"
$PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force
$bccEmailList = ("")
$email = ("")
    # $cred = New-Object -TypeName System.Management.Automation.PSCredential($SMTP_USERNAME, $PWORD)
$SMTPClient = New-Object System.Net.Mail.SmtpClient
$SMTPClient.Host = $SMTP_SERVER
$SMTPClient.Port = $SMTP_PORT
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTP_USERNAME, $PWORD)
$message= New-Object System.Net.Mail.MailMessage
$message.From = "tfsbuild@$($project).com"
    foreach($mailid in $bccEmailList)
    {
    $message.Bcc.Add($mailid)
    }
$message.To.Add($email)
$message.Subject = $emailSubject
$message.Body = $body
    # $message.cc.Add($bccEmailList)
$message.IsBodyHtml = $true
    # Send-MailMessage -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -To $email -Bcc $bccEmailList -From tfsbuild@$($project).com -Subject $emailSubject -Body $body -BodyAsHtml
    try {
        # Send-MailMessage -To $email -SmtpServer $SMTP_SERVER -Port $SMTP_PORT -Credential $cred -Bcc $bccEmailList -From tfsbuild@$($project).com -Subject $emailSubject -Body $body -BodyAsHtml
        $SMTPClient.Send($message)
        $SMTPClient.Dispose()
        $message.Dispose()
        $body = $null
        $approvedWorkitemHtml = $null
        $notapprovedWorkitemHtml = $null
        $childItemResponseresult = $null
        
    }
    catch {
        $_.Exception.message
    }


