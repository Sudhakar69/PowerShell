param ([string]$info="Please provide password")
$queryId = "Query ID"
function Get-PullRequestInfo {
    $PullRequestInfo = New-Object -TypeName PSObject
    $PullRequestInfo | Add-Member -MemberType NoteProperty -Name RepoName -Value $null
    $PullRequestInfo | Add-Member -MemberType NoteProperty -Name PullRequestId -Value $null
    $PullRequestInfo | Add-Member -MemberType NoteProperty -Name creationDate -Value $null
    $PullRequestInfo | Add-Member -MemberType NoteProperty -Name createdby -Value $null
    $PullRequestInfo | Add-Member -MemberType NoteProperty -Name Reviewedby -Value $null
    $PullRequestInfo | Add-Member -MemberType NoteProperty -Name PullrequestStatus -Value $null

    return $PullRequestInfo
    
}
$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryId)"))
}
$projectReportingQueryUri = "https://dev.azure.com/$($Organization)/$($Project)/_apis/git/repositories?api-version=6.0"
$repositoriesResponse = Invoke-RestMethod -Uri $projectReportingQueryUri -Headers $headers -Method Get
$RepoSummaryHtml = "<h1>PullRequest status by Repo</h1>"

$RepoSummaryHtml += "<tbody><table border='1' cellspacing='1' style='font-size:10px; border:1px solid black; border-collapse: collapse' nowrap>
    <tr style='font-size:12px;'>
        <th valign='top' nowrap><b>Repo Name</b></th>
        <th valign='top' nowrap><b>Pull Request ID</b></th>
        <th valign='top' nowrap><b>Created Time(UTC)</b></th>
        <th valign='top' nowrap><b>Created by</b></th>
        <th valign='top' nowrap><b>Needs to be <br>Reviewed/Approved By</b></th>
        <th valign='top' nowrap><b>Pullrequest Status</b></th>
    </tr>
"
$PullRequestData = @()
$sortedRepositories = $repositoriesResponse.value |Sort-Object -Property name
foreach ($repositoryname in $sortedRepositories.name) { 
    $repository = $sortedRepositories |Where-Object{$_.name -eq $repositoryname}
    $RepoUri = "https://dev.azure.com/$($Organization)/$($Project)/_apis/git/repositories/$($repository.Id)/pullrequests?api-version=7.2-preview.2"
    $reporesponse = Invoke-RestMethod -Uri $repouri -Method Get -Headers $headers
    $repodata= $reporesponse.value#|Where-Object{$_ -ne $null}
    $reponame = $repositoryname
    if ($null -ne $repodata.creationDate -or $null -ne $repodata.psobject.Properties['creationDate']) {
        $status = $repodata.status
        try {
            $createdDate = $repodata.creationDate
            if ($null -ne $repodata.Reviewer) {
                $reviewedby = $repodata.reviewer.displayName
            }
            else {
                $reviewedby = $null
            }
        }
        catch {
            $_.Exception.Message
        }
        
        $createdBy = $repodata.createdby.displayName
        $PullRequestInfo = Get-PullRequestInfo
        $PullRequestInfo.RepoName = $reponame
        $PullRequestInfo.PullrequestId = $repodata.pullRequestId
        $PullRequestInfo.creationDate = $createdDate
        $PullRequestInfo.createdby = $createdBy
        $PullRequestInfo.Reviewedby = $reviewedby
        $PullRequestInfo.PullrequestStatus = $status
        $PullRequestData += $PullRequestInfo
        
    }
    else {
        
    }
    
}
$sortedPullrequest = $PullRequestData | Sort-Object -Property creationDate
foreach($PullRequest in $sortedPullrequest){
    $RepoSummaryHtml += "<tr nowrap>
                <td nowrap>$($PullRequest.reponame)</a></td>
                <td nowrap><a href='https://$($Organization).visualstudio.com/$($Project)/_git/$($PullRequest.reponame)/pullrequest/$($PullRequest.pullRequestId)'>$($PullRequest.pullRequestId)</a></td>
                <td nowrap>$($PullRequest.creationDate)</td>
                <td nowrap>$($PullRequest.createdBy)</td>
                <td snowrap>$($PullRequest.reviewedby)</td>
                <td nowrap>$($PullRequest.PullrequestStatus)</td>
            </tr>"
}
$RepoSummaryHtml += "</table></tbody>"
$body += $RepoSummaryHtml
# # compose and send out email messages to individuals who have booked to at least one task in the sprint
Write-Host "-------------------------------------------------------------------------------------------"
Write-Host "Sending out the Outstanding pull requests Details email"
$emailSubject = "Outstanding pull requests Details"
Start-Sleep -s 15

$SMTP_SERVER = "smtp.socketlabs.com"
$SMTP_PORT = 587
$SMTP_USERNAME = "server4507"
$PWORD = ConvertTo-SecureString -String $info -AsPlainText -Force
$bccEmailList = ("nkln@xpansiv.com")
$email = ("kn@$($Project).com")
$SMTPClient = New-Object System.Net.Mail.SmtpClient
$SMTPClient.Host = $SMTP_SERVER
$SMTPClient.Port = $SMTP_PORT
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTP_USERNAME, $PWORD)
$message= New-Object System.Net.Mail.MailMessage
$message.From = "njb@$($Project).com"
foreach($mailid in $bccEmailList)
{
    $message.Bcc.Add($mailid)
}
$message.To.Add($email)
$message.CC.Add($bccEmailList)
$message.Subject = $emailSubject
$message.Body = $body
$message.IsBodyHtml = $true
try {
    $SMTPClient.Send($message)
    $SMTPClient.Dispose()
    $message.Dispose()
    $body = $null
    $RepoSummaryHtml = $null
    $PullRequestData = $null

}
catch {
    $_.Exception.message
}

Get-date

