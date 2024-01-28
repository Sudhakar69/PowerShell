$queryid = "mf5vsoaq5phff4oxnumysgedqgv2wc7gut2cwjkuzre3efowww5q"

$projectreport = "https://dev.azure.com/tarakapersonal/Personal/_apis/git/repositories?api-version=6.0"
$headers = @{
    Authorization= "Basic" + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($queryid)"))
}
$repositoryresponse = Invoke-RestMethod -Uri $projectreport -Method Get -Headers $headers
   # Define SMTP server settings
   $SMTP_SERVER = "smtp.office365.com"
   $SMTP_PORT = 587  # Typically 587 for TLS/STARTTLS, or 465 for SSL
   $SMTP_USERNAME = "taraka.gudise@outlook.com"
   $SMTP_PASSWORD = "MyP0!yc0m@123"
   $UseSsl = $true  # Set to $true for TLS/STARTTLS, or $false for no SSL
   
   # Create a new SMTP client
   $SMTPClient = New-Object System.Net.Mail.SmtpClient
   $SMTPClient.Host = $SMTP_SERVER
   $SMTPClient.Port = $SMTP_PORT
   $SMTPClient.EnableSsl = $UseSsl
   
   # Create credentials
   $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTP_USERNAME, $SMTP_PASSWORD)
   $body = "<h2>Sample Mail</h2>"
   $body += "<h3>Sample Body Mail</h3>"
   $matrixHTML = "<tbody>`n"
   $matrixHTML += "<table boarder='1'cellpadding='1' style='font-size:9px;border: 1px solid black; border-collapse: collapse'><tr>`n"
   $matrixHTML += "<td valign='top' ><b>Branch Name </b></td>`n"
   $matrixHTML += "<td valign='top' bgcolor='#ff0000' nowrap ><b>Behind </b></td>`n"
   $matrixHTML += "<td valign='top' bgcolor='#ff0000' nowrap ><b>Ahead </b></td>`n"
   $matrixHTML += "</tr>`n"
   $body += $matrixHTML
   
foreach($repository in $repositoryresponse.value)
{
    if ($repositoryresponse.defaultbranch -eq "refs/heads/master") {
        <# Action to perform if the condition is true #>
    
        $branch="https://dev.azure.com/tarakapersonal/Personal/_apis/git/repositories/$($repository.id)/stats/branches?baseVersionDescriptor.version=master&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
        $branchresponse = Invoke-RestMethod -Uri $branch -Method Get -Headers $headers
    }
    else #($repositoryresponse.defaultbranch -eq "refs/heads/main") 
    {
        $branch="https://dev.azure.com/tarakapersonal/Personal/_apis/git/repositories/$($repository.id)/stats/branches?baseVersionDescriptor.version=main&baseVersionDescriptor.versionOptions=none&baseVersionDescriptor.versionType=branch&api-version=7.1-preview.1"
        $branchresponse = Invoke-RestMethod -Uri $branch -Method Get -Headers $headers
        <# Action when this condition is true #>
    }
    $body += "<tr style='font-size:9px;border: 1px solid black;'><b>Repo:</b> " + $repository.name + "<b>; Default Branch:</b>" + $repository.defaultbranch + "</tr>"
    foreach($commit in $branchresponse.value)
    {
        $body += "<tr><td align='right' style='font-size:9px;border: 1px solid black; border-collapse: collapse'>" +  $commit.name +" </td>
        <td align='right' style='font-size:9px;border: 1px solid black; border-collapse: collapse'>" + $commit.behindCount +" </td>
        <td align='right' style='font-size:9px;border: 1px solid black; border-collapse: collapse'>" + $commit.aheadcount +"</td></tr>"
    }
   
# Create a new email message

}
$body += "</tbody></table>"
$Message = New-Object System.Net.Mail.MailMessage
$Message.From = $SMTP_USERNAME
$Message.To.Add("taraka.gudise@outlook.com")
$Message.Subject = "Subject"
$Message.Body = $body
$Message.IsBodyHtml = $true

# Send the email
$SMTPClient.Send($Message)

# Dispose of the SMTP client and message
$SMTPClient.Dispose()
$Message.Dispose()