param ([string]$info="PLEASE PROVIDE PASSWORD AS AN ARGUMENT TO THIS SCRIPT")

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

 

function sendMail {
    
  $EmailFrom = "tfsbuild@apx.com"                                                                                          
# $status = "Test Mail"
  $EmailTo = "c-tgudise@xpansiv.com"

  $SMTPServer = "smtp.socketlabs.com"                                                                                            

  $SMTPAuthUsername = "server4507"                                                                                                                                                                                                  

  $SMTPAuthPassword = $info                                                 

  $mailmessage = New-Object system.net.mail.mailmessage                                                                          

  $mailmessage.from = ($EmailFrom)                                                                                              

  $mailmessage.To.add($EmailTo)                                                                                                  

  $mailmessage.Subject = $status                                                                                                

  $body = $status                                                                                                                

  $mailmessage.Body = $body                                                                                                      

  $mailmessage.Priority = [System.Net.Mail.MailPriority]::High                                                                  

  $SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, 587)                                                                  

  $SMTPClient.Credentials = New-Object System.Net.NetworkCredential("$SMTPAuthUsername", "$SMTPAuthPassword")                    

  $SMTPClient.Send($mailmessage)                                                                                                

                                                                                                                                           

}

$status = "Test Mail - Pedro "

sendMail