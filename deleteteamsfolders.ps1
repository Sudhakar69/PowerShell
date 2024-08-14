$path = "C:\Users\itadmin\AppData\Local\Microsoft"
$items=Get-ChildItem -Path $path
if ($items.name -match "TeamsMeetingAddin" -or $items.name -match "TeamsPresenceAddin") {
    Write-Host "We found TeamsMeetingAddin and/or TeamsPresenceAddin "
    $items |Where-Object{$_.name -match "TeamsMeetingAddin" -or $_.name -match "TeamsPresenceAddin"}|Remove-Item
    Write-Host "we are deleting TeamsMeetingAddin and/or TeamsPresenceAddin"
}
else {
    Write-Host "we are unable to find TeamsMeetingAddin or TeamsPresenceAddin"
}

