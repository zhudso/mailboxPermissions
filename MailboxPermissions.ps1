#Future configurations.
#-----------------------
#Connect to EAC with new authenctation built in the script.
#Provided numbered choices to the user to enter to prevent and misspellings.
#Add other functions, like email forwarding and auto reply
#Change if, elseif, else to switch statements.


$mailbox = Read-Host "Who's mailbox do we need get access to?"
$requestersMailbox = Read-Host "Who needs access to this mailbox"
$accessRights = Read-Host "Which access level (FullAccess, SendAs, SendonBehalf)"
$delay = 8
 
#Testing stored variables in IF statement
Write-Host "Checking for correct inputs.."

if ($accessRights -eq "FullAccess") {
    Write-Output "Successful Input for FullAccess.. Now configuring"
    Add-MailboxPermission -Identity $mailbox -User $requestersMailbox -AccessRights $accessRights
}
elseif ($accessRights -eq "SendAs") {
    Write-Output "Successful Input for Send As.. Now configuring"
    Add-MailboxPermission -Identity $mailbox -User $requestersMailbox -AccessRights $accessRights
}
elseif ($accessRights -eq "SendonBehalf") {
    Write-Output "Successful Input for SendonBehalf.. Now configuring"
    Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo @{Add=$requestersMailbox}
}
#Ask the user again for another input
else {
    Write-Output """$accessRights"" Is an incorrect input, running script again."
    $scriptToRun = $psScriptRoot+".\MailboxPermissions.ps1"
    Invoke-Expression -Command $scriptToRun
    break
}

Write-Host ""$requestersMailbox" "$accessRights" is now configured!"
Write-Host "Giving EAC time to reflect the new permissions.."

while ($delay -ge 0)
{
  Write-Host "Seconds Remaining: $($delay)"
  start-sleep 1
  $delay -= 1
  
}

Write-Host "Checking our work.. "
$getMailboxpermissions = Get-MailboxPermission -Identity $mailbox -User $requestersMailbox
Write-Output ($getMailboxpermissions)
