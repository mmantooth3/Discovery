$OutFile = "C:\Temp\PermissionExport.txt"
"DisplayName" + "," + "Alias" + "," + "Full Access" + "," + "Send As" | Out-File $OutFile -Force
 
$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Select Identity, Alias, DisplayName, DistinguishedName
ForEach ($Mailbox in $Mailboxes) {
              $SendAs = Get-ADPermission $Mailbox.DistinguishedName | ? {$_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITY\SELF" -and !$_.IsInherited} | % {$_.User}
              $FullAccess = Get-MailboxPermission $Mailbox.Identity | ? {$_.AccessRights -eq "FullAccess" -and !$_.IsInherited} | % {$_.User}
 
              $Mailbox.DisplayName + "," + $Mailbox.Alias + "," + $FullAccess + "," + $SendAs | Out-File $OutFile -Append
}
