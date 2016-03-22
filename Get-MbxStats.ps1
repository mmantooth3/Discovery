$Report=@()
$mailbox=Get-mailbox -resultsize unlimited 
$mailbox| foreach-object{$DisplayName=$_.DisplayName; $SmtpAddress=$_.PrimarySmtpAddress; $mbstats=(get-mailboxstatistics -identity $DisplayName); $obj=new-object System.Object; $obj|add-member -membertype NoteProperty -name "DisplayName" -value $DisplayName; $obj|add-member -membertype NoteProperty -name "SamAccountName" -value $_.SamAccountName; $obj|add-member -membertype NoteProperty -name "Alias" -value $_.Alias;$obj|add-member -membertype NoteProperty -name "PrimarySmtpAddress" -value $_.PrimarySmtpAddress;$obj|add-member -membertype NoteProperty -name "ServerName" -value $mbstats.ServerName; $obj|add-member -membertype NoteProperty -name "Database" -value $mbstats.Database; $obj|add-member -membertype NoteProperty -name "OrganizationalUnit" -value $_.OrganizationalUnit; $obj|add-member -membertype NoteProperty -name "ItemCount" -value $mbstats.ItemCount; $obj|add-member -membertype NoteProperty -name "TotalItemSize" -value $mbstats.TotalItemSize; $obj|add-member -membertype NoteProperty -name "TotalDeletedItemSize" -value $mbstats.TotalDeletedItemSize; $Report+=$obj}
$Report|export-csv mailboxstatistics.csv -notype
