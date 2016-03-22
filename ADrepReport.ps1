#========================================================================
# Created on:   08/19/2015
# Created by:   EN Pointe Messaging Team 
# Filename:     Repadmin-ShowRepl-HTMLReports.ps1
#
# Description:	Generate HTML reports based on 
#				"repadmin /showrepl * /csv" command return,
#				format it and send it by email
#========================================================================


###################################################
#		VARIABLES
###################################################

# Path of CSV/HTML exports
$ExportPath = "c:\temp"

# Path of Repadmin binary (only necessary if not found in $env:path)
$RepadminExec = $null

# Random name of CSV/HTML Export
$BaseFilename = "Repadmin-"+(Get-Date -Format yyMMdd)+"-"+(Get-Random)
$CSVFilename = $BaseFilename+".csv"
$HTMLFilename = $BaseFilename+".html"

# CSV Headers
$Header = "showrepl_COLUMNS","Destination DSA Site","Destination DSA","Naming Context","Source DSA Site","Source DSA", `
			"Transport Type","Number of Failures","Last Failure Time","Last Success Time","Last Failure Status"

# HTML Style
$style = @"
	<style>
		BODY{font-family:"Segoe UI"}
		P{text-decoration:underline;text-indent:40px;font-size:20px;color:DarkSlateGray;font-weight:bold}
		TABLE{border-width: 2px;border-style: solid;border-color: Black;border-collapse: collapse;}
		TABLE.error{border-color:red}
		TABLE.warning{border-color:darkorange}
		TABLE.success{border-color:green}
		TH{border-width: 2px;padding: 0px;border-style: solid;padding:5px}
		TD{border-width: 1px;padding: 0px;border-style: solid;text-align:center}
		TH.error{border-color:red;background-color:salmon}
		TH.warning{border-color:darkorange;background-color:orange}
		TH.success{border-color:green;background-color:limegreen}
		TD.error{border-color: red}
		TD.warning{border-color:darkorange}
		TD.success{border-color:green}
	</style>
"@

# Email notification parameters
$Sender = "zrjbutt@enpointe.com"
$Recipient = "rjbutt@enpointe.com"
# You can specify an array if needed 
$Server = "mail.enpointe.com"
$subject = "Repadmin HTML Report"


###################################################
#		FUNCTIONS
###################################################

function EmailNotification($Sender, $Recipient, $Server, $Subject, $Body)
{
	$SMTPclient = new-object System.Net.Mail.SmtpClient $Server

	# SMTP Port (if needed)
	# $SMTPClient.port = 587

	# Enabling SSL (if needed)
	# $SMTPclient.EnableSsl = $true

	# Specify authentication parameters (if needed)
	$SMTPAuthUsername = "zrjbutt@enpointe.com"
	$SMTPAuthPassword = "Mcitp@oo*"
	$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTPAuthUsername, $SMTPAuthPassword)

	$Message = new-object System.Net.Mail.MailMessage
	$Message.From = $Sender
	$Recipient | %{ $Message.To.Add($_) }
	$Message.Subject = $Subject
	$Message.Body = $Body
	$Message.IsBodyHtml = $true
	
	$SMTPclient.Send($Message)
}


###################################################
#		MAIN
###################################################

# Construct CSV file full path
$CSVExport = Join-Path $ExportPath $CSVFilename

# Find repadmin binary if not defined
if ( !($RepadminExec) )
{
	foreach ( $path in ($env:path).split(";") ) 
	{ 
		if (Test-Path $path\repadmin.exe)
		{
			$RepadminExec = "$path\repadmin.exe"
		}
	}
}

if ( $RepadminExec )
{
	# Command to generate replications status
	$StrCmd = "$RepadminExec /showrepl * /csv"
	
	# Generate CSV output
	Invoke-Expression $StrCmd | Out-File $CSVExport
	
	if ( Test-Path $CSVExport )
	{
		$ReplicationsState = Import-Csv -Path $CSVExport -Header $Header -Delimiter ","

		$ServersInError = @()
		$ErrorContent = @()

		# Filtering error messages
		$ReplicationsState | Where-Object { $_."showrepl_COLUMNS" -match "showrepl_ERROR" } | %{
			if ($_.showrepl_COLUMNS -match "showrepl_ERROR") {
				$ServersInError += $_."Destination DSA"
				$ErrorContent += $_
			} 
		}

		# Format errors content
		if ( $ErrorContent )
		{
			$MailObject="AD REPLICATIONS STATUS - $(Get-Date -Format yy/MM/dd) - error(s) found"

			$MailBody += $ErrorContent | Select-Object -Property "Destination DSA",@{ Name="Error Message"; Expression={ $_."Naming Context" } } | ConvertTo-Html -As table -Fragment -PreContent "<p>AD Replications status with errors</p>" -PostContent "<br><br>" | Out-String
			$MailBody = $MailBody.Replace("<th>","<th class=""error"">")
			$MailBody = $MailBody.Replace("<td>","<td class=""error"">")
			$MailBody = $MailBody.Replace("<table>","<table class=""error"">")
		}
		
		# Filtering warning messages
		$WarningContent = $ReplicationsState | Where-Object {  @("showrepl_ERROR","showrepl_COLUMNS") -notcontains $_."showrepl_COLUMNS" -and $_."Last Failure Status" -ne 0 }  
		
		# Format warning content
		if ( $WarningContent )
		{
			if (!$MailObject)
			{
				$MailObject="AD REPLICATIONS STATUS - $(Get-Date -Format yy/MM/dd) - warning(s) found"
			}
			
			$MailBody += $WarningContent | Select-Object -ExcludeProperty "showrepl_COLUMNS","Transport Type" -Property * | ConvertTo-Html -As table -Fragment -PreContent "<p>AD Replications status with warning</p>" -PostContent "<br><br>" | Out-String
			$MailBody = $MailBody.Replace("<th>","<th class=""warning"">")
			$MailBody = $MailBody.Replace("<td>","<td class=""warning"">")
			$MailBody = $MailBody.Replace("<table>","<table class=""warning"">")
		}

		# Filtering success messages (uncomment the line if you want to see them in the email report)
		$SuccessContent = $ReplicationsState | Where-Object { @("showrepl_ERROR","showrepl_COLUMNS") -notcontains $_."showrepl_COLUMNS" -and $_."Last Failure Status" -eq 0} 

		# Format success content
		if ( $SuccessContent )
		{
			if (!$MailObject)
			{
				$MailObject="AD REPLICATIONS STATUS - $(Get-Date -Format yy/MM/dd) - OK"
			}
			
			$MailBody += $SuccessContent | select-object -ExcludeProperty "showrepl_COLUMNS","Transport Type" -Property * | ConvertTo-Html -As table -body "<p>AD Replications status with success</p>" -PreContent $style | Out-String
			$MailBody = $MailBody.Replace("<th>","<th class=""success"">")
			$MailBody = $MailBody.Replace("<td>","<td class=""success"">")
			$MailBody = $MailBody.Replace("<table>","<table class=""success"">")
		}

		# Generate HTML file (if wanted)
		$HTMLExport = Join-Path $ExportPath $HTMLFilename
		ConvertTo-Html -Head $style -body $MailBody -Title $MailObject | Out-File $HTMLExport
		
		if ( $MailBody )
		{
			EmailNotification $Sender $Recipient $Server $MailObject (ConvertTo-Html -Head $style -Body $MailBody -Title $MailObject | Out-String)
		}
	}
}
