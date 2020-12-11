Import-Module c:\Module\CredentialManager.psm1

$Logfilepath="\\belw-rs3617-2\install\Work Utils\hr_users\Gedemin\"
$NewUsersFile=$Logfilepath+(Get-Date  -Format "yyyyMMdd")+"log_new_cards.txt"
$DisabledUsersFile=$Logfilepath+(Get-Date  -Format "yyyyMMdd")+"log_block_cards.txt"

$CountNewUsers = @(Get-Content $NewUsersFile).length
If (Test-Path $NewUsersFile)
{
	$text = $null
	if($CountNewUsers -gt 0)
	{
		$text = Get-Content $NewUsersFile | ConvertTo-HTML -Property @{Label="Новые карты сотрудников:";Expression={$_}}
	}
	if($text -ne $null)
	{
		$creds = Get-StoredCredential MailSend -StorePath c:\Credentials\
		Send-MailMessage -From "gedemin@belwest.com" -To "it@belwest.com" -Subject "Столовая: новые карты сотрудников" -Body "$text" -BodyAsHtml -Credential $creds -SmtpServer "mail.belwest.com" -Attachments $NewUsersFile -UseSsl -Encoding UTF8
	}
}

$CountDisabledUsers = @(Get-Content $DisabledUsersFile).length
If (Test-Path $DisabledUsersFile)
{
	$text = $null
	if($CountDisabledUsers -gt 2)
	{
		$text = Get-Content $DisabledUsersFile | ConvertTo-HTML -Property @{Label="Заблокированнные карты сотрудников:";Expression={$_}}
	}
	if($text -ne $null)
	{
		$creds = Get-StoredCredential MailSend -StorePath c:\Credentials\
		Send-MailMessage -From "gedemin@belwest.com" -To "it@belwest.com" -Subject "Столовая: заблокированные карты сотрудников" -Body "$text" -BodyAsHtml -Credential $creds -SmtpServer "mail.belwest.com" -Attachments $DisabledUsersFile -UseSsl -Encoding UTF8
	}
}