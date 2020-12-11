Import-Module c:\Module\CredentialManager.psm1

$Logfilepath="\\belw-rs3617-2\install\Work Utils\hr_users\Log\"
$NewAccountFile=$Logfilepath+"����� ������� ������ "+(Get-Date  -Format "yyyyMMdd")+".txt"

$CountNewUsers = @(Get-Content $NewAccountFile).length
If (Test-Path $NewAccountFile)
{
	$text = $null
	if($CountNewUsers -gt 4)
	{
		$text = Get-Content $NewAccountFile | ConvertTo-HTML -Property @{Label="����� ������� ������";Expression={$_}}
	}
	if($CountNewUsers -gt 20)
	{
		$text = "������� "+$CountNewUsers+" ������� �������! ������ �� ��������."
	}
	if($text -ne $null)
	{
		$creds = Get-StoredCredential MailSend -StorePath c:\Credentials\
		Send-MailMessage -From "ps_log@belwest.com" -To "it@belwest.com" -Subject "����� ������� ������ AD" -Body "$text" -BodyAsHtml -Credential $creds -SmtpServer "mail.belwest.com" -Attachments $NewAccountFile -UseSsl -Encoding UTF8
	}
}
