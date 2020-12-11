Import-Module c:\Module\CredentialManager.psm1
$creds = Get-StoredCredential MailSend -StorePath c:\Credentials\
$InputFilepath="\\belw-rs3617-2\install\Work Utils\hr_users\" 
$InputFile = $InputFilepath+(Get-Date  -Format "yyyyMMdd")+'personalu.txt' # новый путь и каждые день новый файл
$Logfilepath="\\belw-rs3617-2\install\Work Utils\hr_users\Log\"
$BWADPath = "OU=СООО Белвест,DC=belwest,DC=corp","OU=ТУП Белвест Ритейл,DC=belwest,DC=corp" # OU для поиска
$BWADFiredPath = "OU=Уволенные сотрудники,DC=belwest,DC=corp"

$all_pass_file="\\172.29.1.200\ostp\Группа технического сопровождения\УЧЕТКИ\облако_glpi.txt"
$tempfile=".\tempfile.txt"
$tempcsv=".\temp.csv"
Get-Content $InputFile -Encoding UTF8 | Out-File $tempcsv -Encoding unicode

$header = "HZ","HZ1","HZ2","CodeOrg","Org","DateIn","DateOut","TubNum","FIO","Familia","Name","Otch","IDProfFull","ProfFull","IDDepatFull","DepatFull","SapLogin","WSTel","WMTel","PMTel","PSTel","WMail","PMail","WSkype","PSkype","Prof","Depat","Birthday","EmploymentDate"
$csv = Import-CSV $tempcsv -header $header -Delimiter ';'

# файлы логов
$DisabledUsersFile=$Logfilepath+'Отключенные учетные записи '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$PassedUsersFile=$Logfilepath+'Не отключенные учетные записи '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$FullLog=$Logfilepath+(Get-Date  -Format "yyyyMMdd")+' Отключение учетных записей ImportLog.txt'

Out-File $FullLog -InputObject ("+++++++++++++++++++++++++++++++++++++++++++++++++++++++") -Append -Encoding "Default"
Out-File $FullLog -InputObject ("Обработка на дату "+(Get-Date -Format "dd.MM.yyyy")) -Append -Encoding "Default"
Out-File $FullLog -InputObject ("Используется файл "+$InputFile) -Append -Encoding "Default"

$magazines = @()
$magazine_passwords = @()
$filials = @()
$csv | Select-Object CodeOrg, Org, FIO, ProfFull, DepatFull, Tubnum, Prof, Depat, Birthday, EmploymentDate, Familia, Name, Otch, DateIn | ForEach-Object{
    $FIO = $_.FIO
    $TubNum = $_.TubNum
    Out-File $FullLog -InputObject ("Обработка "+$_.FIO+"...") -Append -Encoding "Default"
        $user = $BWADPath | foreach {get-ADUser -Searchbase $_ -Filter {(displayname -eq $FIO) -and (enabled -eq "True")} -Properties EmployeeID, Department, Title, Company, Description, fullDeparment, sAMAccountName, birthDay, comment, distinguishedName, info, employeeType}
	#Если пользователь найден
        If ($null -ne $user) 
       	{
            	#Проверка табельного
           	 if (($user.EmployeeID -eq $_.TubNum) -or ($user.EmployeeID -eq $null))
            	{
	    		#$user | set-ADUser -Enabled $false
	    		Out-File $DisabledUsersFile -InputObject ($_.FIO+", "+$user.description+", "+$user.department+", №"+$_.TubNum) -Append -Encoding "Default"
	    		$user | set-ADUser -Replace @{description="Дата увольнения: "+$_.DateIn}
	    		#Move-ADObject -Identity $user.distinguishedName -TargetPath $BWADFiredPath
			if (($user.info -gt '') -and ($user.employeeType -match "Магазин") -and (-not ($user.info -in $magazines)))
			{
				$password = -join (1..7 | % { [char[]]'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789' | Get-Random }) + -join (1..1 | % { [char[]]'0123456789' | Get-Random })
				$user_magaz = get-ADUser -Searchbase "OU=Магазины,OU=Филиалы,OU=ТУП Белвест Ритейл,DC=belwest,DC=corp" -Filter {(displayname -eq $user.info) -and (enabled -eq "True")} -Properties physicalDeliveryOfficeName
				$user_magaz | Set-ADAccountPassword -NewPassword (ConvertTo-SecureString -String $password -AsPlainText -Force)
				$user_magaz | set-ADUser -Replace @{physicalDeliveryOfficeName=$password}
				$magazines += $user.info
				$magazine_passwords += ($user.info+"   "+$password)
				$filials += $user.info.substring(0,2)
			}
	    	}
            	else
            	{
            		Out-File $PassedUsersFile -InputObject ("Пропуск: Не совпадают табельные для "+$_.FIO + " (в AD - " + $user.EmployeeID + ", в выгрузке - " + $TubNum) -Append -Encoding "Default"
		}
	}
        else
	{
	}
} | Out-Null

$filials = $filials | select -uniq

foreach ($filial in $filials)
{
	foreach ($magazine_password in $magazine_passwords)
	{
		if ($magazine_password.StartsWith($filial))
		{
			Out-File $tempfile -InputObject ($magazine_password) -Append -Encoding "Default"
		}
	}
	if ($filial -eq "sh")
	{
		$mail = "ivan.pilugaev@belwest.com"
		$subject = 'Магазины РБ - изменены пароли для облака и GLPI'
	}
	else
	{
		$acc = $filial+"1direktor"
		$user_filial = get-ADUser -Searchbase "OU=Филиалы,OU=ТУП Белвест Ритейл,DC=belwest,DC=corp" -Filter {sAMAccountName -eq $acc} -Properties mail
		$mail = $user_filial.mail
		$subject = 'Филиал '+$filial+'100 - изменены пароли отделений для облака и GLPI'
	}

	#$mail = "ok@belwest.com"
	$copy = "ok@belwest.com"
	$text = Get-Content $tempfile | ConvertTo-HTML -Property @{Label=$subject;Expression={$_}}
	Send-MailMessage -From "ps_log@belwest.com" -To $mail, $copy -Subject "$subject" -Body "$text" -BodyAsHtml -Credential $creds -SmtpServer "mail.belwest.com" -UseSsl -Encoding UTF8
	Remove-Item $tempfile
}

If (Test-Path $DisabledUsersFile)
{
	$text = Get-Content $DisabledUsersFile | ConvertTo-HTML -Property @{Label='Отключенные учетные записи';Expression={$_}}
	Send-MailMessage -From "ps_log@belwest.com" -To "ok@belwest.com" -Subject "Отключенные учетные записи AD" -Body "$text" -BodyAsHtml -Credential $creds -SmtpServer "mail.belwest.com" -Attachments $DisabledUsersFile -UseSsl -Encoding UTF8
}

Get-ADUser  -SearchBase "OU=Магазины,OU=Филиалы,OU=ТУП Белвест Ритейл,DC=belwest,DC=corp" -Properties displayName, physicalDeliveryOfficeName -Filter 'enabled -eq $true' | Sort displayName | FT @{Label="Login"; Expression={$_.displayName}}, @{Label="Password"; Expression={$_.physicalDeliveryOfficeName}} -Autosize > $all_pass_file
