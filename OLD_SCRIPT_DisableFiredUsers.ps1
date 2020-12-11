Import-Module c:\Module\CredentialManager.psm1
$creds = Get-StoredCredential MailSend -StorePath c:\Credentials\
$InputFilepath="\\belw-rs3617-2\install\Work Utils\hr_users\" 
$InputFile = $InputFilepath+(Get-Date  -Format "yyyyMMdd")+'personalu.txt' # новый путь и каждые день новый файл
$Logfilepath="\\belw-rs3617-2\install\Work Utils\hr_users\Log\"
$BWADPath = "OU=СООО Белвест,DC=belwest,DC=corp","OU=ТУП Белвест Ритейл,DC=belwest,DC=corp" # OU для поиска
$BWADFiredPath = "OU=Уволенные сотрудники,DC=belwest,DC=corp"
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

$csv | Select-Object CodeOrg, Org, FIO, ProfFull, DepatFull, Tubnum, Prof, Depat, Birthday, EmploymentDate, Familia, Name, Otch, DateIn | ForEach-Object{
    $FIO = $_.FIO
    $TubNum = $_.TubNum
    Out-File $FullLog -InputObject ("Обработка "+$_.FIO+"...") -Append -Encoding "Default"
        $user = $BWADPath | foreach {get-ADUser -Searchbase $_ -Filter {(displayname -eq $FIO) -and (enabled -eq "True")} -Properties EmployeeID, Department, Title, Company, Description, fullDeparment, sAMAccountName, birthDay, comment, distinguishedName, info}
	#Если пользователь найден
        If ($null -ne $user) 
       	{
            	#Проверка табельного
           	 if (($user.EmployeeID -eq $_.TubNum) -or ($user.EmployeeID -eq $null))
            	{
	    		$user | set-ADUser -Enabled $false
	    		Out-File $DisabledUsersFile -InputObject ($_.FIO+", "+$user.description+", "+$user.department+", №"+$_.TubNum) -Append -Encoding "Default"
	    		$user | set-ADUser -Replace @{description="Дата увольнения: "+$_.DateIn}
	    		Move-ADObject -Identity $user.distinguishedName -TargetPath $BWADFiredPath
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

If (Test-Path $DisabledUsersFile)
{
	$text = Get-Content $DisabledUsersFile | ConvertTo-HTML -Property @{Label='Отключенные учетные записи';Expression={$_}}
	Send-MailMessage -From "ps_log@belwest.com" -To "it@belwest.com" -Subject "Отключенные учетные записи AD" -Body "$text" -BodyAsHtml -Credential $creds -SmtpServer "mail.belwest.com" -Attachments $DisabledUsersFile -UseSsl -Encoding UTF8
}
