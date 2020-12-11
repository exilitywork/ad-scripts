$BWADPath = "OU=СООО Белвест,DC=belwest,DC=corp","OU=Уволенные,OU=Сервисные аккаунты,DC=belwest,DC=corp"

$InputFilepath="\\belw-rs3617-2\install\Work Utils\hr_users\" 
$InputFile = $InputFilepath+(Get-Date  -Format "yyyyMMdd")+'personalu.txt' # новый путь и каждые день новый файл
# тестовый файл
#$InputFile="c:\Scripts\Blocktest.txt" 

$Logfilepath="\\belw-rs3617-2\install\Work Utils\hr_users\Log\"
#$BWADPath = "OU=СООО Белвест,DC=belwest,DC=corp","OU=ТУП Белвест Ритейл,DC=belwest,DC=corp" # OU для поиска
$tempcsv=".\temp.csv"

Get-Content $InputFile -Encoding UTF8 | Out-File $tempcsv -Encoding unicode
$header = "HZ","HZ1","HZ2","CodeOrg","Org","DateIn","DateOut","TubNum","FIO","Familia","Name","Otch","IDProfFull","ProfFull","IDDepatFull","DepatFull","SapLogin","WSTel","WMTel","PMTel","PSTel","WMail","PMail","WSkype","PSkype","Prof","Depat","Birthday","EmploymentDate"
$csv = Import-CSV $tempcsv -header $header -Delimiter ';'

# файлы логов
$DisabledUsersFile=$Logfilepath+'O365 Отключенные лицензии'+(Get-Date  -Format "yyyyMMdd")+'.txt'
#$PassedUsersFile=$Logfilepath+'O365 Не отключенные учетные записи '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$FullLog=$Logfilepath+'O365 Полный лог '+(Get-Date  -Format "yyyyMMdd")+'.txt'

#Out-File $DisabledUsersFile -InputObject ("+++++++++++++++++++++++++++++++++++++++++++++++++++++++") -Append -Encoding "Default"
#Out-File $DisabledUsersFile -InputObject ("Обработка на дату "+(Get-Date -Format "dd.MM.yyyy")) -Append -Encoding "Default"
#Out-File $DisabledUsersFile -InputObject ("Используется файл "+$InputFile) -Append -Encoding "Default"
Out-File $FullLog -InputObject ("+++++++++++++++++++++++++++++++++++++++++++++++++++++++") -Append -Encoding "Default"
Out-File $FullLog -InputObject ("Обработка на дату "+(Get-Date -Format "dd.MM.yyyy")) -Append -Encoding "Default"
Out-File $FullLog -InputObject ("Используется файл "+$InputFile) -Append -Encoding "Default"

#подключение к сервисам мс
Import-Module c:\Module\CredentialManager.psm1
#если нет сохраненных кредов то надо сделать этой командой
#Get-StoredCredential o365
$creds = Get-StoredCredential O365 -StorePath c:\Credentials\
Connect-AzureAD -Credential $creds 
Connect-MsolService -Credential $creds

$csv | Select-Object CodeOrg, FIO | ForEach-Object{
    $FIO = $_.FIO
    $CodeOrg = $_.CodeOrg
    If ($CodeOrg.Equals('1000'))
    {
        #Get-AzureADUser
        #Out-File $DisabledUsersFile -InputObject ("Обработка "+$_.FIO+"...") -Append -Encoding "Default"
        Out-File $FullLog -InputObject ("Обработка "+$_.FIO+"...") -Append -Encoding "Default"
        $ADuser = $BWADPath | foreach {get-ADUser -Searchbase $_ -Filter {(displayname -eq $FIO) -and (enabled -eq "False")}}
        if ($null -ne $ADuser )
        {
        $AZuser = Get-AzureADUser -Filter "DisplayName eq '$FIO'"
        #Если пользователь найден
        If ($null -ne $AZuser ) 
            {
            #Write-Host $user
            Get-MsolAccountSku | Select-Object AccountSkuId | ForEach-Object {
                                            $lic = $_.AccountSkuId
                                            Try {
                                                # ловим ошибку отключения лицензии
                                                Set-MsolUserLicense -UserPrincipalName $AZuser.UserPrincipalName -RemoveLicenses $_.AccountSkuId -ErrorAction Stop
                                                Out-File $DisabledUsersFile -InputObject ($FIO+" Удалена лицензия "+$lic) -Append -Encoding "Default"
                                                Out-File $FullLog -InputObject ("Удалена лицензия "+$lic) -Append -Encoding "Default"
                                                #Write-Host $_.AccountSkuId
                                                } Catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException] {
                                                Out-File $FullLog -InputObject ("Лицензия "+$lic+" не была предоставлена") -Append -Encoding "Default"
                                                #Write-Host "an error while removing lic"
                                                }
                                                                             }
            #Get-MsolUser -UserPrincipalName $user.UserPrincipalName | Format-List DisplayName,Licenses
            }
        else
        {
        Out-File $FullLog -InputObject ("Не зарегистрирован в Azure") -Append -Encoding "Default"
        }
        }
     }
     Clear-Variable -Name AZuser 
     Clear-Variable -Name ADuser
                                                    }| Out-Null
#В конец что по итогу 
Out-File $DisabledUsersFile -InputObject (Get-MsolAccountSku | where {$_.AccountSkuId -eq "belwst:O365_BUSINESS"}) -Append -Encoding "Default"
Out-File $FullLog -InputObject (Get-MsolAccountSku | where {$_.AccountSkuId -eq "belwst:O365_BUSINESS"}) -Append -Encoding "Default"

#откл от msol не надо отключать
Disconnect-AzureAD

#Отправка на it@belwest.com если есть что
If (Test-Path $DisabledUsersFile)
{
$creds = Get-StoredCredential MailSend -StorePath c:\Credentials\
Send-MailMessage -From "ps_log@belwest.com" -To "it@belwest.com" -Subject "Office365 User Block" -Body "Office365 User Block. Full lof you can see in $FullLog" -Credential $creds -SmtpServer "mail.belwest.com" -Attachments $DisabledUsersFile -UseSsl
}
