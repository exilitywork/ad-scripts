#сама функция транслитерации
function global:Translit {
	param([string]$inString)
	$Translit = @{
		[char]'а' = "a"
		[char]'А' = "A"
		[char]'б' = "b"
		[char]'Б' = "B"
		[char]'в' = "v"
		[char]'В' = "V"
		[char]'г' = "g"
		[char]'Г' = "G"
		[char]'д' = "d"
		[char]'Д' = "D"
		[char]'е' = "e"
		[char]'Е' = "E"
		[char]'ё' = "yo"
		[char]'Ё' = "Yo"
		[char]'ж' = "zh"
		[char]'Ж' = "Zh"
		[char]'з' = "z"
		[char]'З' = "Z"
		[char]'и' = "i"
		[char]'И' = "I"
		[char]'й' = "j"
		[char]'Й' = "J"
		[char]'к' = "k"
		[char]'К' = "K"
		[char]'л' = "l"
		[char]'Л' = "L"
		[char]'м' = "m"
		[char]'М' = "M"
		[char]'н' = "n"
		[char]'Н' = "N"
		[char]'о' = "o"
		[char]'О' = "O"
		[char]'п' = "p"
		[char]'П' = "P"
		[char]'р' = "r"
		[char]'Р' = "R"
		[char]'с' = "s"
		[char]'С' = "S"
		[char]'т' = "t"
		[char]'Т' = "T"
		[char]'у' = "u"
		[char]'У' = "U"
		[char]'ф' = "f"
		[char]'Ф' = "F"
		[char]'х' = "h"
		[char]'Х' = "H"
		[char]'ц' = "c"
		[char]'Ц' = "C"
		[char]'ч' = "ch"
		[char]'Ч' = "Ch"
		[char]'ш' = "sh"
		[char]'Ш' = "Sh"
		[char]'щ' = "sch"
		[char]'Щ' = "Sch"
		[char]'ъ' = ""
		[char]'Ъ' = ""
		[char]'ы' = "y"
		[char]'Ы' = "Y"
		[char]'ь' = ""
		[char]'Ь' = ""
		[char]'э' = "e"
		[char]'Э' = "E"
		[char]'ю' = "yu"
		[char]'Ю' = "Yu"
		[char]'я' = "ya"
		[char]'Я' = "Ya"
	}
	$outCHR=""
	foreach ($CHR in $inCHR = $inString.ToCharArray())
	{
		if ($Translit[$CHR] -cne $Null )
			{$outCHR += $Translit[$CHR]}
		else
			{$outCHR += $CHR}
	}
	Write-Output $outCHR
}

# Генерация пароля и проверка на его соответствие политике
Function global:Generate-Complex-Domain-Password ([Parameter(Mandatory=$true)][int]$PassLenght)
{
	Add-Type -AssemblyName System.Web
	$requirementsPassed = $false
	do {
		$newPassword=[System.Web.Security.Membership]::GeneratePassword($PassLenght,1)
		If ( ($newPassword -cmatch "[A-Z\p{Lu}\s]") `
			-and ($newPassword -cmatch "[a-z\p{Ll}\s]") `
			-and ($newPassword -match "[\d]") `
			-and ($newPassword -match "[^\w]") `
			-and ($newPassword -match "[^_]") `
			-and (([char[]]$newPassword -match "[^\w]").count -eq 1)
		)
		{
			$requirementsPassed=$True
		}
	} While ($requirementsPassed -eq $false)
	return $newPassword
}

# Обработка пользователей
$InputFilepath="\\belw-rs3617-2\install\Work Utils\hr_users\" 
$InputFile = $InputFilepath+(Get-Date  -Format "yyyyMMdd")+'personal.txt' # новый путь и каждые день новый файл
$Logfilepath="\\belw-rs3617-2\install\Work Utils\hr_users\Log\"
$GedeminPath="\\belw-appsvr\Report\"

$BWAD='DC=belwest,DC=corp'
$BWADPath = 'OU=Офис,OU=ТУП Белвест Ритейл,DC=belwest,DC=corp'
$BWADPathGroup = 'OU=ТУП Белвест Ритейл,DC=belwest,DC=corp'
$BWADPathNewUsers = 'OU=Новые учетные записи,OU=Офис,OU=ТУП Белвест Ритейл,DC=belwest,DC=corp'
#clear-content -path $Logfilepath* -filter *.txt -force  #не надо чистить пусть будет история
#$Logfilepath=".\Log\"
#$InputFile = ".\testrf.csv"
#$InputFile = $InputFilepath+'tempfile.txt'
#$InputFile="i:\Windows\Powershell\Корректировка персонала\testv131.csv" 
#$InputFile="i:\Work Utils\hr_users\20200226personal.txt" 
$tempcsv=".\temp.csv"
Get-Content $InputFile -Encoding UTF8 | Out-File $tempcsv -Encoding unicode

$header = "HZ","HZ1","HZ2","CodeOrg","Org","DateIn","DateOut","TubNum","FIO","Familia","Name","Otch","IDProfFull","ProfFull","IDDepatFull","DepatFull","SapLogin","WSTel","WMTel","PMTel","PSTel","WMail","PMail","WSkype","PSkype","Prof","Depat","Birthday","EmploymentDate"
$csv = Import-CSV $tempcsv -header $header -Delimiter ';'

#$goodoutfile=$Logfilepath+$_.Depat+' +.txt' #пока не надо
# файлы логов
$BadOutFile=$Logfilepath+'Ритейл Офис Не обработаные '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$NewGroupFile=$Logfilepath+'Ритейл Офис Новые группы '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$NewUserFile=$Logfilepath+'Ритейл Офис Изменения у пользователей '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$FullLog=$Logfilepath+'Ритейл Офис ImportLog '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$UserErrorFile=$Logfilepath+'Ритейл Офис ошибки обработки '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$NewAccountFile=$Logfilepath+'Новые учетные записи '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$gedemin=$GedeminPath+'Сотрудники.txt'

Out-File $FullLog -InputObject ("+++++++++++++++++++++++++++++++++++++++++++++++++++++++") -Append -Encoding "Default"
Out-File $FullLog -InputObject ("Обработка на дату "+(Get-Date -Format "dd.MM.yyyy")) -Append -Encoding "Default"
Out-File $FullLog -InputObject ("Используется файл "+$InputFile) -Append -Encoding "Default"

Out-File $BadOutFile -InputObject ("+++++++++++++++++++++++++++++++++++++++++++++++++++++++") -Append -Encoding "Default"
Out-File $BadOutFile -InputObject ("Обработка на дату "+(Get-Date -Format "dd.MM.yyyy")) -Append -Encoding "Default"

Out-File $NewAccountFile -InputObject ("--------------------Ритейл Офис--------------------") -Append -Encoding "Default"

$csv | Select-Object CodeOrg, Org, FIO, ProfFull, DepatFull, Tubnum, Prof, Depat, Birthday, EmploymentDate, Familia, Name, Otch | ForEach-Object{
	$empDate = $null
	$CodeOrg = $_.CodeOrg
	$FIO = $_.FIO
	$currentDate = [Int64](Get-Date (Get-Date -Format dd.MM.yyyy) -UFormat %s)
	$empDate = [Int64](Get-Date $_.EmploymentDate -UFormat %s)
	$TubNum = $_.TubNum
	$birthday = $_.Birthday
	if($empDate -eq $null) {Out-File $UserErrorFile -InputObject ($_.FIO + " - некорректная дата приема на работу: " + $_.EmploymentDate) -Append -Encoding "Default"}
	If ($CodeOrg -ne '1000')
	{
		Out-File $FullLog -InputObject ("Обработка "+$_.FIO+'...') -Append -Encoding "Default"
		$user = get-ADUser -Searchbase $BWADPath -Filter {EmployeeID -eq $TubNum} -Properties EmployeeID, Department, Title, Company, Description, fullDeparment, comment, birthDay
		#write-host $user
		#write-host $_.TubNum.Substring(0,3)
		#Если пользователь найден
		If ($null -ne $user ) 
		{
			#Экспорт сотрудников для столовой
			Out-File $gedemin -InputObject ($TubNum.TrimStart("0")+";"+$FIO+";"+$_.DepatFull) -Append -Encoding "Default"
			#Проверка отдела
			#Write-Host $user.Department
			#Write-Host $_.Depat
			if (($user.Department -eq $_.Depat) -or ($_.Depat -eq ''))
			{
				Out-File $FullLog -InputObject ("Отдел верный") -Append -Encoding "Default"
			}
			else
			{
				$user | set-ADUser -Department $_.Depat
				Out-File $FullLog -InputObject ("Отдел обновлен") -Append -Encoding "Default"
				Out-File $NewUserFile -InputObject ($FIO+" изменен отдел") -Append -Encoding "Default"
			}
			#Добавочное поле FullDepartment - полное название отдела
			#Write-Host $user.fullDeparment
			#Write-Host $_.DepatFull
			#$fd = $_.Depatfull.ToString()
			#Write-Host $fd
			if (($user.fullDeparment -eq $_.DepatFull) -or ($_.DepatFull -eq ''))
			{
				Out-File $FullLog -InputObject ("Полный отдел верный") -Append -Encoding "Default"
			}
			else
			{
				$user | set-ADUser -Replace @{fullDeparment=$_.DepatFull}
				Out-File $FullLog -InputObject ("Полный отдел обновлен") -Append -Encoding "Default"
				Out-File $NewUserFile -InputObject ($FIO+" изменен полный отдел") -Append -Encoding "Default"
			}
			#Проверка должности
			#Write-Host $_.Prof
			#Write-Host $user.Title
			if (($user.Title -eq $_.Prof) -or ($_.Prof -eq ''))
			{
				Out-File $FullLog -InputObject ("Должность верная") -Append -Encoding "Default"
			}
			else
			{
				$user | set-ADUser -Title $_.Prof
				Out-File $FullLog -InputObject ("Должность обновлена") -Append -Encoding "Default"
				Out-File $NewUserFile -InputObject ($FIO+" изменена должность") -Append -Encoding "Default"
			}
			#Проверка Организации
			if ($user.Company -eq $_.Org)
			{
				Out-File $FullLog -InputObject ("Организация верная") -Append -Encoding "Default"
			}
			else
			{
				$user | set-ADUser -Company $_.Org
				Out-File $FullLog -InputObject ("Организация обновлена") -Append -Encoding "Default"
				Out-File $NewUserFile -InputObject ($FIO+" изменена организация") -Append -Encoding "Default"
			}
            #Описание
			if (($user.Description -eq $_.ProfFull) -or ($_.ProfFull -eq ''))
            {
				Out-File $FullLog -InputObject ("Описание верное") -Append -Encoding "Default"
			}
            else
            {
				$user | set-ADUser -Description $_.ProfFull
				Out-File $FullLog -InputObject ("Описание обновлено") -Append -Encoding "Default"
				Out-File $NewUserFile -InputObject ($FIO+" изменено описание") -Append -Encoding "Default"
			}
			#Дата рождения 
			if ($user.birthDay -eq $_.Birthday)
			{
				Out-File $FullLog -InputObject ("Дата рождения верная") -Append -Encoding "Default"
			}
			else
			{
				#$user | set-ADUser -birthDay $_.BirthDay
				Set-ADUser $user -Replace @{birthday=$_.BirthDay}
				Out-File $FullLog -InputObject ("Дата рождения обновлена") -Append -Encoding "Default"
				Out-File $NewUserFile -InputObject ($FIO+" изменена дата рождения") -Append -Encoding "Default"
			}
			#дата устройства на работу
			if (($empDate -gt ($currentDate - 2592000)) -and ($user.comment -eq $null))
			{
				Set-ADUser $user -Replace @{comment = $empDate}
				Out-File $FullLog -InputObject ("Дата приема на работу обновлена") -Append -Encoding "Default"
				Out-File $NewUserFile -InputObject ($FIO+" изменена дата приема на работу") -Append -Encoding "Default"
			}
            else
            {
				if (($empDate -lt ($currentDate - 2592000)) -and ($user.comment -ne $null))
				{
					Set-ADUser $user -Clear comment
					Out-File $FullLog -InputObject ("Дата приема на работу удалена") -Append -Encoding "Default"
					Out-File $NewUserFile -InputObject ($FIO+" дата приема на работу удалена") -Append -Encoding "Default"
				}
			}
			# раньше бахалось без проверок
			# $user | set-ADUser -EmployeeID $_.TubNum -Department $_.Depat -Title $_.Prof -Company $_.Org -Description $_.Prof
			# Out-File $goodoutfile -InputObject ($_.FIO+" найден и его данные обновлены ") -Append -Encoding "Default" 
			# Out-File $FullLog -InputObject ($_.FIO+" найден и его данные обновлены ") -Append -Encoding "Default"
			# Clear-Variable -Name user
			# Clear-Variable -Name CodeOrg
			$TempDepat = $_.Depat.Replace('«','')
			$TempDepat = $TempDepat.Replace('»','')
			$ADGroup = 'Ритейл '+$TempDepat
			$ADGroupFull = 'Ритейл '+$_.DepatFull
			$GroupExist = Get-ADGroup -SearchBase $BWADPathGroup -LDAPFilter “(name=$ADGroup)” -Properties Description
			Out-File $FullLog -InputObject ("Обработка группы пользователя") -Append -Encoding "Default"
			#Write-Host ---
			#Write-Host $GroupExist.Description
			#Write-Host $_.DepatFull
			if ($GroupExist -ne $null)
			{
				if ($GroupExist.Description -eq $ADGroupFull)
				{
				Out-File $FullLog -InputObject ("Описание группы "+$ADGroup+" верное") -Append -Encoding "Default"
				}
				else
				{
				$GroupExist | Set-ADGroup -Description $ADGroupFull
				Out-File $FullLog -InputObject ("Описание группы "+$ADGroup+" обновлено на "+$ADGroupFull) -Append -Encoding "Default"
				}
				$UserObject = Get-ADUser -Filter {displayname -eq $FIO } -Property "MemberOf"
				#$UserObject.MemberOf | ForEach-Object {
				#     Write-Host $UserObject.MemberOf
				If ($UserObject.MemberOf.Value -eq $GroupExist )
				#  if ((Get-ADUser $user -Properties memberof) | where {$_ -like $GroupExist})  
				{
					Out-File $FullLog -InputObject ("Группа существует и пользователь уже в ней ") -Append -Encoding "Default"
					#Write-Host 'Пользователь уже в группе' 
				}
				else
				{
					Add-ADGroupMember -Identity $ADGroup -Members @user
					Out-File $FullLog -InputObject ($FIO+" новый член группы "+$ADGroup) -Append -Encoding "Default"
					#Write-Host $FIO' новый член группы ' $ADGroup
				}
				#}
			}
			else 
			{
				Out-File $newgroupfile -InputObject ($ADGroup +" добавлена в AD ") -Append -Encoding "Default"
				Out-File $FullLog -InputObject ($ADGroup +" добавлена в AD ") -Append -Encoding "Default"
				New-ADGroup $ADGroup -Path $BWADPathGroup -GroupScope Global -Description $ADGroupFull -PassThru -Verbose
				Add-ADGroupMember -Identity $ADGroup -Members @user
				Out-File $FullLog -InputObject ($FIO+" новый член группы "+$ADGroup) -Append -Encoding "Default"
				# Write-Host 'Создали'$ADGroup
			}
			Clear-Variable -Name user
			Clear-Variable -Name CodeOrg
			Clear-Variable -Name ADGroup
			Out-File $FullLog -InputObject (" --------------- ") -Append -Encoding "Default"
		}
		else 
		{
			$checkUser = get-ADUser -Searchbase $BWAD -Filter {(displayname -eq $FIO) -and (birthDay -eq $birthday)} -Properties EmployeeID, Department, Title, Company, Description, fullDeparment, comment, birthDay, employeeType
			If (($CodeOrg.Substring(2) -eq "N0") -and ($checkUser -eq $null))
			{
				$transLastName=Translit($_.Familia.Replace('-',''))
				$transFirstName=Translit($_.Name.Chars(0))
				$transInitials=Translit($_.Otch.Chars(0))
				$samAccountName = $transLastName + $transFirstName + $transInitials
				$checkSamAccountName = get-ADUser -Searchbase $BWAD -Filter {SamAccountName -eq $samAccountName} -Properties EmployeeID, Department, Title, Company, Description, fullDeparment, comment, birthDay, employeeType
				if ($checkSamAccountName -ne $null)
				{
					$transFirstName=Translit($_.Name.Substring(0,2))
					$transInitials=Translit($_.Otch.Substring(0,2))
					$samAccountName = $transLastName + $transFirstName + $transInitials
				}
				$samAccountName=$samAccountName.Remove(20)
				$uname = $_.Familia + " " + $_.Name + " " + $_.Otch.Chars(0) + "."
				$checkAccountName = get-ADUser -Searchbase $BWAD -Filter {name -eq $uname} -Properties EmployeeID, Department, Title, Company, Description, fullDeparment, comment, birthDay, employeeType
				if ($checkAccountName -ne $null)
				{
					$uname = $_.Familia + " " + $_.Name + " " + $_.Otch.Substring(0,2) + "."
				}
				$Familia = $_.Familia
				$upn = $samAccountName + “@belwest.corp”
				$newPass = Generate-Complex-Domain-Password(8)
				#write-host $samAccountName " --- " $newPass
				New-ADUser -SamAccountName $samAccountName `
					-name $uname `
					-Path $BWADPathNewUsers `
					-DisplayName $_.FIO `
					-GivenName $_.Name `
					-Surname $Familia `
					-Initials $_.Otch.Chars(0) `
					-AccountPassword (ConvertTo-SecureString $newPass -AsPlainText -force) `
					-Department $_.Depat `
					-Title $_.Prof `
					-Description $_.ProfFull `
					-Company $_.Org `
					-UserPrincipalName $upn `
					-OtherAttributes @{'EmployeeID'=$_.TubNum;'birthDay'=$_.BirthDay;'physicalDeliveryOfficeName'=$newPass}
				$NewUser = Get-ADUser $samAccountName
				Set-ADUser $samAccountName -ChangePasswordAtLogon $True
				if ($NewUser -eq $null)
				{
					Out-File $NewAccountFile -InputObject ("ОШИБКА СОЗДАНИЯ! - "+$_.FIO+", №"+$_.TubNum+", "+$samAccountName) -Append -Encoding "Default"
				}
				Out-File $NewAccountFile -InputObject ($_.FIO+", "+$_.ProfFull+", "+$_.Depat+", №"+$_.TubNum+", "+$samAccountName) -Append -Encoding "Default"
				Add-ADGroupMember -Identity ("Ритейл ТУП Белвест Ритейл") -Members $NewUser
				#Экспорт нового сотрудника для столовой
				Out-File $gedemin -InputObject ($_.TubNum.TrimStart("0")+";"+$_.FIO+";"+$_.DepatFull) -Append -Encoding "Default"
			}
			else
			{
				Out-File $badoutfile -InputObject ($_.FIO+" с №"+$_.TubNum+" не найден в AD в OU Ритейл Офис") -Append -Encoding "Default" 
				Out-File $FullLog -InputObject ($_.FIO+" с №"+$_.TubNum+" не найден в AD в OU Ритейл Офис") -Append -Encoding "Default"
			}
        }
    }
	else
	{
		Out-File $FullLog -InputObject ($FIO+" с №"+$_.TubNum+" не работает на Ритейле") -Append -Encoding "Default"
	}
} | Out-Null




