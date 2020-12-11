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
#$Logfilepath=".\Log\"
#$InputFile = ".\testrf.csv"
#$InputFile = $InputFilepath+'tempfile.txt'
$BWAD='DC=belwest,DC=corp'
$BWADPath = 'OU=РБ,OU=Сотрудники магазинов,OU=Филиалы,OU=ТУП Белвест Ритейл,DC=belwest,DC=corp'
$BWADPathNewUsers = 'OU=РБ,OU=Сотрудники магазинов,OU=Филиалы,OU=ТУП Белвест Ритейл,DC=belwest,DC=corp'
$BWADPathGroup = 'OU=ТУП Белвест Ритейл,DC=belwest,DC=corp'

$tempcsv=".\temp.csv"
Get-Content $InputFile -Encoding UTF8 | Out-File $tempcsv -Encoding unicode

$header = "HZ","HZ1","HZ2","CodeOrg","Org","DateIn","DateOut","TubNum","FIO","Familia","Name","Otch","IDProfFull","ProfFull","IDDepatFull","DepatFull","SapLogin","WSTel","WMTel","PMTel","PSTel","WMail","PMail","WSkype","PSkype","Prof","Depat","Birthday","EmploymentDate"
$csv = Import-CSV $tempcsv -header $header -Delimiter ';'

# файлы логов
$BadOutFile=$Logfilepath+'Магазины РБ Не обработаные '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$NewGroupFile=$Logfilepath+'Магазины РБ Новые группы '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$NewUserFile=$Logfilepath+'Магазины РБ Изменения у пользователей '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$UserErrorFile=$Logfilepath+'Магазины РБ ошибки обработки '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$NewAccountFile=$Logfilepath+'Новые учетные записи '+(Get-Date  -Format "yyyyMMdd")+'.txt'

$FullLog=$Logfilepath+(Get-Date  -Format "yyyyMMdd")+' Магазины РБ ImportLog.txt'

Out-File $FullLog -InputObject ("+++++++++++++++++++++++++++++++++++++++++++++++++++++++") -Append -Encoding "Default"
Out-File $FullLog -InputObject ("Обработка на дату "+(Get-Date -Format "dd.MM.yyyy")) -Append -Encoding "Default"
Out-File $FullLog -InputObject ("Используется файл "+$InputFile) -Append -Encoding "Default"

Out-File $BadOutFile -InputObject ("+++++++++++++++++++++++++++++++++++++++++++++++++++++++") -Append -Encoding "Default"
Out-File $BadOutFile -InputObject ("Обработка на дату "+(Get-Date -Format "dd.MM.yyyy")) -Append -Encoding "Default"

Out-File $NewAccountFile -InputObject ("--------------------Магазины РБ--------------------") -Append -Encoding "Default"

$csv | Select-Object CodeOrg, Org, FIO, ProfFull, DepatFull, Tubnum, Prof, Depat, Birthday, EmploymentDate, Familia, Name, Otch, HZ1 | ForEach-Object{
	Clear-Variable -Name empDate
	$empType = "РБ Магазин"
	$CodeOrg = $_.CodeOrg
	$FIO = $_.FIO
	$TubNum = $_.TubNum
	$exTubNum = $_.TubNum.Substring(0,3)
	$birthday = $_.Birthday
	if ($_.Depat -ne '')
	{
		if ((($_.Depat -match "\d+") | %{$matches[0]}).length -eq 1) {
			$magazin = "shop0"+(($_.Depat -match "\d+") | %{$matches[0]})
		} else {
			$magazin = "shop"+(($_.Depat -match "\d+") | %{$matches[0]})
		}
	}
	else
	{
		$magazin = ''
	}
	# приведение даты приема на работу в UNIXtime и преобразование его в INT
	$currentDate = [Int64](Get-Date (Get-Date -Format dd.MM.yyyy) -UFormat %s)
	$empDate = [Int64](Get-Date $_.EmploymentDate -UFormat %s)
	if($empDate -eq $null) {Out-File $UserErrorFile -InputObject ($_.FIO + " - некорректная дата приема на работу: " + $_.EmploymentDate) -Append -Encoding "Default"}
	# проверка сотрудника на соответствие критериям:
	# - 1000 - код организации Белвест
	# - фильтрация по наличию в названии отдела слова Магазин
	# - исключение табельных с первыми цифрами 090, 330, 320
	If (($CodeOrg -eq "1000") -and ($_.Depat -match "магазин") -and !(($exTubNum -eq "090") -or ($exTubNum -eq "330") -or ($exTubNum -eq "320")))
        {
		Out-File $FullLog -InputObject ("Обработка "+$_.FIO+"...") -Append -Encoding "Default"
		$user = get-ADUser -Searchbase $BWADPath -Filter {EmployeeID -eq $TubNum} -Properties EmployeeID, Department, Title, Company, Description, fullDeparment, sAMAccountName, birthDay, comment, employeeType
		#Если пользователь найден
		If ($null -ne $user )
		{
			#Проверка отдела
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
            	Set-ADUser $user -Replace @{birthday=$_.BirthDay}
            	Out-File $FullLog -InputObject ("Дата рождения обновлена") -Append -Encoding "Default"
            	Out-File $NewUserFile -InputObject ($FIO+" изменена дата рождения") -Append -Encoding "Default"
            }
			#Тип работника - возможно уже ненужное поле
            if ($user.employeeType -eq $empType)
            {
				Out-File $FullLog -InputObject ("Тип работника верный") -Append -Encoding "Default"
			}
            else
            {
            	Set-ADUser $user -Replace @{employeeType=$empType}
            	Out-File $FullLog -InputObject ("Тип работника обновлен") -Append -Encoding "Default"
            	Out-File $NewUserFile -InputObject ($FIO+" изменен тип работника") -Append -Encoding "Default"
            }
            #дата устройства на работу - обновляется дата только за последние 30 дней, старше - удаляется
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
				}
            }
		#Отделение + Фамилия
		if ($user.info -eq $magazin)
		{
			Out-File $FullLog -InputObject ("Отделение верное") -Append -Encoding "Default"
		}
		else
		{
			Set-ADUser $user -Replace @{info=$magazin}
			Set-ADUser $user -Replace @{sn=$magazin+" "+$_.Familia}
			Out-File $FullLog -InputObject ("Отдлеление и фамилия обновлены") -Append -Encoding "Default"
			Out-File $NewUserFile -InputObject ($FIO+" изменены отдлеление и фамилия") -Append -Encoding "Default"
		}
            $TempDepat = $_.Depat.Replace('«','')
            $TempDepat = $TempDepat.Replace('»','')
			# Обработка групп
            $ADGroup = 'Ритейл '+$TempDepat
            $ADGroupFull = 'Ритейл '+$_.DepatFull
            $GroupExist = Get-ADGroup -SearchBase $BWADPathGroup -LDAPFilter “(name=$ADGroup)” -Properties Description
            Out-File $FullLog -InputObject ("Обработка группы пользователя") -Append -Encoding "Default"
	    	$UserObject = Get-ADUser -Filter {EmployeeID -eq $TubNum } -Property "MemberOf"
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
				If ($UserObject.MemberOf.Value -eq $GroupExist )
				{
					Out-File $FullLog -InputObject ("Группа существует и пользователь уже в ней ") -Append -Encoding "Default"
                }
                else
                {
                	Add-ADGroupMember -Identity $ADGroup -Members @user
                	Out-File $FullLog -InputObject ($FIO+" новый член группы "+$ADGroup) -Append -Encoding "Default"
		        }
            }
            else
            {
            	Out-File $newgroupfile -InputObject ($ADGroup +" добавлена в AD ") -Append -Encoding "Default"
            	Out-File $FullLog -InputObject ($ADGroup +" добавлена в AD ") -Append -Encoding "Default"
            	New-ADGroup $ADGroup -Path $BWADPathGroup -GroupScope Global -Description $ADGroupFull -PassThru -Verbose
            	Add-ADGroupMember -Identity $ADGroup -Members @user
            	Out-File $FullLog -InputObject ($FIO+" новый член группы "+$ADGroup) -Append -Encoding "Default"
            }
            Clear-Variable -Name user
            Clear-Variable -Name CodeOrg
            Clear-Variable -Name ADGroup
			Clear-Variable -Name empDate
            Out-File $FullLog -InputObject (" --------------- ") -Append -Encoding "Default"
        }
		else
        {
			# создание новой учетки
			# ищем в AD пользователя с таким же ФИО и датой рождения, если их нет - создаем учетку
        	$checkUser = get-ADUser -Searchbase $BWAD -Filter {(displayname -eq $FIO) -and (birthDay -eq $birthday)} -Properties EmployeeID, Department, Title, Company, Description, fullDeparment, comment, birthDay, employeeType
			if (!(($exTubNum -eq "090") -or ($exTubNum -eq "330") -or ($exTubNum -eq "320")) -and $checkUser -eq $null)
			{
				# транслитерация фамилии и инициалов
				$transLastName=Translit($_.Familia.Replace('-',''))
				$transFirstName=Translit($_.Name.Chars(0))
				$transInitials=Translit($_.Otch.Chars(0))
				$samAccountName = $transLastName + $transFirstName + $transInitials
				# проверка полученного логина на одинаковые в AD и его повторная генерация, при этом вместо инициалов - по 2 буквы имени и отчества
				$checkSamAccountName = get-ADUser -Searchbase $BWAD -Filter {SamAccountName -eq $samAccountName} -Properties EmployeeID, Department, Title, Company, Description, fullDeparment, comment, birthDay, employeeType
				if ($checkSamAccountName -ne $null)
				{
					$transFirstName=Translit($_.Name.Substring(0,2))
					$transInitials=Translit($_.Otch.Substring(0,2))
					$samAccountName = $transLastName + $transFirstName + $transInitials
				}
				# если логин больше 20 символов - удаляем лишние
				$samAccountName=$samAccountName.Remove(20)
				# формирование названия учетки и его проверка на дубли в AD; если есть - повторная генерация
				$uname = $_.Familia + " " + $_.Name + " " + $_.Otch.Chars(0) + "."
				$checkAccountName = get-ADUser -Searchbase $BWAD -Filter {name -eq $uname} -Properties EmployeeID, Department, Title, Company, Description, fullDeparment, comment, birthDay, employeeType
				if ($checkAccountName -ne $null)
				{
					$uname = $_.Familia + " " + $_.Name + " " + $_.Otch.Substring(0,2) + "."
				}
				# добавление для магазинов РБ в поле фамилии приставки shopXX - необходимо для GLPI
				if ($_.Depat -ne '')
				{
					$Familia =$magazin+" "+$_.Familia
				}
				else
				{
					$Familia = $_.Familia
				}
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
					-OtherAttributes @{'EmployeeID'=$_.TubNum;'birthDay'=$_.BirthDay;'physicalDeliveryOfficeName'=$newPass} `
					-Enabled $true
				# проверка результата создания учетки и добавление ее в необходимые группы 
				$NewUser = Get-ADUser $samAccountName
				if ($NewUser -eq $null)
				{
					Out-File $NewAccountFile -InputObject ("ОШИБКА СОЗДАНИЯ! - "+$_.FIO+", №"+$_.TubNum+", "+$samAccountName) -Append -Encoding "Default"
				}
				Add-ADGroupMember -Identity ("Shop By") -Members $NewUser
				Out-File $NewAccountFile -InputObject ($_.FIO+", "+$_.ProfFull+", "+$_.Depat+", №"+$_.TubNum+", "+$samAccountName) -Append -Encoding "Default"
			}
			else
			{
				Out-File $badoutfile -InputObject ($_.FIO+" с №"+$_.TubNum+" не найден в AD в OU Ритейл Филиалы РБ") -Append -Encoding "Default" 
				Out-File $FullLog -InputObject ($_.FIO+" с №"+$_.TubNum+" не найден в AD в OU Ритейл Филиалы РБ") -Append -Encoding "Default"
			}
			Clear-Variable -Name checkUser
			Clear-Variable -Name CodeOrg
			Clear-Variable -Name empDate
		}
	}
	else
	{
		Out-File $FullLog -InputObject ($_.FIO+" с №"+$_.TubNum+" не является сотрудником сети магазинов РБ"+$ADGroup) -Append -Encoding "Default"
	}
} | Out-Null
