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
#$Logfilepath=".\Log\"
#InputFile = $InputFilepath+'tempfile.txt'

$BWAD='DC=belwest,DC=corp'
$BWADPath = 'OU=СООО Белвест,DC=belwest,DC=corp'
$BWADPathGroup = 'OU=СООО Белвест,DC=belwest,DC=corp'
$BWADPathNewUsers = 'OU=Новые учетные записи,OU=СООО Белвест,DC=belwest,DC=corp'

$tempcsv=".\temp.csv"
Get-Content $InputFile -Encoding UTF8 | Out-File $tempcsv -Encoding unicode
$header = "HZ","HZ1","HZ2","CodeOrg","Org","DateIn","DateOut","TubNum","FIO","Familia","Name","Otch","IDProfFull","ProfFull","IDDepatFull","DepatFull","SapLogin","WSTel","WMTel","PMTel","PSTel","WMail","PMail","WSkype","PSkype","Prof","Depat","Birthday","EmploymentDate","CardNum","ManagerNum"
$csv = Import-CSV $tempcsv -header $header -Delimiter ';'

# файлы логов
$BadOutFile=$Logfilepath+'Не обработаные Белвест Офис '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$NewGroupFile=$Logfilepath+'Новые группы Белвест Офис '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$NewUserFile=$Logfilepath+'Изменения у пользователей Белвест Офис '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$FullLog=$Logfilepath+(Get-Date  -Format "yyyyMMdd")+'BelwestOfficeImportLog.txt'
$UserErrorFile=$Logfilepath+'Белвест Офис ошибки обработки '+(Get-Date  -Format "yyyyMMdd")+'.txt'
$NewAccountFile=$Logfilepath+'Новые учетные записи '+(Get-Date  -Format "yyyyMMdd")+'.txt'

Out-File $FullLog -InputObject ("+++++++++++++++++++++++++++++++++++++++++++++++++++++++") -Append -Encoding "Default"
Out-File $FullLog -InputObject ("Обработка на дату "+(Get-Date -Format "dd.MM.yyyy")) -Append -Encoding "Default"
Out-File $FullLog -InputObject ("Используется файл "+$InputFile) -Append -Encoding "Default"

Out-File $BadOutFile -InputObject ("+++++++++++++++++++++++++++++++++++++++++++++++++++++++") -Append -Encoding "Default"
Out-File $BadOutFile -InputObject ("Обработка на дату "+(Get-Date -Format "dd.MM.yyyy")) -Append -Encoding "Default"

Out-File $NewAccountFile -InputObject ("--------------------Белвест Офис--------------------") -Append -Encoding "Default"

# Сопоставление SAP-идентификаторов подразделений  SAP-идентификаторам руководящих должностей
$managers=@{}
$top_managers=@{}
$csv | Select-Object CodeOrg, Org, FIO, ProfFull, DepatFull, Tubnum, Prof, Depat, Birthday, EmploymentDate, Familia, Name, Otch, ManagerNum, HZ2, IDProfFull, IDDepatFull | ForEach-Object{
	$CodeOrg = $_.CodeOrg
	$id_prof = $_.IDProfFull
	$id_dep = $_.IDDepatFull
	$TubNum = $_.TubNum
	If ($CodeOrg -eq "1000")
	{
		if ($id_prof -eq 50009336)
		{
			$managers['50000317']=$TubNum		# Цех №1 - Начальник цеха №1,5
			$top_managers['50000457']=$TubNum	# Цех №1 смена А - Начальник цеха №1,5
			$top_managers['50000459']=$TubNum	# Цех №1 смена Б - Начальник цеха №1,5
			$top_managers['50000471']=$TubNum	# Цех №5 смена А - Начальник цеха №1,5
			$top_managers['50000472']=$TubNum	# Цех №5 смена Б - Начальник цеха №1,5
		}
		elseif ($id_prof -eq 50037990)
		{
			$managers['50000457']=$TubNum		# Цех №1 смена А - Старший мастер смены А цеха №1,5
		}
		elseif ($id_prof -eq 50037992)
		{
			$managers['50000459']=$TubNum		# Цех №1 смена Б - Старший мастер смены Б цеха №1,5
		}
		elseif ($id_prof -eq 50009713)
		{
			$managers['50000319']=$TubNum		# Цех №2 - Начальник цеха №2
			$top_managers['50000461']=$TubNum	# Цех №2 смена А - Начальник цеха №2
			$top_managers['50032499']=$TubNum	# Цех №2 смена Б - Начальник цеха №2
			$top_managers['50000473']=$TubNum	# Цех №6 - Начальник цеха №2
		}
		elseif ($id_prof -eq 50009694)
		{
			$managers['50000461']=$TubNum		# Цех №2 смена А - Старший мастер смены A цеха №2
		}
		elseif ($id_prof -eq 50009591)
		{
			$managers['50032499']=$TubNum		# Цех №2 смена Б - Старший мастер смены Б цеха №2
		}
		elseif ($id_prof -eq 50010333)
		{
			$managers['50000321']=$TubNum		# Цех №3 - Начальник цеха №3,4
			$top_managers['50000464']=$TubNum	# Цех №3 смена А - Начальник цеха №3,4
			$top_managers['50000463']=$TubNum	# Цех №3 смена Б - Начальник цеха №3,4
			$top_managers['50000466']=$TubNum	# Цех №4 смена А - Начальник цеха №3,4
			$top_managers['50000467']=$TubNum	# Цех №4 смена Б - Начальник цеха №3,4
		}
		elseif ($id_prof -eq 50000464)
		{
			$managers['50000464']=$TubNum		# Цех №3 смена А - Старший мастер смены А цеха №3,4
		}
		elseif ($id_prof -eq 50009383)
		{
			$managers['50000463']=$TubNum		# Цех №3 смена Б - Старший мастер смены Б цеха №3,4
		}
		elseif ($id_prof -eq 50010360)
		{
			$managers['50000466']=$TubNum		# Цех №4 смена А - Старший мастер смены А цеха №3,4
		}
		elseif ($id_prof -eq 50009383)
		{
			$managers['50000467']=$TubNum		# Цех №4 смена Б - Старший мастер смены Б цеха №3,4
		}
		elseif ($id_prof -eq 50037990)
		{
			$managers['50000471']=$TubNum		# Цех №5 смена А - Старший мастер смены А цеха №1,5
		}
		elseif ($id_prof -eq 50037992)
		{
			$managers['50000472']=$TubNum		# Цех №5 смена Б - Старший мастер смены Б цеха №1,5
		}
		elseif ($id_prof -eq 50046489)
		{
			$managers['50000473']=$TubNum		# Цех №6 - Старший мастер производственного участка
		}
		elseif ($id_prof -eq 50010291)
		{
			$managers['50000456']=$TubNum		# Участок литья - Начальник участка литья
			$managers['50000478']=$TubNum
			$managers['50000479']=$TubNum
			$managers['50001845']=$TubNum
			$managers['50001846']=$TubNum
			$managers['50045747']=$TubNum
			$managers['50045745']=$TubNum		
		}
		elseif ($id_prof -eq 50030531)
		{
			$managers['50030386']=$TubNum		# Участок подготовки кож - Начальник производственного участка
			$managers['50000460']=$TubNum
		}
	}
}

$csv | Select-Object CodeOrg, Org, FIO, ProfFull, DepatFull, Tubnum, Prof, Depat, Birthday, EmploymentDate, Familia, Name, Otch, ManagerNum, HZ2, IDProfFull, IDDepatFull | ForEach-Object{
	$id_prof = $_.IDProfFull
	$id_dep = $_.IDDepatFull
	$level = $_.HZ2
	$CodeOrg = $_.CodeOrg
	$FIO = $_.FIO
	$TubNum = $_.TubNum
	$exTubNum = $_.TubNum.Substring(0,3)
	$birthday = $_.Birthday
	$currentDate = [Int64](Get-Date (Get-Date -Format dd.MM.yyyy) -UFormat %s)
	$empDate = [Int64](Get-Date $_.EmploymentDate -UFormat %s)
	$empType = "Офис"
	If (($CodeOrg -eq "1000") -and !(($exTubNum -eq "090") -or ($exTubNum -eq "330") -or ($exTubNum -eq "320") -or ($_.Depat -match "магазин")))
	{
		Out-File $FullLog -InputObject ("Обработка "+$_.FIO+'...') -Append -Encoding "Default"
		$user = get-ADUser -Searchbase $BWADPath -Filter {EmployeeID -eq $TubNum} -Properties EmployeeID, Department, Title, Company, Description, fullDeparment, birthDay, comment, manager, distinguishedName, employeeType
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
					Out-File $NewUserFile -InputObject ($FIO + ": дата приема на работу удалена") -Append -Encoding "Default"
				}
			}
			#Руководитель
			if ($_.ManagerNum -ne '')
				{
				$id_prof = $_.IDProfFull
				$id_dep = $_.IDDepatFull
				$ManagerNum = $_.ManagerNum
				if ($TubNum -eq $ManagerNum)
				{
					if ($managers[$id_dep] -ne $null)
					{
						$ManagerNum = $managers[$id_dep]
					}
					elseif ($top_managers[$id_dep] -ne $null)
					{
						$ManagerNum = $top_managers[$id_dep]
					}
				}
				$manager = get-ADUser -Searchbase $BWADPath -Filter {EmployeeID -eq $ManagerNum} -Properties distinguishedName
				if (($user.manager -eq $manager.distinguishedName) -or ($manager.distinguishedName -eq ''))
				{
					Out-File $FullLog -InputObject ("Руководитель верный") -Append -Encoding "Default"
				}
				else
				{
					$user | set-ADUser -manager $manager.distinguishedName
					Out-File $FullLog -InputObject ("Руководитель обновлен") -Append -Encoding "Default"
					Out-File $NewUserFile -InputObject ($FIO+" изменен руководитель") -Append -Encoding "Default"
				}
			#Назначение группы и аттрибута Офис или Производство
			if ((($level -eq "v5") -or ($level -eq "v6") -or ($level -eq "v7")) -and ($user.employeeType -ne "Офис"))
			{
				Set-ADUser $user -Replace @{employeeType="Офис"}
            			Out-File $FullLog -InputObject ("Тип работника обновлен") -Append -Encoding "Default"
            			Out-File $NewUserFile -InputObject ($FIO+" изменен тип работника") -Append -Encoding "Default"
				Add-ADGroupMember -Identity ("Офис") -Members $user
			}
			else
			{
				if ((($level -eq "v1") -or ($level -eq "v2") -or ($level -eq "v3") -or ($level -eq "v4")) -and ($user.employeeType -ne "Рабочие"))
				{
					Set-ADUser $user -Replace @{employeeType="Рабочие"}
	            			Out-File $FullLog -InputObject ("Тип работника обновлен") -Append -Encoding "Default"
	            			Out-File $NewUserFile -InputObject ($FIO+" изменен тип работника") -Append -Encoding "Default"
					Add-ADGroupMember -Identity ("Рабочие") -Members $user
				}
				else
				{
					Out-File $FullLog -InputObject ("Тип работника верный") -Append -Encoding "Default"
				}
			}
			}
			# раньше бахалось без проверок
			# $user | set-ADUser -EmployeeID $_.TubNum -Department $_.Depat -Title $_.Prof -Company $_.Org -Description $_.Prof
			# Out-File $goodoutfile -InputObject ($_.FIO+" найден и его данные обновлены ") -Append -Encoding "Default" 
			# Out-File $FullLog -InputObject ($_.FIO+" найден и его данные обновлены ") -Append -Encoding "Default"
			# Clear-Variable -Name user
			# Clear-Variable -Name CodeOrg
			$ADGroup = $_.Depat
			$GroupExist = Get-ADGroup -SearchBase 'DC=belwest,DC=corp' -LDAPFilter “(name=$ADGroup)” -Properties Description
			Out-File $FullLog -InputObject ("Обработка группы пользователя") -Append -Encoding "Default"
			#Write-Host ---
			#Write-Host $GroupExist.Description
			#Write-Host $_.DepatFull
			if ($GroupExist -ne $null)
			{
				if ($GroupExist.Description -eq $_.DepatFull)
				{
					Out-File $FullLog -InputObject ("Описание группы "+$_.Depat+" верное") -Append -Encoding "Default"
				}
				else
				{
					$GroupExist | Set-ADGroup -Description $_.DepatFull
					Out-File $FullLog -InputObject ("Описание группы "+$_.Depat+" обновлено на "+$_.DepatFull) -Append -Encoding "Default"
				}
				$UserObject = Get-ADUser -Filter {displayname -eq $FIO } -Property "MemberOf"
				#$UserObject.MemberOf | ForEach-Object {
				# Write-Host $UserObject.MemberOf
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
			}
			else 
			{
				Out-File $newgroupfile -InputObject ($ADGroup +" добавлена в AD ") -Append -Encoding "Default"
				New-ADGroup $_.Depat -Path $BWADPath -GroupScope Global -Description $_.DepatFull -PassThru -Verbose
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
			If (!(($exTubNum -eq "090") -or ($exTubNum -eq "330") -or ($exTubNum -eq "320") -or ($_.Depat -match "магазин")) -and ($checkUser -eq $null))
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
				#Назначение группы и аттрибута Офис
				if (($level -eq "v5") -or ($level -eq "v6") -or ($level -eq "v7"))
				{
					Set-ADUser $user -Replace @{employeeType="Офис"}
					Add-ADGroupMember -Identity ("Офис") -Members $user
				}
				else
				{
					Set-ADUser $user -Replace @{employeeType="Рабочие"}
					Add-ADGroupMember -Identity ("Рабочие") -Members $user
				}
				if ($NewUser -eq $null)
				{
					Out-File $NewAccountFile -InputObject ("ОШИБКА СОЗДАНИЯ! - "+$_.FIO+", №"+$_.TubNum+", "+$samAccountName) -Append -Encoding "Default"
				}
				Out-File $NewAccountFile -InputObject ($_.FIO+", "+$_.ProfFull+", "+$_.Depat+", №"+$_.TubNum+", "+$samAccountName) -Append -Encoding "Default"
				#Экспорт нового сотрудника для столовой
				Out-File $gedemin -InputObject ($_.TubNum.TrimStart("0")+";"+$_.FIO+";"+$_.DepatFull) -Append -Encoding "Default"
			}
			else
			{
				Out-File $badoutfile -InputObject ($_.FIO+" с №"+$_.TubNum+" не найден в AD в OU СООО Белвест") -Append -Encoding "Default" 
				Out-File $FullLog -InputObject ($_.FIO+" с №"+$_.TubNum+" не найден в AD в OU СООО Белвест") -Append -Encoding "Default"
			}
			Clear-Variable -Name checkUser
			Clear-Variable -Name CodeOrg
			Clear-Variable -Name empDate
		}
	}
	else
	{
		Out-File $FullLog -InputObject ($_.FIO+" с №"+$_.TubNum+" не является сотрудником Офис СООО Белвест"+$ADGroup) -Append -Encoding "Default"
	}
} | Out-Null

$GedeminFile=$InputFilepath+'Gedemin\SAPusers.txt'
$InputFiredFile=$InputFilepath+(Get-Date  -Format "yyyyMMdd")+'personalu.txt'
Get-Content $InputFile -Encoding UTF8 | Out-File $GedeminFile -encoding default
$GedeminFiredFile=$InputFilepath+'Gedemin\SAPusersfired.txt'
Get-Content $InputFiredFile -Encoding UTF8 | Out-File $GedeminFiredFile -encoding default
