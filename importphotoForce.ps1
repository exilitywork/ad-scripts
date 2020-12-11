Import-Module ActiveDirectory
$PhotoPath="\\belw-rs3617-2\install\Work Utils\hr_users\" # путь к папке с фото
$InputFilepath="\\belw-rs3617-2\install\Work Utils\hr_users\" # путь к папку с файлами выгрузки
#ниже есть ссылка на тестовый файл если будут изменения
$InputFile = $InputFilepath+(Get-Date  -Format "yyyyMMdd")+'personal.txt' # новый путь и каждые день новый файл
$Logfilepath="\\belw-rs3617-2\install\Work Utils\hr_users\Log\" #папка логов
$BWADPath = 'DC=belwest,DC=corp' # OU для поиска
$ImageImportLog=$Logfilepath+'ImageImportLog'+(Get-Date  -Format "yyyyMMdd")+'.txt' #файл лога

#тестовый файл импорта
#$InputFile="\\belw-rs3617-2\install\Windows\Powershell\Корректировка персонала\testv131.csv" 

Out-File $ImageImportLog -InputObject ("Обработка на дату "+(Get-Date -Format "dd.MM.yyyy")) -Append -Encoding "Default"
Out-File $ImageImportLog -InputObject ("Используется файл "+$InputFile) -Append -Encoding "Default"

#Делаем временный файл с нужной кодировкой
$tempcsv=".\temp.csv" 
Get-Content $InputFile -Encoding UTF8 | Out-File $tempcsv -Encoding unicode
$header = "HZ","HZ1","HZ2","CodeOrg","Org","DateIn","DateOut","TubNum","FIO","Familia","Name","Otch","IDProfFull","ProfFull","IDDepatFull","DepatFull","SapLogin","WSTel","WMTel","PMTel","PSTel","WMail","PMail","WSkype","PSkype","Prof","Depat","IDDepat"
$csv = Import-CSV $tempcsv -header $header -Delimiter ';'

$csv | Select-Object CodeOrg, Org, FIO, ProfFull, DepatFull, Tubnum, Prof, Depat | ForEach-Object{
    $CodeOrg = $_.CodeOrg
    $FIO = $_.FIO
    Out-File $ImageImportLog -InputObject ("Обработка "+$_.FIO+'...') -Append -Encoding "Default"
    #If ($CodeOrg.Equals('1000')) #берем только белвест
    #    {
        $user = get-ADUser -Searchbase $BWADPath -Filter {displayname -eq $FIO } -Properties EmployeeID , Department, Title, Company, Description, fullDeparment, thumbnailPhoto
        If ($null -ne $user ) #если есть такой пользователь
            {
            #Out-File $ImageImportLog -InputObject ($_.FIO+" обрабатывается") -Append -Encoding "Default"
            $PhotoPathFull= $PhotoPath+($_.TubNum)+'.jpg'
            #Write-Host $PhotoPathFull
            If (Test-Path  -Path $PhotoPathFull)
                {
                    #write-host 'foto yes'
                    $photo = [byte[]](Get-Content $PhotoPathFull -Encoding byte)
                    #Убрана проверка наличия фото. Часто стали менять фотографии. 16.11.2020
                    #if ($user.thumbnailPhoto -eq $null)
                    #    {
                        Set-ADUser $user -Replace @{thumbnailPhoto=$photo}
                        Out-File $ImageImportLog -InputObject ($_.FIO+" фото обновлено") -Append -Encoding "Default"
                    #    Write-Host +
                    #    }
                    #else
                    #    {
                    #    Out-File $ImageImportLog -InputObject ($_.FIO+" фото уже было загружено ранее") -Append -Encoding "Default"
                    #    } 
                }
            else
                {
                    #Write-Host 'foto nain'
                    Out-File $ImageImportLog -InputObject ($_.FIO+" фото для импорта остутствует") -Append -Encoding "Default"
                }
            }
        else 
            {
                Out-File $ImageImportLog -InputObject ('Пользователь отсутствует в АД') -Append -Encoding "Default"
            }
    #    }
    #else
    #    {
    #    Out-File $ImageImportLog -InputObject ("Не работает на Белвесте") -Append -Encoding "Default"
    #    }
        Clear-Variable -Name user
        Clear-Variable -Name CodeOrg
        
                                                                                                 }