# Stop-Computer -computer belw-wsus -Force -Credential belwestcorp\fedorovaa
# В принципе работает и так, если запускать с контролера домена возьмет привелегии пользователя от которого запустили.
# -Force использовать в случае: Stop-Computer : Не удается выполнить команду на конечном компьютере ("belw-mdt")
# из-за следующей ошибки: Невозможно инициировать завершение работы системы, так
# как компьютер используется другими пользователями. - Лбио завершать сеансы, либо Force
# Вин ниже 2008 не потушим черех повершелл.
# Дописать/убрать и запустить.
Write-Host ″Скрипт экстенного выключения серверных ОС на Windows. Если вы уверены что хотите это сделать - наберите YES … или Ctrl+C для отмены″
$abort=read-host "Ваш выбор?"
if ($abort -eq "yes")
{ 
Write-Host "Спасибо. Команды отправляются."
Stop-Computer -computer belw-wsus -Force
Stop-Computer -computer belw-mdt -Force
Stop-Computer -computer belw_sql -Force
Stop-Computer -computer BELW-1C82 -Force
Stop-Computer -computer belw-printer -Force
Stop-Computer -computer belw-rdp -Force
Stop-Computer -computer belw-spapp -Force
Stop-Computer -computer belw-spdb -Force
Stop-Computer -computer belw-spwebapp -Force
Stop-Computer -computer fileserv1 -Force
Stop-Computer -computer bwexcel -Force
Stop-Computer -computer belw-bo -Force
Stop-Computer -computer belw-sqlbw -Force
Write-Host "Команды на отключение отправлены."
Write-Host "Удачного дня"
Start-Sleep 3
}
else 
{
write-host "До свидания"
Start-Sleep 3
}
