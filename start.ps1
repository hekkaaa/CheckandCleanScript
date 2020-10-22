$comp = (Import-Csv C:\Temp\explorefile\Hosts1.csv).Name

write-Host "Файл CSV импортирован" -ForegroundColor "Green"

function FuncInfo {

Param($Computer)

$Error.Clear()

#***************** ИНТЕРФЕЙС **************************************************************************************************#

#Данные нужно указывать в кавычках что бы они были копределены в переменной как СТРОКИ.

$Programms = 'notepad*' #Какую программу ищем

$version1 = '8.8.0' #Версия программы

$Module = "0" #модуль поиска - "0" / модуль поиска удаления/переустановки - "1"

# настройки дистрибутива

$Share = "\\$Computer\c$\Temp" #Не менять шару развертования!

$sourcePath = "C:\Temp\Notepad++7_8_9.msi" # Указать путь откуда брать установщик

$CMDLine1 = "/S" #ключи запуска. Для exe пример: '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-'

$CMDLineMSI = "/i `"C:\Temp\Notepad++7_8_9.msi`" /qb- /norestart" #ключи запуска. MSI пример: "/i `"C:\Temp\Notepad++7_8_9.msi`" /qb- /norestart"

$UninstalCMDLine = '/S' #Ключи удаления exe

#*****************************************************************************************************************************#

function ExplorerModule {

Invoke-Command -ComputerName $Computer -ScriptBlock{

param($Programms,$version1,$Module)

$SelectInfoProgramm = Get-Package -Name $Programms -MaximumVersion $version1

if ($SelectInfoProgramm.Name){

if ($SelectInfoProgramm.Version -eq $version1){

$exit = 'Программа '+'"'+$SelectInfoProgramm.Name +'"' +' актуальной версии ' + $version1

return $exit

}

else {

$VariableCheck = '' # очистка переменной

$VariableCheck = '*'+$Programms

$InstallProviderCheck = Get-Package -Name $Programms

$exit ="Программа "+ $SelectInfoProgramm.Name +" Версии " + $SelectInfoProgramm.Version + " ниже необходимой: " + $version1 + " | Установлен с помощью: " + $InstallProviderCheck.ProviderName

return $exit

}

}

else {

$exit = 'Программа '+ '"' + $Programms + '"' +' не найдена на ПК либо версии выше указанной ' + $version1

return $exit

}

} -ArgumentList $Programms,$version1,$Module -ErrorAction SilentlyContinue

}

function UninstalInstalModule {

Invoke-Command -ComputerName $Computer -ScriptBlock{

param($Programms,$version1,$CMDLine1,$sourcePath,$Module,$UninstalCMDLine)

$SelectInfoProgramm = Get-Package -Name $Programms -MaximumVersion $version1

if ($SelectInfoProgramm.Name){

if ($SelectInfoProgramm.Version -eq $version1){

$exit = 'Программа '+'"'+$SelectInfoProgramm.Name +'"' +' актуальной версии ' + $version1

return $exit

}

else {

$VariableCheck = '' # очистка переменной

$VariableCheck = '*'+$Programms

$CheckProcess = Get-Process -Name $VariableCheck

if ($CheckProcess){

###### Если процесс запущен то сразу конец

$ex1 = 'Запущенные процессы ' + $CheckProcess.Name +' найдены. Процесс остановлен.'

$exit ='Программа '+ $SelectInfoProgramm.Name +' Версии ' + $SelectInfoProgramm.Version + ' ниже необходимой:' + $version1 + ' Требуется переустановка ' + $ex1

return $exit

}

else {

$Pkg = Get-Package -Name $SelectInfoProgramm.Name

# Тут начинаются ветки.

if ($Pkg.ProviderName -eq "Programs") {

# тут ветка удаления EXE

$ex2 = "Запущенных процессов нет"

$Pkg = Get-Package -Name $SelectInfoProgramm.Name

$UninstallCommand = $Pkg.Meta.Attributes['UninstallString']

$code = (Start-Process -FilePath $UninstallCommand -ArgumentList $UninstalCMDLine -Verb runAs -Wait -Passthru).ExitCode

if ($code -eq 0){

########## ЛОГИКА УСТАНОВКИ MSI и EXE

$ProgramName = Split-Path $sourcePath -Leaf

$VariableExtension = Get-ItemProperty -Path "C:\Temp\$ProgramName"

if ($VariableExtension.Extension -eq ".exe"){

#Установщих exe

$code = (Start-Process "C:\Temp\$ProgramName" -ArgumentList $CMDLine1 -Verb runAs -Wait -Passthru).ExitCode;

if ($code -eq 0)

{

$ex2_2 = "$ProgramName Установлен!"

$exit ='Программа '+ $SelectInfoProgramm.Name +' Версии ' + $SelectInfoProgramm.Version + ' ниже необходимой: ' + $version1 + ' Переустановлено ' + " | " + $ex2_2

Remove-Item -Path "C:\Temp\$ProgramName" -Force | Out-Null

return $exit

}

else

{

$ex2_2_2 ="BAD - TOTALY FAILED - $ProgramName не установлен! $code"

$exit ='Программа '+ $SelectInfoProgramm.Name +' Версии ' + $SelectInfoProgramm.Version + ' ниже необходимой: ' + $version1 + ' Нужна ручная установка ' + $ex2 + " | " + $ex2_2_2

Remove-Item -Path "C:\Temp\$ProgramName" -Force | Out-Null

return $exit

}

}

elseif ($VariableExtension.Extension -eq ".msi"){

#Установка MSI

$code = (Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList $CMDLineMSI -Verb runAs -Wait -Passthru).ExitCode;

if ($code -eq 0)

{

$ex2_2 = "$ProgramName Установлен!"

$exit ='Программа '+ $SelectInfoProgramm.Name +' Версии ' + $SelectInfoProgramm.Version + ' ниже необходимой: ' + $version1 + ' Переустановлено ' + " | " + $ex2_2

Remove-Item -Path "C:\Temp\$ProgramName" -Force | Out-Null

return $exit

}

else

{

$ex2_2_2 ="BAD - TOTALY FAILED - $ProgramName не установлен! $code"

$exit ='Программа '+ $SelectInfoProgramm.Name +' Версии ' + $SelectInfoProgramm.Version + ' ниже необходимой: ' + $version1 + ' Нужна ручная установка ' + $ex2 + " | " + $ex2_2_2

Remove-Item -Path "C:\Temp\$ProgramName" -Force | Out-Null

return $exit

}

}

else {

Write-Host 'ERROR FORMAT install FILES!' -ForegroundColor 'Red'

}

}

else {

$ex1_2 = 'Что то пошло не так ПО еще на месте при удалении программы с ProviderName EXE '

$exit ='Программа '+ $SelectInfoProgramm.Name +' Версии ' + $SelectInfoProgramm.Version + ' ниже необходимой: ' + $version1 + ' Переустановка ' + $ex2 +' | ' + $ex1_2

return $exit

}

}

elseif ($Pkg.ProviderName -eq "msi"){

# ветка удаление/установки MSI

$ex2 = "Запущенных процессов нет"

Get-Package -Name $SelectInfoProgramm.Name | Uninstall-Package -Force | Out-Null

$PostInistallCheck = Get-Package -Name $SelectInfoProgramm.Name

if ($PostInistallCheck){

$ex1_2 = 'Что то пошло не так ПО еще на месте при удалении при удалении программы с ProviderName MSI'

$exit ='Программа '+ $SelectInfoProgramm.Name +' Версии ' + $SelectInfoProgramm.Version + ' ниже необходимой: ' + $version1 + ' Переустановка ' + $ex2 +' | ' + $ex1_2

return $exit

}

else {

########## ЛОГИКА УСТАНОВКИ MSI и EXE

$ProgramName = Split-Path $sourcePath -Leaf

$VariableExtension = Get-ItemProperty -Path "C:\Temp\$ProgramName"

if ($VariableExtension.Extension -eq ".exe"){

#Установщих exe

# Тут требуется проверка и доработка импорта ключей.

$code = (Start-Process "C:\Temp\$ProgramName" -ArgumentList $CMDLine1 -Verb runAs -Wait -Passthru).ExitCode;

if ($code -eq 0)

{

$ex2_2 = "$ProgramName Установлен!"

$exit ='Программа '+ $SelectInfoProgramm.Name +' Версии ' + $SelectInfoProgramm.Version + ' ниже необходимой: ' + $version1 + ' Переустановлено ' + " | " + $ex2_2

Remove-Item -Path "C:\Temp\$ProgramName" -Force | Out-Null

return $exit

}

else

{

$ex2_2_2 ="BAD - TOTALY FAILED - $ProgramName не установлен! $code"

$exit ='Программа '+ $SelectInfoProgramm.Name +' Версии ' + $SelectInfoProgramm.Version + ' ниже необходимой: ' + $version1 + ' Нужна ручная установка ' + $ex2 + " | " + $ex2_2_2

Remove-Item -Path "C:\Temp\$ProgramName" -Force | Out-Null

return $exit

}

}

elseif ($VariableExtension.Extension -eq ".msi"){

#Установка MSI

$code = (Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList $$CMDLineMSI -Verb runAs -Wait -Passthru).ExitCode;

if ($code -eq 0)

{

$ex2_2 = "$ProgramName Установлен!"

$exit ='Программа '+ $SelectInfoProgramm.Name +' Версии ' + $SelectInfoProgramm.Version + ' ниже необходимой: ' + $version1 + ' Переустановлено ' + " | " + $ex2_2

Remove-Item -Path "C:\Temp\$ProgramName" -Force | Out-Null

return $exit

}

else

{

$ex2_2_2 ="BAD - TOTALY FAILED - $ProgramName не установлен! $code"

$exit ='Программа '+ $SelectInfoProgramm.Name +' Версии ' + $SelectInfoProgramm.Version + ' ниже необходимой: ' + $version1 + ' Нужна ручная установка ' + $ex2 + " | " + $ex2_2_2

Remove-Item -Path "C:\Temp\$ProgramName" -Force | Out-Null

return $exit

}

}

else {

Write-Host 'ERROR FORMAT install FILES!' -ForegroundColor 'Red'

}

}

}

else {

Write-Host 'ERROR FORMAT Uninstal FILES!' -ForegroundColor 'Red'

}

}

}

}

else {

$exit = 'Программа '+ '"' + $Programms + '"' +' не найдена на ПК либо версии выше указанной ' + $version1

return $exit

}

} -ArgumentList $Programms,$version1,$CMDLine1,$sourcePath,$Module,$UninstalCMDLine -ErrorAction SilentlyContinue

}

if ($Module -eq "0"){

return ExplorerModule

}

else {

Copy-Item $sourcePath -Destination $Share -Force

return UninstalInstalModule

}

#Ловля ошибок при подклчюении invoke-Command

$ErrorLog = $Error[0].Exception

if ($ErrorLog){

$exit1 = $ErrorLog

return $exit1

}

else {

Out-Null #заглушка.

}

############# Конец функции. ###################

}

#####################################################################################################################

######### создание файла отчета #########

$TimeIndex = Get-Date -Format "dd/MM/yyyy_HH/mm"

$TimeIndex = $TimeIndex +'.txt'

New-Item -Name $TimeIndex -Path "C:\Temp\lotusUPdate" -Force | Out-Null

##### цикл ########

$i = 0 # Переменная для счетчика.

foreach($Computer in $comp) {

$i++

Write-Host "Обрабатывается: " $i " из " $comp.Length #Счетчик

$test1 = Test-Connection -computer $Computer -quiet -Count 2 -Delay 1

if ($test1 -eq 'True') {

$SelectInfoFunc = FuncInfo $Computer

#Запись в лог

#разделено логикой потому как вывод ошибки из функции одновременно летит с успешной проверкой. из invoke-command $error.clear() не работает.

if ($SelectInfoFunc.Length -gt 2){

$Computer + ' | ' + $SelectInfoFunc | Out-File -FilePath C:\Temp\explorefile\$TimeIndex -Append

}

else {

$Computer + ' | ' + $SelectInfoFunc[0] | Out-File -FilePath C:\Temp\explorefile\$TimeIndex -Append

}

}

#Если машина не в сети.

else {

#Запись в лог

"Машина " + $Computer + " не в сети" | Out-File -FilePath C:\Temp\explorefile\$TimeIndex -Append

}

}

Write-Host "Конец процесса" -ForegroundColor "Yellow"
