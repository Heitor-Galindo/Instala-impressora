Write-Output "`n"
Write-Output 'Desenvolvido por Heitor Augusto Galindo'
Write-Output 'heitor.augusto.galindo@live.com'
Write-Output '17/09/2021'
Write-Output 'Versao 1.0'
Write-Output "`n"

$nome_impressora = Read-Host -Prompt "Digite o nome da impressora"
$ip_impressora = Read-Host -Prompt "Digite o IP da impressora"
Write-Host 'O nome da impressora e: '$nome_impressora -ForegroundColor Green
Write-Host 'O ip da impressora e: '$ip_impressora -ForegroundColor Green
#-ErrorAction:SilentlyContinue

$driver = 'https://support.ricoh.com/bb/pub_e/dr_ut_e/0001327/0001327395/V43200/z94536L1a.exe'
$pasta_driver = 'C:/temp/z94536L1a.exe'
$nome_driver = 'PCL6 Driver for Universal Print'
$nome_porta = "$nome_impressora - $ip_impressora"

Write-Output '.....PREPARANDO O AMBIENTE'
Remove-Item C:\temp\z94536L1a.exe -ErrorAction:SilentlyContinue
Remove-Item C:\temp\z94536L1a.zip -ErrorAction:SilentlyContinue
Remove-Item C:\temp\z94536L1a -Confirm:$false -Force -Recurse -ErrorAction:SilentlyContinue
Remove-Printer -Name $nome_impressora -ErrorAction:SilentlyContinue
Remove-PrinterDriver -Name $nome_driver -ErrorAction:SilentlyContinue
Remove-PrinterPort -Name $nome_porta -ErrorAction:SilentlyContinue
Start-Sleep -Seconds 2

Write-Output '.....BAIXANDO DRIVER'
Invoke-WebRequest -Uri $driver -OutFile $pasta_driver
Start-Sleep -Seconds 2

Write-Output '.....CONVERTENDO'
Move-Item C:\temp\z94536L1a.exe C:\temp\z94536L1a.zip
Start-Sleep -Seconds 2

Write-Output '.....EXTRAINDO'
Expand-Archive -Path 'C:\temp\z94536L1a.zip' -DestinationPath C:\temp\z94536L1a
Start-Sleep -Seconds 2

Write-Output '.....ADICIONANDO DRIVER'
pnputil.exe -a C:\temp\z94536L1a\disk1\oemsetup.inf > $null
Start-Sleep -Seconds 2

Write-Output '.....ADICIONANDO IMPRESSORA'
Add-PrinterDriver -Name $nome_driver
Add-PrinterPort -Name $nome_porta -PrinterHostAddress $ip_impressora
Add-Printer -Name $nome_impressora -DriverName $nome_driver -PortName $nome_porta
Start-Sleep -Seconds 2

Write-Output '.....PRESSIONE ENTER PARA FINALIZAR.....'
Pause