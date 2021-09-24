#texto de introducao
$texto1 = "`n Desenvolvido por Heitor Augusto Galindo
 heitor.augusto.galindo@live.com
 https://github.com/Extars01/Instala-impressora
 Versao 1.1
 
 DRIVER: PCL6 DRIVER FOR UNIVERSAL PRINT Ver. 4.32.0.0
 http://support.ricoh.com/bb/html/dr_ut_e/rc3/model/p_i/p_i.htm?lang=en`n"

#aviso ao usuario
$aviso1 = "     PARA EXECUTAR ESSE PROGRAMA VOCE DEVE EXECUTAR ELE EM MODO ADMINISTRADOR"
$aviso2 = "     PARA EXECUTAR ESSE PROGRAMA EXECUTE O COMANDO ABAIXO NO POWERSHELL EM MODO ADMINISTRADOR"
$comando1 = "                Set-ExecutionPolicy Unrestricted -Force`n"
$texto2 = "     ATENCAO, SIGA OS PASSOS ABAIXO PARA EXECUTAR ESSE PROGRAMA:
$aviso1`n$aviso2`n"
Write-Host $texto1 -ForegroundColor Magenta
#Write-Host $texto2
#Write-Host $comando1 -ForegroundColor Green

#configuracao dos repositorios
$driver = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001327/0001327395/V43200/z94536L1a.exe"
$arquivo_exe = "C:/temp/z94536L1a.exe"
$arquivo_zip = "C:/temp/z94536L1a.zip"
$pasta_driver = "C:/temp/z94536L1a"
$nome_driver = "PCL6 Driver for Universal Print"

#imput usuario
do {
    do {
        $nome_impressora = Read-Host -Prompt "Digite o nome da impressora"
        $ip_impressora = Read-Host -Prompt "Digite o IP da impressora"
        Write-Host "O nome da impressora e: "$nome_impressora -ForegroundColor Green
        Write-Host "O ip da impressora e: "$ip_impressora -ForegroundColor Green
        $confirma = Read-Host -Prompt "As informacoes estao corretas? [y/n]"
        if ($confirma -eq "n" ) {
        }
    } until($confirma -eq "y")
    $nome_porta = "$nome_impressora - $ip_impressora"

    #validacao de ambiente
    $validacao_exe = Test-Path $arquivo_exe -PathType Leaf
    $validacao_zip = Test-Path $arquivo_zip -PathType Leaf
    $validacao_pasta = Test-Path $pasta_driver

    #condicao de download
    if ($validacao_exe -eq $false -And $validacao_zip -eq $false) {
        Write-Output "...REALIZANDO DOWNLOAD..."
        Invoke-WebRequest -Uri $driver -OutFile $arquivo_exe
    }
    else {
        Write-Output "...DOWNLOAD DESNECESSARIO..."
    }

    #condicao de conversao
    if ($validacao_exe -eq $true -Or $validacao_zip -eq $false) {
        Write-Output "...CONVERTENDO .EXE EM .ZIP..."
        Move-Item $arquivo_exe $arquivo_zip
    }
    else {
        Write-Output "...CONVERSAO DESNECESSARIA..."
    }

    #condicao de expansao
    if ($validacao_pasta -eq $false) {
        Write-Output "...EXPANDINDO ARQUIVOS..."
        Expand-Archive $arquivo_zip $pasta_driver
    }
    else {
        Write-Output "...EXPANSAO DESNECESSARIA..."
    }

    #Remocao da impressora/driver/porta quem podem gerar conflitos
    Remove-Printer -Name $nome_impressora -ErrorAction:SilentlyContinue
    Remove-PrinterPort -Name $nome_porta -ErrorAction:SilentlyContinue

    #preparando driver
    Write-Output "...PREPARANDO DRIVER..."
    pnputil.exe -a C:\temp\z94536L1a\disk1\oemsetup.inf > $null

    #adicionando driver
    Write-Output "...ADICIONANDO DRIVER PCL6...`n"
    Add-PrinterDriver -Name $nome_driver

    #adicionando porta
    Write-Output "...CRIANDO PORTA: $nome_porta ..."
    Add-PrinterPort -Name $nome_porta -PrinterHostAddress $ip_impressora
    
    Write-Host "...ADICIONANDO IMPRESSORA: $nome_impressora...`n"
    Add-Printer -Name $nome_impressora -DriverName $nome_driver -PortName $nome_porta

$finaliza = Read-Host -Prompt "DESEJA FINALIZAR A INSTALACAO? [y/n]"
    if ($finaliza -eq "n") {
    }
} until($finaliza -eq "y")