if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
  Exit
 }
}

echo "Installation du generateur de facture d'encan"
$python_path = $PSScriptRoot + "\python-3.9.0.exe"
$requirements_path = $PSScriptRoot + "\requirements.txt"

echo "--- Installation de python 3.9 ---"
echo "Telechargement de l'installeur python."
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.9.0/python-3.9.0-amd64.exe" -OutFile $python_path

echo "starting install job"
$proc = Start-Process $python_path -NoNewWindow -PassThru
$proc.WaitForExit()
echo "install job done"

echo "Installation des modules de dependances."
py -m pip install -r $requirements_path
Read-Host "Installation complete appuyez sur [Entrer] pour continuer."


	