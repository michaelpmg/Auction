if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
  Exit
 }
}

echo "Installation du generateur de facture d'encan"
Pause
$DIR=$PSScriptRoot

$python_install_job = Start-Job { .\python-3.9.0.exe InstallAllUsers=0 PrependPath=1 Include_test=0 }

if (-Not (Get-Command 'py' -errorAction SilentlyContinue)){
	echo "--- Installation de python 3.9 ---"
	Pause
	echo "Telechargement de l'installeur python."
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.9.0/python-3.9.0-amd64.exe" -OutFile ".\python-3.9.0.exe"
	echo "Installing Python, please rerun the installer once completed by running .\installer.ps1"

	Wait-Job $python_install_job
	Receive-Job $python_install_job
	Pause
}

rm .\python-3.9.0.exe -Force 2>$null
echo "Installation des modules de dependances."
Pause
py -m pip install -r requirements.txt
Read-Host "Installation complete appuyez sur [Entrer] pour continuer."
Pause


	