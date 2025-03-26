set "file=%~dp0installer_shell_script.ps1"
echo %file%
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%file%""' -Verb RunAs}"
pause