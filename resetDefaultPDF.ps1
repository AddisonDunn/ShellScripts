for /f “usebackq” %i in (`reg query hkcr /f “AppX”`) do reg query %i\DefaultIcon | find “MicrosoftEdgePDF” && reg add %i /v NoOpenWith /t REG_SZ /f

$key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion'
(Get-ItemProperty -Path $key -Name ProgramFilesDir).ProgramFilesDir


$registryPath = "HKEY_CURRENT_USER\SOFTWARE\Classes\" + $key

$Name = "NoOpenWith"


New-ItemProperty -Path $registryPath -Name $name -Value $value `-PropertyType DWORD -Force | Out-Null