$username = "addison.dunn"
$password ="Dakoto86"
$secstr = New-Object -TypeName System.Security.SecureString
$password.ToCharArray() | ForEach-Object {$secstr.AppendChar($_)}
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $secstr

$computer_name = "IRPW-BD064Q11"



Enable-PSRemoting -Force #IMPORTANT
Set-ExecutionPolicy UnRestricted
#Enter-PSSession
#psexec \\IRPW-DH7DD72 -u $username -p $password dir

$scriptLocation = "C:\Users\addison.dunn\Documents"

# psexec \\IRPW-BD064Q11 -h -i -c PowerShell C:\Users\addison.dunn\Documents\dummyScript.ps1


# psexec \\IRPW-BD064Q11 -h  help

#@run_file  Run command on every computer listed in the text file specified. ->>>http://ss64.com/nt/psexec.html



#Test-WsMan $IP
#Enter-PSSession
#Invoke-Command -ComputerName COMPUTER -ScriptBlock { COMMAND } -credential USERNAME
#remotely connect AS ADMIN

#$user = false
#function Test-Administrator  
#{  
#    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
#    (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
#}
#if (not $user) {

