$username = "USERNAME"
$password ="PASSWORD"
$secstr = New-Object -TypeName System.Security.SecureString
$password.ToCharArray() | ForEach-Object {$secstr.AppendChar($_)}
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $secstr

# The connection name can be found at Control Panel -> Printers -> Properties of the desired printer
$connectionName = "\\PRINTER_CONNECTION\PRINTER_NAME"

# Adds printer based on connection name
add-printer -connectionname $connectionName
