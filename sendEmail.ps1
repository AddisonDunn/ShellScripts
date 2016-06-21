$username = "DOMAIN_USERNAME"
$password ="DOMAIN_PASSWORD"
$secstr = New-Object -TypeName System.Security.SecureString
$password.ToCharArray() | ForEach-Object {$secstr.AppendChar($_)}
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $secstr
$subject = "SUBJECT_LINE"
$body = Get-Content "informationToBeMailed.txt"
 
Send-MailMessage -to "DOMAIN_USERNAME <DOMAIN_USERNAME@DOMAIN.com>" -From "DOMAIN_USERNAME <DOMAIN_USERNAME@DOMAIN.com>" -Subject $subject -Body $body -SmtpServer "SERVER_NAME"