# for storing credential outside of github
$credential = Get-Credential

#-------------------- or --------------------------

$SMTPUsername = "youremail@gmail.com"
$SMTPPassword = "yourpassword"
$credential = (New-Object System.Management.Automation.PSCredential ($SMTPUsername, (ConvertTo-SecureString $SMTPPassword -AsPlainText -Force)))