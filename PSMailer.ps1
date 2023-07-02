$ExcelFile = "C:\Users\user\Documents\EmailList.xlsx"
$SheetName = "Sheet1"
$Recipients = Import-Excel -Path $ExcelFile -WorksheetName $SheetName
$SMTPServer = "smtp.gmail.com"
$SMTPPort = 587
$SMTPUsername = "youremail@gmail.com"
$SMTPPassword = "yourpassword"
$From = "youremail@gmail.com"
$Subject = "Email Subject"
$BodyTemplate = Get-Content -Path "C:\Users\user\Documents\EmailBodyTemplate.html" -Raw

foreach ($Recipient in $Recipients) {
    $To = $Recipient.EmailAddress
    $Body = $BodyTemplate -replace "__FirstName__", $Recipient.FirstName -replace "__LastName__", $Recipient.LastName -replace "__Email__", $Recipient.EmailAddress
    Send-MailMessage -To $To -From $From -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -UseSsl -Credential (New-Object System.Management.Automation.PSCredential ($SMTPUsername, (ConvertTo-SecureString $SMTPPassword -AsPlainText -Force))) -BodyAsHtml
}
