<#
.SYNOPSIS
    Simple PowerShell for programmers that know how to read PowerShell, write HTML, and have a variable data file ready.
.DESCRIPTION
    A longer description of the function, its purpose, common use cases, etc.
.NOTES
    Information or caveats about the function e.g. 'This function is not supported in Linux'
.LINK
    Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE
    Test-MyTestFunction -Verbose
    Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines
#>

. "_private\credential.ps1"

$ExcelFile = "C:\Users\user\Documents\EmailList.xlsx"
$SheetName = "Sheet1"
$Recipients = Import-Excel -Path $ExcelFile -WorksheetName $SheetName
$SMTPServer = "smtp.gmail.com"
$SMTPPort = 587
$From = "youremail@gmail.com"
$Subject = "Email Subject"
$BodyTemplate = Get-Content -Path "C:\Users\user\Documents\EmailBodyTemplate.html" -Raw

foreach ($Recipient in $Recipients) {
    $To = $Recipient.EmailAddress
    $Body = $BodyTemplate -replace "__FirstName__", $Recipient.FirstName -replace "__LastName__", $Recipient.LastName -replace "__Email__", $Recipient.EmailAddress
    Send-MailMessage -To $To -From $From -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -UseSsl -Credential $credential -BodyAsHtml
}
