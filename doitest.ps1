<#
.Synopsis
	Automated Web Site Login and Forwarding
.Description
	A PowerShell script for logging into a testing site, extracting the
	relevant information, then sending it through an email-to-SMS service
	to one or more recipients via text message.
.Notes
	Author:		Keith Dobbelaere
	Version:	1.01
	Created:	2019-09-08
	Modified:	2019-09-17
#>
. .\logger.ps1 -logFile 'doitest.log'

Write-Log '(SCRIPT STARTED)' 'INFO'
Write-Log 'doitest.ps1 - PowerShell Script' 'INFO'
Write-Log '(c)2019 Keith Dobbelaere' 'INFO'

#-------------------------------  Configuration ------------------------------------

# Milliseconds to wait for page to load
$timeoutMillis = 20000
# Pin Number (this code is the same for all participants)
$pin = '1590'
# ID Number (this code is unique to you)
$id = '01201970'
# Gmail account (must be set to allow access for less secure apps)
$emailAddr = 'yourusername@gmail.com'
# Gmail password
$password = 'gmailpassword'
# Cell Phone number(s) - Select carrier from table below.
$phoneNums = @(
	[PSCustomObject]@{ number = '5071234567'; carrier = 'Verizon Wireless'; }
	[PSCustomObject]@{ number = '5072345678'; carrier = 'AT&T'; }
)
# Email to SMS services
$SMSTable = @{
	'Alltel' = 'sms.alltelwireless.com'
	'AT&T' = 'txt.att.net'
	'Boost Mobile' = 'sms.myboostmobile.com'
	'Cricket Wireless' = 'mms.cricketwireless.net'
	'MetroPCS' = 'mymetropcs.com'
	'Google Fi' = 'msg.fi.google.com'
	'Republic Wireless' = 'text.republicwireless.com'
	'Sprint' = 'messaging.sprintpcs.com'
	'T-Mobile' = 'tmomail.net'
	'U.S. Cellular' = 'email.uscc.net'
	'Verizon Wireless' = 'vzwpix.com'
	'Virgin Mobile' = 'vmobl.com'
}

#---------------------------------  Load Page  -------------------------------------
Write-Log 'Loading page...' 'INFO'
try {
	$ie = New-Object -ComObject 'internetExplorer.Application'
	#$ie.visible = $true # Uncomment to make Internet Explorer window visible
	$ie.Navigate("doi.testday.com")
}
catch {
	Write-Log 'There was a problem opening Internet Explorer.' 'ERROR'
}
$timeStart = Get-Date
# Wait for page to load
$exitFlag = $false
do {
	sleep -milliseconds 100
	if ($ie.ReadyState -eq 4) {
		$elements = $ie.Document.getElementById('web_checkin')
		$elementMatch = $elements.readyState -match 'complete'
		if ($elementMatch) { $loadTime = (Get-Date).subtract($timeStart) }
	}
	$timeout = ((Get-Date).subtract($timeStart)).TotalMilliseconds -gt $timeoutMillis
	$exitFlag = $timeout -or $elementMatch
} until ($exitFlag)
if ($timeout) {
	Write-Log "There was a problem loading the page.  Operation timed out after $($timeoutMillis/1000) seconds." 'ERROR'
}
else {
	Write-Log "Load Time: $loadTime" 'INFO'
}

#--------------------------------  Submit IDs  ------------------------------------
Write-Log 'Submitting identification...' 'INFO'
try {
	$inputs = $elements.getElementsByTagName('input')
	$inputs[1].value = $pin
	$inputs[2].value = $id
	$inputs[3].click()
}
catch {
	Write-Log 'There was a problem entering identification to web site.' 'ERROR'
}
# Wait for page to load
sleep -milliseconds 4000


#------------------------------  Extract Info  ------------------------------------
Write-Log 'Extracting information...' 'INFO'
$elements = $ie.Document.getElementsByClassName('w3-container w3-padding w3-pale-green')
if ($elements.length -eq 0) {
	$elements = $ie.Document.getElementsByClassName('w3-container w3-padding w3-pale-red')
}
if ($elements.length -eq 0) {
	Write-Log 'There was a problem extracting information from web site.' 'ERROR'
	$mustTest = 'There was a problem retrieving data from the site.  Please call testing number.'
	$initials = 'NULL'
	$confNumr = 'NULL'
	$dateTime = $timeStart.GetDateTimeFormats()[53]
}
else {
	$mustTest = $elements[0].outerText
	$elements = $ie.Document.getElementsByClassName('w3-right-align w3-col s6')
	$initials = $elements[0].outerText
	$elements = $ie.Document.getElementsByClassName('w3-right-align w3-col s3')
	$confNumr = $elements[0].outerText
	$elements = $ie.Document.getElementsByClassName('w3-right-align w3-col s9')
	$dateTime = $elements[0].outerText
}


#------------------------------  Send Message  -----------------------------------
Write-Log 'Constructing message...' 'INFO'

Write-Log "Message:
`t$mustTest
`tInitials:$($initials.PadLeft(20))
`tConfirmation:$($confNumr.PadLeft(16))
`tDate&Time:$($dateTime.PadLeft(19))" 'INFO'

$textMsg = "$mustTest
Initials: $initials
Conf-Num: $confNumr
Time: $dateTime"

Write-Log 'Creating credentials...' 'INFO'
$username = $emailAddr
$securePwd = $password | ConvertTo-SecureString -AsPlainText -Force
$emailCred = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securePwd
Write-Log "Sending email through $emailAddr..." 'INFO'
try {
	$properties = @{
		smtpserver = 'smtp.gmail.com'
		port = 587
		subject = 'Daily Testing Message'
		body = $textMsg
		from = $emailAddr
		to = ''
	}
	foreach($phone in $phoneNums) {
		$properties.to = "$($phone.number)@$($SMSTable[$phone.carrier])"
		Send-MailMessage @properties -UseSsl -Credential($emailCred) -ErrorAction Stop
		sleep -milliseconds 10000
		Write-Log "Message sent to $($phone.number)." 'INFO'
	}
}
catch {
	Write-Log 'There was a problem accessing email server.' 'FATAL'
	Write-Log '(SCRIPT FAILED)' 'FATAL'
	exit 1 # Exit with error code
}
$ie.Quit()
Write-Log '(SCRIPT FINISHED)' 'INFO'

#Read-Host -Prompt "Press Enter to exit"  # Uncomment to have script prompt for input after running
exit 0
