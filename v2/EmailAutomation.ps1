clear
#$SUBJECT = $null

$Date = Get-Date -Format "dd/MM/yyyy"
$refer_count = 0
$recruit = 0
$ATTACHMENT = "C:\Users\djmay\Downloads\Mayur Chavan Resume Linkdin.pdf"
$a = import-csv -Path C:\Users\djmay\OneDrive\Desktop\EmailAutomation\v2\EmailAutomateDatabase.csv -Header Mode,Mail,Name,Role,Company,Portal,MyName,MyEmailid,CurrentRole,CurrentCompany,CurrentYOE,CurrentSkills,NewRole,CareerUrl,LinkdinUrl -Delimiter `t

$creds = New-Object   System.Management.Automation.PsCredential(get-credential)

foreach($b in $a)
{
$Name = $b.Name
$FROM = $b.MyEmailid
$TO = $b.Mail
$ROLE = $b.Role
$PORTAL = $b.Portal
$COMPANY = $b.Company
$MYNAME = $b.MyName
$CurrentRole = $b.CurrentRole
$CurrentCompany = $b.CurrentCompany
$CurrentYOE = $b.CurrentYOE
$CurrentSkills = $b.CurrentSkills
$NewRole = $b.NewRole
$Mode = $b.Mode
$CareerUrl = $b.CareerUrl
$Careerlink = "<a href=$CareerUrl>Careers</a>"
$LinkdinUrl = $b.LinkdinUrl
$link = "<a href=$LinkdinUrl>Linkedin</a>"

if ( $Mode -eq "Recruitment" ) {

$SUBJECT = "Regarding : Career Opportunity for "+$ROLE+ " - "+$COMPANY+ " | " + $Date

$BODY_1 = "Hello " + $Name.split(' ')[0]
$BODY_2 = "<br><br>My name is "+$MYNAME+" and from "+$PORTAL+" I found that you are actively recruiting "+$ROLE+" for " + $COMPANY +"."
$BODY_3 = "<br><br>I have been working as a "+$CurrentRole+" with "+$CurrentCompany+" for the last "+$CurrentYOE+" years and in that time I've had great hands-on experience on "+$CurrentSkills+"`
<br><br>If you have any opportunities available for "+$NewRole+", I would greatly appreciate discussing further about role, work culture and how we may work together. `
<br><br>Please take your time to review my attached resume. I believe that I would be an excellent candidate for the available position and I'm waiting for the opportunity to meet you in person and discuss how my skills and experience can benefit $COMPANY .`
<br><br>If you would like to know better, you can reach out to me on $link.
<br><br>Thank you for the opportunity. I look forward to hearing from you.
<br><br>Thanks and Regards,
<br>$MYNAME"

$recruit = $recruit + 1

$BODY = $BODY_1 + $BODY_2 + $BODY_3
$SUBJECT
$BODY
}

elseif ( $Mode -eq "Referral" ) {

$SUBJECT = "Looking for a referral in "+$Company+" - "+$Role + " | " + $Date

$refer_1 = "Hi "+$Name.split(' ')[0]
$refer_2 = "<br><br>I hope this message finds you well! I've been with "+$CurrentCompany+" for "+$CurrentYOE+" years as a "+$CurrentRole+", and I've just started looking for a new opportunity. I've been following "+$COMPANY+" for some time and there are a few roles I've come across that seem like a great fit."

$refer_3 = "<br><br>I'd be so grateful if you could help me with referring to the available position in "+$COMPANY

$refer_4 = "<br><br>Please find the career link that matches my role - $Careerlink.

<br><br>If you would like to know better, you can reach out to me on $link.

<br><br>Appreciate your help !

<br><br>Thanks and Regards,
<br>$MYNAME"

$refer_count = $refer_count + 1

$BODY = $refer_1 + $refer_2 + $refer_3 + $refer_4
$SUBJECT
$BODY
}

else {
write-host("We are unable to read your mode for `""+$COMPANY+"/"+$ROLE+"`". Can you please help us in defining proper mode in our database") -ForegroundColor red
}

#Send-MailMessage -From $FROM -Credential ($creds) -Port 587 -UseSsl -To $TO -Subject $SUBJECT -Body $BODY -BodyAsHtml -Attachments $ATTACHMENT -SmtpServer 'smtp.gmail.com' -DeliveryNotificationOption OnSuccess, OnFailure
}

write-host("I have sent "+$refer_count+ " refferal emails and "+$recruit+" cold emails.") -ForegroundColor green

# -Priority High `
# -DeliveryNotificationOption OnSuccess, OnFailure `
# -SmtpServer 'smtp-mail.outlook.com `