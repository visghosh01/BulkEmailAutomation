clear
#$SUBJECT = $null

$Date = Get-Date -Format "dd/MM/yyyy"

$ROLE = $null
$ATTACHMENT = "C:\Users\djmay\Downloads\Mayur Chavan Resume Linkdin.pdf"
$a = import-csv -Path C:\Users\djmay\OneDrive\Desktop\EmailAutomation\EmailAutomateDatabase.csv -Header Mail,Name,Role,Company,Portal,MyName,MyEmailid -Delimiter `t

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

$SUBJECT = "Regarding : Career Opportunity for "+$ROLE+ " - "+$COMPANY+ " | " + $Date

$BODY_1 = "Hello " + $Name.split(' ')[0]
$BODY_2 = "`n`nMy name is "+$MYNAME+" and from "+$PORTAL+" I found that you are actively recruiting "+$ROLE+" for " + $COMPANY +"."
$BODY_3 = "`n`nI have been working as a System Engineer/DevOps Engineer with Infosys Limited for the last 4.6 years and in that time I've had a great hands-on experience on AWS, DevOps(Terraform, Ansible and CircleCI) as well as basic knowledge on Docker and Kubernetes. `
`nIf you have any opportunities available for DevOps Engineer/Cloud Infrastructure Engineer/Cloud Automation Engineer then I would greatly appreciate to talk further about role, work culture and how we may work together. `
`nPlease take your time to review my attached resume. I believe that I would be an excellent candidate for the available position, and I'm waiting for the opportunity to meet you in person and discuss how my skills and experience can benefit $COMPANY .`
`nThank you for the opportunity. I look forward to hearing from you.
`nThanks and Regards,
`Mayur Chavan"

$BODY = $BODY_1 + $BODY_2 + $BODY_3
$BODY
#Send-MailMessage -From $FROM -Credential ($creds) -Port 587 -UseSsl -To $TO -Subject $SUBJECT -Body $BODY -Attachments $ATTACHMENT -SmtpServer 'smtp.gmail.com'
}

# -Priority High `
# -DeliveryNotificationOption OnSuccess, OnFailure `
# -SmtpServer 'smtp-mail.outlook.com `
