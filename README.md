# SCCM_Collection_Notification

## Summary

This script send mail to members of the collection of SCCM. We can personalize mail content and update files before send the mail. 

Usage example : 
- Update notification
- Security alert
- Information

---

### :warning: Prerequisites
Please be carefull to have installed : 
- Notepad++ (for editing html files)
- Microsoft Edge (for open & check html file)
- SCCM Console installed
---

## :dart: Process
The script do : 
- Create folder
- Copy template files
- Open html file with Notepad++ for editing your email content
- Check your file with Microsoft Edge
- Get members of SCCM collection
- Check in Active Directory members and email
- Send email for each member of collection
- Export members & infos in csv file

---
### :hammer: 2. Set variables
#### Edit script
You must editing some variables before execute the script :
- **AD_Server** -> Set your Active Directory server
- **SmtpServer** -> Add your SMTP server
- **EmailSender** -> Add the email sender

#### Parameters
Here is the parameters of the script: 
- **TemplateName**: Name of your template
- **CollectionName**: Name of the SCCM collection target
- **TitleMail**: The title of your mail
- **MailType**: The type of mail you want (possible value : Information, Warning, Alert)

In the body template, we have two variable : 
- **XXXUSERNAMEXXX** :  Get the GivenName of user
- **YYYCOMPUTERNAMEYYY** : Get computer name of user 

**This two variables are not mandatory !** If you don't want in your mail template, you can just delete variables.

---

### :rocket: 3. Execute script
For execute the script, go to the folder containing the script and execute like below :

>.\SCCM_Collec_Notification.ps1 -TemplateName "Your_Template_Name" -CollectionName "Collection_Sccm_1" -TitleMail "Test Email" -MailType "Information"


For example, an email to the computer with Windows updates late :
>.\SCCM_Collec_Notification.ps1 -TemplateName "Update_Windows10_Late" -CollectionName "Windows10_Update_Late" -TitleMail "Windows Update Late" -MailType "Alert"

After, the script open Notepad++ automatly and you will could add your content in the mail body. 