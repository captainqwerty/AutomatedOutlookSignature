# Automated Outlook Signature Script
This project contains two scripts. One PowerShell script is used to automate the creation of Outlook signatures using user information from Active Directory and also sets this signature as the user's default signature for new emails and email replies and the other can be used to set users Outlook Web Signatue.

Outlook desktop signature script currently tested on and working with Outlook 2010, 2016 and 2019.
Outlook web signature script has been tested on Exchange Online.

Here is a tutorial video to assist with the setup of the Outlook desktop signature script: https://www.youtube.com/watch?v=rt9y02iBoPE

### Active Directory
The signatures are dynamic and when the users job title or their maiden name etc. are updated in Active Directory their Outlook signature will also be updated!  A selection of Active Directory attribute are already configured in the script and listed below however more attributes can be easily added. 

The follow details are used from Active Directory within the script:

| Variable in Script | AD Field  | Notes | Optional |
|-------------| ------------- | ------------- | ------------- |
| $displayName | Display name | Users display name | Yes |
| $jobTitle | Job title | Users job title | Yes |
| $email | Email | Users email address  | Yes |
| $telephone | Telephone  | The main site/branch telephone number | Yes |
| $directDial | Home | The users direct dial number | Yes |
| $mobileNumber | Mobile | The users mobile number | Yes |
| $street | Street | Street / First line of address | Yes |
| $poBox | P.O. Box | Site / Branch name which will appear in bold above the address e.g. Head Office | Yes |
| $city | City | City / Town | Yes |
| $state | State/Province | State / County | Yes |
| $zipCode | Zip/Postal Code | Post Code / Zip Code | Yes |
| $office | physicaldeliveryofficename | Office | Yes |
| $website | Website | Website address | Yes |

Additional variables that do not rely on Active Directory

| Variable in Script | Usage | Optional |
|-------------| ------------- | ------------- |
| $companyName | Variable containing the name of the company | Yes |
| $logo | Variable containing the URL of a image to use as a logo in the signature | Yes |

### How to use this script
I recommend using the script in Group Policy as a log-on script.  If you are unaware of how to do this rather than reinvent the wheel explain here I shall point you to this article :) - [Configuring Logon PowerShell Scripts with Group Policy - 4Sysops](https://4sysops.com/archives/configuring-logon-powershell-scripts-with-group-policy/)

### Need help?
If you require help with the script or would like assistance altering it more for your own environment please see my EduGeek thread on this script and feel free to comment on the thread or PM on EduGeek. You could also leave a comment on the tutorial video if you like.

[EduGeek Post](http://www.edugeek.net/forums/scripts/205976-outlook-email-signature-automation-ad-attributes.html#post1760284)
