# Automated Outlook Signature Scripts
This project contains two scripts: 
* GenerateSignature.ps1 - Used to generate and set a user's signature for desktop Outlook
* Set-OutlookWebSignatures.ps1 - Used to gen

Outlook desktop signature script currently tested on and working with Outlook 2010, 2016 and 2019.
Outlook web signature script has been tested on Exchange Online.

## How To Use The Scripts
This is a very basic description on how to use the scripts. For more detail please see the YouTube videos linked earlier 

### Desktop Signature Script
Video guide: https://www.youtube.com/watch?v=rt9y02iBoPE

I recommend using the script in Group Policy as a log-on script.  If you are unaware of how to do this rather than reinvent the wheel explain here I shall point you to this article :) - [Configuring Logon PowerShell Scripts with Group Policy - 4Sysops](https://4sysops.com/archives/configuring-logon-powershell-scripts-with-group-policy/)

### Exchange Online Signature Script
Video guide: Coming soon
Currently this script can be ran manually from a device which has both the ActiveDirectory module and the ExchangeOnlineManagement module.

### Need further help?
If you require help with the script or would like assistance altering it more for your own environment please see my EduGeek thread on this script and feel free to comment on the thread or PM on EduGeek. You could also leave a comment on the tutorial video if you like.

[EduGeek Post](http://www.edugeek.net/forums/scripts/205976-outlook-email-signature-automation-ad-attributes.html#post1760284)

### Active Directory
A selection of Active Directory attribute are already configured in the script and listed below however more attributes can be easily added. 

The following properties are used from Active Directory within the script:

| Variable in Script | AD Field  | Notes |
|-------------| ------------- | ------------- |
| $displayName | Display name | Users display name |
| $jobTitle | Job title | Users job title |
| $email | Email | Users email address  |
| $telephone | Telephone  | The main site/branch telephone number |
| $directDial | Home | The users direct dial number |
| $mobileNumber | Mobile | The users mobile number |
| $street | Street | Street / First line of address |
| $poBox | P.O. Box | Site / Branch name which will appear in bold above the address e.g. Head Office |
| $city | City | City / Town |
| $state | State/Province | State / County |
| $zipCode | Zip/Postal Code | Post Code / Zip Code |
| $office | physicaldeliveryofficename | Office |
| $website | Website | Website address |

Additional variables that do not rely on Active Directory

| Variable in Script | Usage |
|-------------| ------------- |
| $companyName | Variable containing the name of the company |
| $logo | Variable containing the URL of a image to use as a logo in the signature |

