# Autoamted Outlook Signature Script
This PowerShell script can be used to automate the creation of Outlook signatures using Active Directory, it will also set this signature as the users default signature for new emails and email replies.  Currently tested on and working with Outlook 2016 and 2019.

### Active Directory
To ensure users signatures are dynamic and when their job title changes, or their maiden name etc. are updated in Active Directory their Outlook signature will also be updated!

The follow details are used from Active Directory:

| Variable in Script | AD Field  | Notes | Optional |
|-------------| ------------- | ------------- | ------------- |
| $displayName | Display name | Users display name | No |
| $jobTitle | Job title | Users job title | No |
| $email | Email | Users email address  | Yes |
| $telephone | Telephone  | The main site/branch telephone number | Yes |
| $directDial | Home | The users direct dial number | Yes |
| $mobileNumber | Mobile | The users mobile number | Yes |
| $street | Street | Street / First line of address | Yes |
| $poBox | P.O. Box | Site / Branch name which will appear in bold above the address e.g. Head Office | Yes |
| $city | City | City / Town | Yes |
| $state | State/Province | State / County | Yes |
| $zipCode | Zip/Postal Code | Post Code / Zip Code | Yes |
| $website | Website | Website address | Yes |

Additional variables that do not rely on Active Directory

| Variable in Script | Usage |
|-------------| ------------- |
| $logo | Variable containing the URL of a image to use as a logo in the signature |

### How to use this script
I recommend using the script in Group Policy as a log on script.  If you are unaware of how to do this rather than reinvent the wheel explain here I shall point you to this article :) - [Configuring Logon PowerShell Scripts with Group Policy - 4Sysops](https://4sysops.com/archives/configuring-logon-powershell-scripts-with-group-policy/)

### Why a Script and not Mail-Flow?
The reason I decided to use a script rather than Mail-Flow rules in Exchange is the undesired affects that occurred when using Mail-Flow rules: our users didn’t like the fact they couldn’t see their signature while writing emails; it was near impossible to get the signature to be applied to reply emails without just being pinned to the bottom of the thread and; it would repeat the signature at the bottom of the email over and over again and if you used a rule to detect it and not put it on again you only got the one signature at the very bottom of the message thread and not in replies. 
