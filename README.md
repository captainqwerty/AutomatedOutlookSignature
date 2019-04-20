# Autoamted Outlook Signature Script
This PowerShell script can be used to automate the creation of Outlook signatures using Active Directory, it will also set this signature as the users default signature for new emails and email replies.  Currently tested on and working with Outlook 2016 and 2019.

### Active Directory
To ensure users signatures are dynamic and when their job title changes, or their maiden name etc. are updated in Active Directory their Outlook signature will also be updated!

The follow details are used from Active Directory:

| Variable in Script | AD Field  | Notes |
|-------------| ------------- | ------------- |
| $displayName | Display name | Users display name |
| $jobTitle | Job title | Users job title |
| $email | Email | Users email address  |
| $telephone | Telephone  | The main reception telephone number not a direct dial |
| $directDial | Home | The users direct dial number (can be blank) |
| $mobileNumber | Mobile | The users mobile number (can be blank) |
| $street | Street | Street address of users site |
| $poBox | P.O. Box | Site name which will appear in bold above the address e.g. Gen2 Head Office |
| $city | City | The site city/town |
| $state | State/Province | The site county | 
| $zipCode | Zip/Postal Code | Post Code |
| $website | Website | Website address |

Additional variables that do not rely on Active Directory

| Variable in Script | Usage |
|-------------| ------------- |
| $logo | Variable containing the URL of a image to use as a logo in the signature |

### Why a Script and not Mail-Flow?
The reason I decided to use a script rather than Mail-Flow rules in Exchange is the undesired affects that occurred when using Mail-Flow rules: our users didn’t like the fact they couldn’t see their signature while writing emails; it was near impossible to get the signature to be applied to reply emails without just being pinned to the bottom of the thread and; it would repeat the signature at the bottom of the email over and over again and if you used a rule to detect it and not put it on again you only got the one signature at the very bottom of the message thread and not in replies. 
