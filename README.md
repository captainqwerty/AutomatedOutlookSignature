# AutomatedOutlookSignature
PowerShell script to automate the creation of Outlook signatures using Active Directory.

The follow details should be filled out in active directory:

| AD Field  | Contents |
| ------------- | ------------- |
| Display name | Users display name |
| Job title | Users job title |
| Email | Users email address  |
| Telephone  | The main reception telephone number not a direct dial |
| Home | The users direct dial number (can be blank) |
| Pager | The users extension (can be blank) |
| Mobile | The users mobile number (can be blank) |
| ExtensionAttribute1 | Any prefix such as Dr (can be blank) |
| ExtensionAttribute2 | Any qualification letters that need to follow their name (can be blank) |
| Street | Street address of users site |
| P.O. Box | Site name which will appear in bold above the address e.g. Gen2 Head Office |
| City | The site city/town |
| State/Province | The site county | 
| Zip/Postal Code | Post Code |


### Why a Script and not Mail-Flow?
The reason I decided to use a script rather than Mail-Flow rules in Exchange is the undesired affects that occurred when using Mail-Flow rules: our users didn’t like the fact they couldn’t see their signature while writing emails; it was near impossible to get it to back on replies and; it would repeat the signature at the bottom of the email over and over again and if you used a rule to detect it and not put it on again you only got the one signature at the very bottom of the message thread and not in replies. 
