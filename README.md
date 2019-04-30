# Automated Outlook Signature Script
This PowerShell script can be used to automate the creation of Outlook signatures using Active Directory, it will also set this signature as the users default signature for new emails and email replies.  Currently tested on and working with Outlook 2010, 2016 and 2019.

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
| $website | Website | Website address | Yes* |

Additional variables that do not rely on Active Directory

| Variable in Script | Usage | Optional |
|-------------| ------------- | ------------- |
| $logo | Variable containing the URL of a image to use as a logo in the signature | Yes* |

\* If either the logo or the website variable are blank currently this will stop the logo showing.  The next version will not have this restriction.

### How to use this script
I recommend using the script in Group Policy as a log-on script.  If you are unaware of how to do this rather than reinvent the wheel explain here I shall point you to this article :) - [Configuring Logon PowerShell Scripts with Group Policy - 4Sysops](https://4sysops.com/archives/configuring-logon-powershell-scripts-with-group-policy/)

### Need help?
If you require help with the script or would like assistance altering it more for your own environment please see my EduGeek thread on this script and feel free to comment on the thread or PM on EduGeek.

[EduGeek Post](http://www.edugeek.net/forums/scripts/205976-outlook-email-signature-automation-ad-attributes.html#post1760284)

### The Script

The first part of the script ensures the Microsoft Signatures folder exists in the users %appdata% and sets the directory path and file name of the signature the script creates. 

```powershell
$folderlocation = $Env:appdata + '\\Microsoft\\signatures'

if(!(Test-Path -Path $folderlocation )){
    New-Item -ItemType directory -Path $folderlocation
}
```

Next we get the users username from their session and use the DirectorySearcher to search Active Directory for the user and store their account detials in the $ADUser variable.

```powershell
# Getting Active Directory information for current user
$UserName = $env:username
$user = (([adsisearcher]"(&(objectCategory=User)(samaccountname=$UserName))").FindOne().Properties)
```

The attributes from the $ADUser object are split into two sections.  The first section are the details which should be uniquie and would not be the same as any other member of staff such as their email address and the second half are attributes which could be the same for every member of staff such as the address details, these details can be either taken from active directory or could be set statically within the script. 

```powershell
# Get the users properties (These should always be in Active Directory and Unique)
$displayName = $user.name[0]
$jobTitle = $user.title[0]
$directDial = $user.homephone[0]
$mobileNumber = $user.mobile[0]
$email = $user.mail[0]

# These are details you can either get from Active directory or as they might be the same for your entire company could statically set them here. Each has a commented out static example, simply swap the commented lines and alter the example.
$poBox = $user.postofficebox[0]
$street = $user.streetaddress[0]
$city = $user.l[0]
$state = $user.st[0]
$zipCode = $user.postalcode[0]
$telephone = $user.telephonenumber[0]
$website = $user.wwwhomepage[0] # Ensure that if you use this that the website is lead with http:// or https:// or the hyperlinks wont work
$logo = "https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_92x30dp.png"
```

The next part of the script gets fairly wordy.  This part is where the HTML file is built using several If statements.  This is done so that if you do not utilise a particular AD field or not all users have an entry (such as direct dial) it is not added to the script but the resulting signature still looks presentable.  The line breaks before before "@ in some sections are not mandatory I only use them so the HTML file that is generated is more readable.

```powershell
$signature = 
@"
<div style="color: #000; font-family: Arial, Helvetica, sans-serif; font-size: 12">
    <p>
        <b>$displayName</b><br>
        $jobTitle
    </p>
    <p>
        <a href="$website"><img src='$logo' /></a> 
    </p>

    <p>
        $(if($poBox){"<b>"+$poBox+"</b><br>"})
        $(if($street){ $street+",<br>"})
        $(if($city){ $city+",<br>"})
        $(if($state){ $state+",<br>"})
        $(if($state){ $zipCode+"<br>"})
    </p>

    <p>
        <table border="0" style="font-family: Arial, Helvetica, sans-serif; font-size: 12; color: #000">
            $(if($telephone){"<tr><td>t: </td><td>"+$telephone+"</td></tr>"})
            $(if($directDial){"<tr><td>dd: </td><td>"+$directDial+"</td></tr>"})
            $(if($mobileNumber){"<tr><td>m: </td><td>"+$mobileNumber+"</td></tr>"})
            $(if($email){"<tr><td>w: </td><td><a href='mailto:$email'>"+$email+"</td></tr>"})
            $(if($website){"<tr><td>e: </td><td><a href='$website'><b>"+$website+"</b></td></tr>"})
        </table>
    </p>
</div>
"@

```

Now the $signature variable is populated with the HTML required it is output to the directory created earlier with the filename specified. 

```powershell
# Save the HTML to the signature file
$Filename  = "$folderLocation\\signature.htm"
$signature | out-file $Filename -encoding ascii
```

Finally the required registry keys are created and one is removed.  This helps ensure that if the user creates their own signature, each time the script is ran the automated signature will be used.

```powershell
# Setting the regkeys for Outlook 2016
if (test-path "HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General") 
{
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value signature -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value signature -propertytype string -force
    Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook\\Setup -Name "First-Run"
}

# Setting the regkeys for Outlook 2010 - Thank you AJWhite1970 for the 2010 registry keys
if (test-path "HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General") 
{
    get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ MailSettings | new-Itemproperty -name NewSignature -value signature -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ MailSettings | new-Itemproperty -name ReplySignature -value signature -propertytype string -force
    Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\14.0\\Outlook\ \Setup -Name "First-Run"
}
```
