# GitHub: https://github.com/captainqwerty/AutomatedOutlookSignature 
# Author: CaptainQwerty
# Last Modified: 19/04/2019
# Notes: Please refer to the ReadMe in the GitHub Repository for a detailed break down of this script.

# Create the signatures folder and signature file
$folderlocation = $Env:appdata + '\\Microsoft\\signatures'  
mkdir $folderlocation -force
$Filename  = "$folderLocation\\signature.htm"

# Getting Active Directory information for current user
$UserName = $env:username
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))" 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.Filter = $Filter 
$ADUserPath = $Searcher.FindOne() 
$ADUser = $ADUserPath.GetDirectoryEntry()

# Get the users properties (These should be unique per user and should not be static within this script although some are optional entirely) 
$displayName = $ADUser.DisplayName.Value
$jobTitle = $ADUser.title.Value
$directDial = $ADUser.homePhone.Value
$mobileNumber = $ADUser.mobile.Value
$email = $ADUser.mail.Value 

# These are details you can either get from Active directory or as they might be the same for your entire company could statically set them here. Each has a commented out static example, simply swap the commented lines and alter the example.
$poBox = $ADUser.postOfficeBox.Value 
$street = $ADUser.streetaddress.Value 
$city = $ADUser.l.Value 
$state = $aduser.st.Value 
$zipCode = $ADUser.postalCode.Value
$telephone = $ADUser.TelephoneNumber.Value
$website = $ADUser.wWWHomePage.Value # Ensure that if you use this that the website is lead with http:// or https:// or the hyperlinks wont work
$logo = "https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_92x30dp.png"

# Build the HTML 
$signature = 
@"
<div style="color: #000; font-family: Arial, Helvetica, sans-serif; font-size: 12">
<p>
    <b>{0}</b><br>
    {1}
</p>
"@ -f $displayname, $jobTitle

if($website -and $logo)
{
$signature = $signature + 
@"
<p>
    <a href="{0}"><img src='{1}' /></a> 
</p>

"@ -f $website, $logo
}

if($poBox)
{
$signature = $signature + 
@"

<b>{0}</b><br>
"@ -f $poBox
}

if($street)
{
$signature = $signature + 
@"

{0},<br>
"@ -f $street
}

if($city)
{
$signature = $signature + 
@"

{0},<br>
"@ -f $city
}

if($state)
{
$signature = $signature + 
@"

{0},<br>
"@ -f $state
}

if($zipCode)
{
$signature = $signature + 
@"

{0}<br>
"@ -f $zipCode
}

$signature = $signature + 
@"

<p><table border="0" style="font-family: Arial, Helvetica, sans-serif; font-size: 12; color: #000">
"@ 

if($telephone)
{
$signature = $signature + 
@"
    <tr>
        <td>
            t:
        </td>
        <td>
            {0}
        </td>
    </tr>
"@ -f $telephone
} 

# If homePhoneNumber is not blank it will be added (we use this field for our users Direct Dial numbers)
if($directDial)
{
$signature = $signature + 
@"
    <tr>
        <td>
            dd:
        </td>
        <td>
            {0}
        </td>
    </tr>
"@ -f $directDial
} 

# If mobilenumber is not blank it will be added
if($mobileNumber)
{
    $signature = $signature + 
@"
    <tr>
        <td>
            m:
        </td>
        <td>
            {0}
        </td>
    </tr>
"@ -f $mobileNumber
}

# If email is not blank it will be added
if($email)
{
    $signature = $signature + 
@"
    <tr>
        <td>
            e:
        </td>
        <td>
            <a href="mailto:{0}">{0}</a>
        </td>
    </tr>
"@ -f $email
}

# if the website is not blank it will be added
if($website)
{
$signature = $signature +
@"
    <tr>
        <td>
            w:
        </td>
        <td>
            <a href="{0}" style="color: #470a68; font-family: Arial, Helvetica, sans-serif; font-size: 12"><b>{0}</b></a>
        </td>
    </tr>
"@ -f $website
}

# Ends the table
$signature = $signature +
@"

</table>
</p>
</div>
"@

# Save the HTML to the signature file
$signature | out-file $Filename -encoding ascii

# Setting the regkeys for Outlook 2016
if (test-path "HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General") 
{
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value signature -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value signature -propertytype string -force
    Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook\\Setup -Name "First-Run"
}
