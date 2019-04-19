# GitHub: https://github.com/captainqwerty/AutomatedOutlookSignature 
# Author: CaptainQwerty
# Last Modified: 19/04/2019

# Create the signatures folder and signature file - Do not alter this.
$folderlocation = $Env:appdata + '\\Microsoft\\signatures'  
mkdir $folderlocation -force
$Filename  = "$folderLocation\\signature.htm"

# Getting Active Directory information for current user - Do not alter this.
$UserName = $env:username
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))" 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.Filter = $Filter 
$ADUserPath = $Searcher.FindOne() 
$ADUser = $ADUserPath.GetDirectoryEntry()

# Get the properties
$displayName = $ADUser.DisplayName.Value # Display name which will likely be your users full name e.g. Barry Banderos
$title = $ADUser.title.Value # This is their job title e.g. Finance Manager
$poBox = $ADUser.postOfficeBox.Value # We use the PoBox value for our site names such as "Head Quarters" which is then displayed above the address in bold 
$street = $ADUser.streetaddress.Value # Street address 
$city = $ADUser.l.Value # City address
$state = $aduser.st.Value # As we are in the UK we use this for county 
$zipCode = $ADUser.postalCode.Value # Post code
$telephone = $ADUser.TelephoneNumber.Value # We use this as our reception phone number
$homePhoneNumber = $ADUser.homePhone.Value # We use this as the users direct dial (which can be blank and is handled later in the script if blank)
$mobileNumber = $ADUser.mobile.Value # Mobile number for the user if they have one, again can be left blank.
$exAttribute1 = $ADUser.extensionAttribute1.Value # Used if they have a prefix such as Dr before their name. Can be blank.
$exAttribute2 = $ADUser.extensionAttribute2.Value # Used if they want to display qualification letters such as BSch(Hons) after their name. Can be blank.
$email = $ADUser.mail.Value # Email address of the user
$website = $ADUser.wWWHomePage.Value # or can comment this out if you do not use the web page field in AD and can use the below line to set the website statically
#$website = "https://www.google.com"
$logo = "https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_92x30dp.png" # A public image file such as a PNG of the company logo in the correct dimensions

# Build the HTML 
$signature = 
@"
<div style="color: #470a68; font-family: Arial, Helvetica, sans-serif; font-size: 12"><b>{0} {1} {2}</b><br>
{3}</div><br>

<a href="{4}"><img src='{5}' /></a><br><br>


<div style="font-family: Arial, Helvetica, sans-serif; font-size: 12; color: #000"><b>{4}</b><br>
{6},<br>
{7},<br>
{8},<br> 
{9}<br><br>

<p><table border="0" style="font-family: Arial, Helvetica, sans-serif; font-size: 12; color: #000">
    <tr>
        <td>
            t:
        </td>
        <td>
            {10}
        </td>
    </tr>
"@ -f $exAttribute1, $displayname, $exAttribute2, $title, $website, $logo, $poBox, $street, $city, $state, $zipCode, $telephone

# If homePhoneNumber is not blank it will be added (we use this field for our users Direct Dial numbers)
if($homePhoneNumber)
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
"@ -f $homePhoneNumber
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

# Finally the users email address and a link to your website is added to the bottom of the footer.
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
    <tr>
        <td>
            w:
        </td>
        <td>
            <a href="{1}" style="color: #470a68; font-family: Arial, Helvetica, sans-serif; font-size: 12"><b>{1}</b></a>
        </td>
    </tr>
</table></p><br>

"@ -f $email, $website

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
