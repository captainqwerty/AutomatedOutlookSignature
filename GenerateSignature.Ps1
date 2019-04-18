# Antony Bragg
# Updated to add Qual letters at end of name....
# Create the signatures folder in %appdata%
$folderlocation = $Env:appdata + '\\Microsoft\\signatures'  
mkdir $folderlocation -force

# Setting the file name and location
$Filename  = "$folderLocation\\signature.htm"

# Getting Active Directory information for current user
$UserName = $env:username

$Filter = "(&(objectCategory=User)(samAccountName=$UserName))" 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.Filter = $Filter 
$ADUserPath = $Searcher.FindOne() 
$ADUser = $ADUserPath.GetDirectoryEntry()

# Set properties
$displayName = $ADUser.DisplayName.Value
$title = $ADUser.title.Value
$eduLetters = $ADuser.info.Value
$poBox = $ADUser.postOfficeBox.Value
$street = $ADUser.streetaddress.Value
$city = $ADUser.l.Value
$state = $aduser.st.Value
$zipCode = $ADUser.postalCode.Value
$telephone = $ADUser.TelephoneNumber.Value
$homePhoneNumber = $ADUser.homePhone.Value
$mobileNumber = $ADUser.mobile.Value
$email = $ADUser.mail.Value

# Build the HTML file
$signature = 
@"
<div style="color: #470a68; font-family: Arial, Helvetica, sans-serif; font-size: 12"><b>{0} {1}</b><br>
{2}</div><br>

<a href="http://www.gen2.ac.uk"><img src='https://www.gen2.ac.uk/images/logo_with_city_and_guilds.png' /></a><br><br>

<div style="font-family: Arial, Helvetica, sans-serif; font-size: 12; color: #000"><b>{3}</b><br>
{4},<br>
{5}, {6}, {7}.<br><br>

<p><table border="0" style="font-family: Arial, Helvetica, sans-serif; font-size: 12; color: #000">
    <tr>
        <td>
            t:
        </td>
        <td>
            {8}
        </td>
    </tr>
"@ -f $displayname, $eduLetters, $title, $poBox, $street, $city, $state, $zipCode, $telephone

# Direct Dial
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
            <a href="http://www.gen2.ac.uk" style="color: #470a68; font-family: Arial, Helvetica, sans-serif; font-size: 12"><b>www.gen2.ac.uk</b></a>
        </td>
    </tr>
</table></p><br>

"@ -f $email

$signature | out-file $Filename -encoding ascii

# Setting the regkeys for Outlook 2013
if (test-path "HKCU:\\Software\\Microsoft\\Office\\15.0\\Common\\General") 
{
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value signature -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value signature -propertytype string -force
    Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook\\Setup -Name "First-Run"
}

# Setting the regkeys for Outlook 2016
if (test-path "HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General") 
{
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value signature -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value signature -propertytype string -force
    Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook\\Setup -Name "First-Run"
}