<#
.SYNOPSIS
  Dynamic Outlook Signature Creation - https://github.com/captainqwerty/AutomatedOutlookSignature
.DESCRIPTION
  This script uses properties from ActiveDirectory to populate the HTM which is then stored in the $folderlocation with the name $filename.htm. 
  The script can be ran as either a scheduled task at logon or preferably a Group Policy Logon script, more details on this can be found in the GitHub ReadMe.
  The script uses if statements to ensure if a user, for example, doesn't have a mobile number then that entire section of the signature is not uncluded, this makes the signature look much more professional and 
  if the user is given a mobile number, at next logon it will be added to their signature. 
  The script can also add parts based on group membership. 
.INPUTS
  Several properties for the user are taken directly from Active Directory, for more details on this please see the ReadMe on GitHub.
.OUTPUTS
  $folderlocation\$filename.htm
.NOTES
  Version:        2.1
  Author:         CaptainQwerty
  Modified:       22/09/2019
#>

# Create the signatures folder
$folderlocation = $Env:appdata + '\\Microsoft\\signatures'
$filename = "signature"

if(!(Test-Path -Path $folderlocation )){
    New-Item -ItemType directory -Path $folderlocation
}

# Getting Active Directory information for current user
$UserName = $env:username
$user = (([adsisearcher]"(&(objectCategory=User)(samaccountname=$UserName))").FindOne().Properties)

# Get the users properties (These should always be in Active Directory and Unique)
if($user.name.count -gt 0){$displayName = $user.name[0]}
if($user.title.count -gt 0){$jobTitle = $user.title[0]}
if($user.homephone.count -gt 0){$directDial = $user.homephone[0]}
if($user.mobile.count -gt 0){$mobileNumber = $user.mobile[0]}
if($user.mail.count -gt 0){$email = $user.mail[0]}

<# 
These are details you can either get from Active directory or as they might be the same for your entire company could statically set them here. 
#>
if($user.postofficebox.count -gt 0){$poBox = $user.postofficebox[0]}
if($user.streetaddress.count -gt 0){$street = $user.streetaddress[0]}
if($user.l.count -gt 0){$city = $user.l[0]}
if($user.st.count -gt 0){$state = $user.st[0]}
if($user.postalcode.count -gt 0){$zipCode = $user.postalcode[0]}
if($user.telephonenumber.count -gt 0){$telephone = $user.telephonenumber[0]}
if($user.physicaldeliveryofficename.count -gt 0){$office = $user.physicaldeliveryofficename}
if($user.wwwhomepage.count -gt 0){$website = $user.wwwhomepage[0]} # Must include https:// or http:// or the hyperlinks will not function.

# Please note that if you do include a logo it must be located somewhere on the internet that the public has access to, many users upload it to their website.
$logo = "https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_92x30dp.png"

# Group Check Example
$Group = [ADSI]"LDAP://cn=IT Staff,OU=Groups,DC=Example,DC=co,DC=uk"
$Group.Member | ForEach-Object {
if ($user.distinguishedname -match $_) {
      $ItStaff = $true
    }
}

# Building HTML
$signature = 
@"
<div style="color: #000; font-family: Arial, Helvetica, sans-serif; font-size: 12">
    <p>
        <b>$displayName</b>
        $(if($poBox){"<b>"+$jobTitle})
    </p>
    <p>
    $(if($website -And $logo){"<a href="+$website+">"})
        $(if($logo){"<img src='$logo' />"})
    $(if($website -And $logo){"</a>"})
    </p>

    <p>
        $(if($poBox){"<b>"+$poBox+"</b><br>"})
        $(if($street){ $street+"<br>"})
        $(if($city){ $city+"<br>"})
        $(if($state){ $state+"<br>"})
        $(if($zipcode){ $zipCode+"<br>"})
    </p>

    <p>
        $(if($ITStaff){"For IT Helpdesk Call: 019201 20202993"})
        <table border="0" style="font-family: Arial, Helvetica, sans-serif; font-size: 12; color: #000">
            $(if($telephone){"<tr><td>t: </td><td>"+$telephone+"</td></tr>"})
            $(if($directDial){"<tr><td>dd: </td><td>"+$directDial+"</td></tr>"})
            $(if($mobileNumber){"<tr><td>m: </td><td>"+$mobileNumber+"</td></tr>"})
            $(if($email){"<tr><td>e: </td><td><a href='mailto:$email'>"+$email+"</td></tr>"})
            $(if($website){"<tr><td>w: </td><td><a href='$website'><b>"+$website+"</b></td></tr>"})
        </table>
    </p>
</div>
"@

# Save the HTML to the signature file
$Filename  = "$folderLocation\\$filename.htm"
$signature | out-file $Filename -encoding ascii

# Setting the regkeys for Outlook 2016
if (test-path "HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General") 
{
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value $filename -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value $filename -propertytype string -force
    Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook\\Setup -Name "First-Run" -ErrorAction silentlycontinue
}

# Setting the regkeys for Outlook 2010 - Thank you AJWhite1970 for the 2010 registry keys
if (test-path "HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General") 
{
    get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ MailSettings | new-Itemproperty -name NewSignature -value $filename -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\ MailSettings | new-Itemproperty -name ReplySignature -value $filename -propertytype string -force
    Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\14.0\\Outlook\\Setup -Name "First-Run" -ErrorAction silentlycontinue
}