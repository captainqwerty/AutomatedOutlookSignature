<#
.SYNOPSIS
  Dynamic Outlook Signature Creation - https://github.com/captainqwerty/AutomatedOutlookSignature
.DESCRIPTION
  This script uses properties from ActiveDirectory to populate the .htm and .txt file which are then stored in the $folderlocation with the name $filename.htm and $filename.txt. 
  The script can be ran as either a scheduled task at logon or preferably a Group Policy Logon script, more details on this can be found in the GitHub ReadMe.
  The script uses if statements to ensure if a user, for example, doesn't have a mobile number then that entire section of the signature is not uncluded, this makes the signature look much more professional and 
  if the user is given a mobile number, at next logon it will be added to their signature. 
  The script can also add additional parts based on group membership, the example given in this script is being a member of IT Staff adds Helpdesk contact information. 
.INPUTS
  Several properties for the user are taken directly from Active Directory, for more details on this please see the ReadMe on GitHub.
.OUTPUTS
  $folderlocation\$filename.htm - HTML signature for rich text emails
  $folderlocation\$filename.txt - Text signautre for none rich text emails
.NOTES
  Version:        3.0
  Author:         CaptainQwerty
  Modified:       29/02/2020
#>

# Getting Active Directory information for current user
$user = (([adsisearcher]"(&(objectCategory=User)(samaccountname=$env:username))").FindOne().Properties)

# If the user is not found in Active Directory exit the script
if (-not $user) {
  exit
}

# Create the signatures folder and sets the name of the signature file
$folderLocation = Join-Path -Path $Env:appdata -ChildPath 'Microsoft\signatures'
$filename = 'Signature'
$file  = Join-Path -Path $folderLocation -ChildPath $filename

# If the folder does not exist create it
if (-not (Test-Path -Path $folderLocation)) {
  New-Item -ItemType directory -Path $folderLocation
}

# Company name and logo
$companyName = 'Company Name'
$logo = 'https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_92x30dp.png' # Please note that if you do include a logo it must be located somewhere on the internet that the public has access to, many users upload it to their website.

# Get the users properties (These should always be in Active Directory and Unique)
if($user.name.count -gt 0){$displayName = $user.name[0]} # Display Name
if($user.title.count -gt 0){$jobTitle = $user.title[0]} # Job Title
if($user.mail.count -gt 0){$email = $user.mail[0]} # Email Address

if($user.mobile.count -gt 0){$mobileNumber = $user.mobile[0]} # Mobile number
if($user.homephone.count -gt 0){$directDial = $user.homephone[0]} # Home Nnumber
if($user.telephonenumber.count -gt 0){$telephone = $user.telephonenumber[0]} # Office number

if($user.wwwhomepage.count -gt 0){$website = $user.wwwhomepage[0]} # Webpage address

# Address
if($user.postofficebox.count -gt 0){$poBox = $user.postofficebox[0]} # PO box number
if($user.physicaldeliveryofficename.count -gt 0){$office = $user.physicaldeliveryofficename} # Office Name
if($user.streetaddress.count -gt 0){$street = $user.streetaddress[0]} # Street address
if($user.l.count -gt 0){$city = $user.l[0]} # City
if($user.st.count -gt 0){$state = $user.st[0]} # State
if($user.postalcode.count -gt 0){$zipCode = $user.postalcode[0]} # Post code / Zip code

# Group Check Example
$Group = [ADSI]"LDAP://cn=IT Staff,OU=Groups,DC=Example,DC=co,DC=uk"
$Group.Member | ForEach-Object {
if ($user.distinguishedname -match $_) {
      $ItStaff = $true
    }
}

# Building Style Sheet
$style = 
@"
<style>
p, table, td, tr, a, span { 
    font-family: Arial, Helvetica, sans-serif;
    font-size:  12pt;
    color: #28b8ce;
}

span.blue
{
    color: #28b8ce;
}

table {
    margin: 0;
    padding: 0;
}

a { 
text-decoration: none;
}

hr {
border: none;
height: 1px;
background-color: #28b8ce;
color: #28b8ce;
width: 700px;
}

table.main {
    border-top: 1px solid #28b8ce;
}
</style>
"@

# Building HTML
$signature = 
@"
  $(if($displayName){"<span><b>"+$displayName+"</b></span><br />"})
  $(if($jobTitle){"<span>"+$jobTitle+"</span><br /><br />"})

<p>
  <table class='main'>
      <tr>
          <td style='padding-right: 75px;'>$(if($logo){"<img src='$logo' />"})</td>
          <td>
              <table>
                  <tr><td colspan='2' style='padding-bottom: 10px;'>
                    $(if($companyName){ "<b>"+$companyName+"</b><br />" })
                    $(if($street){ $street+", " })
                    $(if($city){ $city+", " })
                    $(if($state){ $state+", " })
                    $(if($zipCode){ $zipCode })
                  </td></tr>
                  $(if($ITMember){"<tr><td td colspan='2'>IT Helpdesk: 0188887 55555 6666</tr></td>"})
                  $(if($telephone){"<tr><td>T: </td><td><a href='tel:$telephone'>$($telephone)</a></td></tr>"})
                  $(if($mobileNumber){"<tr><td>M: </td><td><a href='tel:$mobileNumber'>$($mobileNumber)</a></td></tr>"})
                  $(if($email){"<tr><td>E: </td><td><a href='mailto:$email'>$($email)</a></td></tr>"})
                  $(if($website){"<tr><td>W:</td><td><a href='https://$website'>$($website)</a></td></tr>"})
              </table>
          </td>
      </tr>
  </table>
</p>
<br />
"@

# Save the HTML to the signature file
$style + $signature | out-file "$file.htm" -encoding ascii

# Build the txt version for none rich text emails
$signature = 
@"
$(if($displayName){ $displayName })
$(if($jobTitle){ $jobTitle })
___________________________________________________________

$(if($companyName){ $companyName })
$(if($street){ $street+", " })
$(if($city){ $city+", " })
$(if($state){ $state+", " })
$(if($zipCode){ $zipCode })

$(if($ITMember){"For IT Helpdesk Call 0191231 212313"})
$(if($telephone){"T: "+$telephone})
$(if($mobileNumber){"M: "+$mobileNumber})
$(if($email){"E: "+$email})
$(if($website){"W: "+$website})
"@

# Output the text to the signatures folder
$signature | out-file "$file.txt" -encoding ascii

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
