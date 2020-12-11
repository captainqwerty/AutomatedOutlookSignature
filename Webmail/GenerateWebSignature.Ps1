<#
.SYNOPSIS
  Dynamic Outlook Web Mail Signature Creation - https://github.com/captainqwerty/AutomatedOutlookSignature
.DESCRIPTION
  This script uses properties from ActiveDirectory to generalte a HTML signautre which is then set as the users Outlook web signature.
  The script uses if statements to ensure if a user, for example, doesn't have a mobile number then that entire section of the signature is not uncluded, this makes the signature look much more professional and 
  if the user is given a mobile number, at next logon it will be added to their signature. 
  The script can also add additional parts based on group membership, the example given in this script is being a member of IT Staff adds Helpdesk contact information. 
.INPUTS
  Several properties for the user are taken directly from Active Directory, for more details on this please see the ReadMe on GitHub.
.OUTPUTS
  
.NOTES
  Version:        1.0
  Author:         CaptainQwerty
  Modified:       07/12/2020
#>

#requires -Module ExchangeOnlineManagment
Import-Module ExchangeOnlineManagement

Connect-ExchangeOnline -UserPrincipalName example@example.com

$webEmailSignature = [ADSI]"LDAP://CN=Outlook - Web Signature,OU=Groups,DC=Example,DC=co,DC=uk"

$webEmailSignature.Member | ForEach-Object {
    $Searcher = [adsisearcher]"(distinguishedname=$_)"
    $user = $searcher.FindOne().Properties

    if($user) {
        write-host "Setting siguture for" $user.name[0]

        # Company name and logo
        $companyName = "Company Name"
        $logo = "https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_92x30dp.png" # Please note that if you do include a logo it must be located somewhere on the internet that the public has access to, many users upload it to their website.

        # Get the users properties (These should always be in Active Directory and Unique)
        if($user.name.count -gt 0){$displayName = $user.name[0]}
        if($user.title.count -gt 0){$jobTitle = $user.title[0]}
        if($user.homephone.count -gt 0){$directDial = $user.homephone[0]}
        if($user.mobile.count -gt 0){$mobileNumber = $user.mobile[0]}
        if($user.mail.count -gt 0){$email = $user.mail[0]}
        if($user.wwwhomepage.count -gt 0){$website = $user.wwwhomepage[0]}

        # Address
        if($user.postofficebox.count -gt 0){$poBox = $user.postofficebox[0]}
        if($user.streetaddress.count -gt 0){$street = $user.streetaddress[0]}
        if($user.l.count -gt 0){$city = $user.l[0]}
        if($user.st.count -gt 0){$state = $user.st[0]}
        if($user.postalcode.count -gt 0){$zipCode = $user.postalcode[0]}
        if($user.telephonenumber.count -gt 0){$telephone = $user.telephonenumber[0]}
        if($user.physicaldeliveryofficename.count -gt 0){$office = $user.physicaldeliveryofficename}

        # Group Check Example
        $Group = [ADSI]"LDAP://cn=IT Staff,OU=Groups,DC=Example,DC=co,DC=uk"
        $Group.Member | ForEach-Object {
        if ($user.distinguishedname -match $_) {
                $ItStaff = $true
            }
        }

        # Building address
        if($street){ $address = "$($street), " } 
        if($city){ $address = $address + "$($city), " }
        if($state){ $address = $address + "$($state), " }
        if($zipCode){ $address = $address + $zipCode }

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
                    $(if($website){"<tr><td>W: <a href='https://$website'>$($website)</a></td></tr>"})
                </table>
            </td>
        </tr>
    </table>
  </p>
  <br />
"@
}

  Get-Mailbox $user.mail[0] | Set-MailboxMessageConfiguration -SignatureHTML $signature -AutoAddSignature:$true
}