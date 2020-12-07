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

$webEmailSignature = [ADSI]"LDAP://CN=All Staff,OU=Two Castles,DC=tcha1,DC=co,DC=uk"

# $webEmailSignature.Member | ForEach-Object {
#     $Searcher = [adsisearcher]"(distinguishedname=$_)"
#     $user = $searcher.FindOne().Properties

#     $user.mail[0]
# }


$webEmailSignature.Member | ForEach-Object {
    $Searcher = [adsisearcher]"(distinguishedname=$_)"
    $user = $searcher.FindOne().Properties

    if($user) {
        write-host "Setting siguture for" $user.name[0]

        # Get the users properties (These should always be in Active Directory and Unique)
        $displayName = $user.name[0]
        $jobTitle = $user.title[0]
        $mobileNumber = $user.mobile[0]
        $email = $user.mail[0]
        $namePrefix = $user.extensionattribute1
        $namePostfix = $user.extensionattribute2

        # Address
        $street = $user.streetaddress[0] # Street address
        $city = $user.l[0] # City
        $state = $user.st[0] # State
        $zipCode = $user.postalcode[0] # Postcode 
        $telephone = $user.telephonenumber[0] # Telephone number
        $website = "www.castlesandcoasts.co.uk" # Website
        $logo = "https://castlesandcoasts.co.uk/media/10509/email-logo.jpg" # Logo

        # IT group check
        $GroupIT = [ADSI]"LDAP://cn=IT Staff,OU=Global Security Groups,OU=Two Castles,DC=tcha1,DC=co,DC=uk"
        $ITMember = $false
        $GroupIT.Member | ForEach-Object { 
                if ($user.distinguishedname -match $_) { 
                    $ITMember = $true
                }
        }

        # Hide Mobile Number group check
        $GroupNoMobile = [ADSI]"LDAP://cn=OutLook Signature - Hide Mobile,OU=Global Security Groups,OU=Two Castles,DC=tcha1,DC=co,DC=uk"
        $NoMobile = $false
        $GroupNoMobile.Member | ForEach-Object { 
                if ($user.distinguishedname -match $_) { 
                    $NoMobile = $true
                }
        }

        # CJSM Secure Email users
        $GroupCJSM = [ADSI]"LDAP://cn=CJSM Users,OU=Global Security Groups,OU=Two Castles,DC=tcha1,DC=co,DC=uk"
        $CJSMMember = $false
        $GroupCJSM.Member | ForEach-Object { 
                if ($user.distinguishedname -match $_) { 
                    $secureEmail = "$username@castlesandcoasts.cjsm.net"
                    $CJSMMember = $true
                }
        }

        # Building address
        if($street){ $address = "$($street), " } 
        if($city){ $address = $address + "$($city), " }
        if($state){ $address = $address + "$($state), " }
        if($zipCode){ $address = $address + $zipCode }


# Build the HTML 
$signature = 
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


<span><b>$(if($namePrefix){$namePrefix + " "})$displayName $(if($namePostfix){$namePostfix})</b></span><br>
$(if($jobTitle){"<span>"+$jobTitle+"</span><br /><br />"})

<p>
    <table class='main'>
        <tr>
            <td style='padding-right: 75px;'>$(if($logo){"<img src='$logo' />"})</td>
            <td>
                <table>
                    <tr><td colspan='2' style='padding-bottom: 10px;'>Castles & Coasts Housing Association<br>$($address)</td></tr>
                    $(if($ITMember){"<tr><td td colspan='2'>IT Helpdesk 01228 635426 or Ext 5426</tr></td>"})
                    $(if($telephone){"<tr><td>T:</td><td><a href='tel:$telephone'>$($telephone)</a></td></tr>"})
                    $(if($mobileNumber -And !$NoMobile){"<tr><td>M:</td><td><a href='tel:$mobileNumber'>$($mobileNumber)</a></td></tr>"})
                    $(if($email){"<tr><td>E:</td><td><a href='mailto:$email'>$($email)</a></td></tr>"})
                    $(if($CJSMMember){"<tr><td colspan='2'>Secure Email: <a href='mailto:$secureEmail'>$($secureEmail)</a></td></tr>"})
                    $(if($website){"<tr><td colspan='2'><b>Website: <a href='http://$website'>$($website)</a></b></td></tr>"})
                </table>
            </td>
        </tr>
    </table>
    <!-- Temp banner below signature -->
    <img src="https://castlesandcoasts.co.uk/media/11076/ukha_2020_finalist_small.png" height=150px/>
</p>
<br />
"@
}

        Get-Mailbox $user.mail[0] | Set-MailboxMessageConfiguration -SignatureHTML $signature -AutoAddSignature:$true
}