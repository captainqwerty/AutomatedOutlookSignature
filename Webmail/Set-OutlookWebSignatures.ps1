#Requires -Module ExchangeOnlineManagement
#Requires -Module ActiveDirectory

<#
.SYNOPSIS
  Outlook Web Signature Script - https://github.com/captainqwerty/AutomatedOutlookSignature
.DESCRIPTION
  This script gets each users AD details one at a time, builds their HTML signature and sets it as their Outlook web signature.
.INPUTS
  
.OUTPUTS
  All signatures are stored in the $signatureFolder
.NOTES
  Version:        1.0
  Author:         CaptainQwerty 
  Creation Date:  07/12/2020
#>

#-----[ Configuration ]-----#

# Sets the folder to store all web signatures and if the folder does exists it will be created.
$signatureFolder = "$psscriptroot\Web-Signatures"

if(!(test-path $signatureFolder)) {
    New-Item -ItemType "directory" -Path $signatureFolder
}

#-----[ Functions ]-----#

# This is the function that creates their signature file if it does not exists, or updates it if it differs.
function Create-WebSignatures {

    # Gets each of the users one at a time from AD and creates their signature
    $signaturesToUpdate = @()

    # Gets all users in the Outlook Signature Group
    $allStaff = Get-ADGroupMember "Outlook Signature"

    # For each user in the All Staff group the below will be ran
    $allStaff | ForEach-Object {

        # Stores the users details in $user
        $user = Get-Aduser -Identity $_.distinguishedname -Properties Title, MobilePhone, EmailAddress, extensionattribute1, extensionattribute2, streetaddress, st, l, postalcode, telephonenumber
        
        # If the user is disabled they are skipped
        if(!$user.Enabled) 
        {
            Write-Host "$($user.Name)'s account is disabled and will be skipped"
            return
        }
        
        # Saving the users properties in slightly more user frinedly named variables
        $username = ($user.userprincipalName).Substring(0, $user.userprincipalname.IndexOf('@'))
        $displayName = $user.Name
        $jobTitle = $user.Title
        $mobileNumber = $user.MobilePhone
        $email = $user.EmailAddress
        $namePrefix = $user.extensionattribute1 # Dr etc.
        $namePostfix = $user.extensionattribute2 # Bs(Hons) etc.

        # These are details you can either get from Active directory or as they might be the same for your entire company could statically set them here. Each has a commented out static example, simply swap the commented lines and alter the example.
        $street = $user.streetaddress # Street address
        $city = $user.l # City
        $state = $user.st # State
        $zipCode = $user.postalcode # Postcode 
        $telephone = $user.telephonenumber # Telephone number
        $website = "www.example.co.uk" # Website
        $logo = "https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_92x30dp.png" # Logo

        # Gathers a list of all groups the user is a member of
        $groups = Get-ADPrincipalGroupMembership $_ | select name

        # Group Check Example
        $Group = [ADSI]"LDAP://cn=IT Staff,OU=Groups,DC=Example,DC=co,DC=uk"
        $Group.Member | ForEach-Object {
        if ($user.distinguishedname -match $_) {
                $ItStaff = $true
            }
        }

        $address = $null

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

        # If the file exists it will compare them for changes and if there are changes it will mark them as needing an update
        if(test-path "$signatureFolder\$email.html"){          
            $currentSig = (Get-Content "$signatureFolder\$email.html" | out-string).TrimEnd()

            if($currentSig -eq $signature)
            {
                write-host "Signature found for $displayname - No update required." -ForegroundColor green
            } else {
                write-host "Signature found for $displayname - Update required." -ForegroundColor yellow
                Remove-Item -Path "$signatureFolder\$email.html" -Force
                $signature | Out-File "$signatureFolder\$email.html"
                $signaturesToUpdate += $email
            }
        } else {
            write-host "No Signature found for $displayName. Creating Signature" -ForegroundColor Red
            $signature | Out-File "$signatureFolder\$email.html"
            $signaturesToUpdate += $email
        }
    }
    return $signaturesToUpdate
}

# This function is passed all the users who have no signature or need an update and uses their signature file to update their web signature.
function Update-WebSignatures {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $userEmailAddress
    )

    write-host "Setting signature on mailbox: $userEmailAddress"
    
    # Gets their signature content out of the file.
    $signature = Get-Content "$signatureFolder\$userEmailAddress.html"

    # Sets the users signature to the content from within the file. 
    Get-Mailbox $userEmailAddress | Set-MailboxMessageConfiguration -SignatureHTML $signature -AutoAddSignature:$true
}

#-----[ Execution ]-----#

# Creates all new/modified signatures and outputs a list of who needs them altered online
$usersToUpdate = Create-WebSignatures

# If any users need a new / modified signature this section will run
if($usersToUpdate.Count -gt 0) {
    try {
        Write-Host "Connecting to Exchange Online"
        Connect-ExchangeOnline

        foreach ($usertoUpdate in $usersToUpdate)
        {
            # This calls the function for each of the users that needs an updated signature
            Update-WebSignatures $usertoUpdate
        }
    } catch {
        write-host "Oh dear something went wrong" -ForegroundColor Red
    } finally {
        Write-Host "Disconneting Exchange Online"
        Disconnect-ExchangeOnline -Confirm:$false
    }
}