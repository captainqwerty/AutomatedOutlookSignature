[CmdletBinding()]
param (
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]
    $CompanyName,
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]
    $Website,
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("ascii", "utf8nobom", "unicode", "utf32")]
    [string]
    $Encoding = "utf8",
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [bool]
    $ForceSignatureOnNewEmails = $true,
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [bool]
    $ForceSignatureOnReplies = $true
)

<#
.SYNOPSIS
    Automated Outlook Signature

.DESCRIPTION
    Using values from the users on-prem Active Directory account this script will create 
    generate a HTML signature and set this signature as the users default signature for 
    new emails and replies.  The HTML and CSS can be customised. 

.NOTES
    Author           : CaptainQwerty
    Release Date     : 19/05/2024
    Script Version   : 5.0.1
    GitHub Repo      : https://github.com/CaptainQwerty/AutomatedOutlookSignature
    ReadMe           : https://github.com/CaptainQwerty/AutomatedOutlookSignature/blob/main/README.md
#>

<#-----[ Latest Patch Notes ]-----#

Version 5.0.1
    * Minor spelling mistakes fixed.
    * Completed the detailed description.
    * Added -ForceSignatureOnNewEmails and -ForceSignatureOnReplies launch parameters.
    * Default encoding swapped out for UTF-8 when -Encoding launch parameter not specified.
    * Neatened the HTML and TXT saving. 
    * Swapped UTF8 for utf8nobom to stop random character showing up on the new Outlook

#>

# [ Configuration ]

Write-Verbose "Version 5.0.1"
$logo = "https://example.com/logo.png"
$folderLocation = Join-Path -Path $Env:appdata -ChildPath 'Microsoft\signatures'
$fileName = 'Signature'
$file  = Join-Path -Path $folderLocation -ChildPath $fileName

# [ Classes ]

class UserAccount {
    [string]$Name
    [string]$DistinguishedName
    [string]$UserPrincipalName
    [string]$DisplayName
    [string]$GivenName
    [string]$Initials
    [string]$Surname
    [string]$Description
    [string]$JobTitle
    [string]$Department
    [string]$Company
    [string]$EmailAddress
    [string]$StreetAddress
    [string]$City
    [string]$State
    [string]$PostalCode
    [string]$Country
    [string]$TelephoneNumber
    [string]$Mobile
    [string]$Pager
    [string]$Fax
    [string]$HomePhoneNumber
    [string]$OtherHomePhoneNumber
    [string]$HomeFax
    [string]$OtherFax
    [string]$IPPhone
    [string]$OtherIPPhone
    [string]$WebPage
    [string]$ExtensionAttribute1
    [string]$ExtensionAttribute2
    [string]$ExtensionAttribute3
    [string]$ExtensionAttribute4
    [string]$ExtensionAttribute5
    [string]$ExtensionAttribute6
    [string]$ExtensionAttribute7
    [string]$ExtensionAttribute8
    [string]$ExtensionAttribute9
    [string]$ExtensionAttribute10
    [string]$ExtensionAttribute11
    [string]$ExtensionAttribute12
    [string]$ExtensionAttribute13
    [string]$ExtensionAttribute14
    [string]$ExtensionAttribute15
}

# [ Functions ]

function Get-UserDetails {
    [CmdletBinding()]
    param ()

    try {
        $user = (([adsisearcher]"(&(objectCategory=User)(samaccountname=$env:username))").FindOne().Properties)
        Write-Verbose "User found for $($env:username)"

        $userAccount = New-Object UserAccount
        $userAccount.Name = $user['name'] -join ''
        $userAccount.DistinguishedName = $user['distinguishedname'] -join ''
        $userAccount.UserPrincipalName = $user['userprincipalname'] -join ''
        $userAccount.DisplayName = $user['displayname'] -join ''
        $userAccount.GivenName = $user['givenname'] -join ''
        $userAccount.Initials = $user['initials'] -join ''
        $userAccount.Surname = $user['sn'] -join ''
        $userAccount.Description = $user['description'] -join ''
        $userAccount.JobTitle = $user['title'] -join ''
        $userAccount.Department = $user['department'] -join ''
        $userAccount.EmailAddress = $user['mail'] -join ''
        $userAccount.StreetAddress = $user['streetaddress'] -join ''
        $userAccount.City = $user['l'] -join ''
        $userAccount.State = $user['st'] -join ''
        $userAccount.PostalCode = $user['postalcode'] -join ''
        $userAccount.Country = $user['c'] -join ''
        $userAccount.TelephoneNumber = $user['telephonenumber'] -join ''
        $userAccount.Mobile = $user['mobile'] -join ''
        $userAccount.Pager = $user['pager'] -join ''
        $userAccount.Fax = $user['facsimiletelephonenumber'] -join ''
        $userAccount.HomePhoneNumber = $user['homephone'] -join ''
        $userAccount.OtherHomePhoneNumber = $user['otherhomephone'] -join ''
        $userAccount.HomeFax = $user['homefax'] -join ''
        $userAccount.OtherFax = $user['otherfacsimiletelephonenumber'] -join ''
        $userAccount.IPPhone = $user['ipphone'] -join ''
        $userAccount.OtherIPPhone = $user['otheripphone'] -join ''

        $userAccount.Company = if ($null -eq $CompanyName) { $user['company'] } else { $CompanyName }
        $userAccount.WebPage = if ($null -eq $Website) { $user['wWWHomePage'] } else { $Website }
        
        $userAccount.ExtensionAttribute1 = $user['extensionattribute1'] -join ''
        $userAccount.ExtensionAttribute2 = $user['extensionattribute2'] -join ''
        $userAccount.ExtensionAttribute3 = $user['extensionattribute3'] -join ''

        return $userAccount
    } catch {
        Write-Error "Unable to query Active Directory for user information. Details: $($_.Exception.Message)"
        exit 1
    }
}

function Set-RegistryEntries {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Version,
        [Parameter(Mandatory = $true)]
        [string] $SignatureName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [bool]
        $ForceSignatureOnNewEmails,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [bool]
        $ForceSignatureOnReplies
    )

    $regPathGeneral = "HKCU:\Software\Microsoft\Office\$Version\Common\General"
    $regPathMailSettings = "HKCU:\Software\Microsoft\Office\$Version\Common\MailSettings"
    $regPathOutlookSetup = "HKCU:\Software\Microsoft\Office\$Version\Outlook\Setup"

    Write-Verbose "Setting registry keys"
    New-ItemProperty -Path $regPathGeneral -Name "Signatures" -Value "signatures" -PropertyType String -Force | Out-Null
    if($ForceSignatureOnNewEmails){ New-ItemProperty -Path $regPathMailSettings -Name "NewSignature" -Value $SignatureName -PropertyType String -Force | Out-Null }
    if($ForceSignatureOnReplies){ New-ItemProperty -Path $regPathMailSettings -Name "ReplySignature" -Value $SignatureName -PropertyType String -Force | Out-Null }

    Write-Verbose "Removing First-Run registry key"
    Remove-ItemProperty -Path $regPathOutlookSetup -Name "First-Run" -ErrorAction SilentlyContinue | Out-Null
}

function Get-OfficeVersion {
    [CmdletBinding()]
    param ()

    Write-Verbose "Getting office version"
    $officeVersions = @("16.0", "15.0", "14.0")
    $basePath = "HKCU:\Software\Microsoft\Office\"

    foreach ($version in $officeVersions) {
        $path = $basePath + $version + "\Outlook"
        if (Test-Path -Path $path) {
            Write-Verbose "Office Version: $version"
            return $version
        }
    }

    Write-Error "No compatible version of Microsoft Office found."
    exit 1
}

function Get-SignatureHTML {
    [CmdletBinding()]
    param (
        [UserAccount]$user
    )

    Write-Verbose "Building HTML"

    return @"
$(if($user.DisplayName){"<span><b>$($user.DisplayName)</b></span><br />"})
$(if($user.JobTitle){"<span>$($user.JobTitle)</span><br /><br />"})

<p>$(if($logo){"<img src='$logo' />"})</p>

<p>
$(if($user.Company){ "<b>$($user.Company)</b><br />" })
$(if($user.StreetAddress){ "$($user.StreetAddress), " })
$(if($user.City){ "$($user.City), " })
$(if($user.State){ "$($user.State), " })
$(if($user.PostalCode){ $user.PostalCode })
</p>

<table>
    $(if($($user.TelephoneNumber)){"<tr><td><b>T:</b></td><td><a href='tel:$($user.TelephoneNumber)'>$($user.TelephoneNumber)</a></td></tr>"})
    $(if($($user.Mobile)){"<tr><td><b>M:</b></td><td><a href='tel:$($user.Mobile)'>$($user.Mobile)</a></td></tr>"})
    $(if($user.EmailAddress){"<tr><td><b>E:</b></td><td><a href='mailto:$($user.EmailAddress)'>$($user.EmailAddress)</a></td></tr>"})
    $(if($($user.WebPage)){"<tr><td><b>W:</b></td><td><a href='https://$($user.WebPage)'>$($user.WebPage)</a></td></tr>"})
</table>
<br />
"@
}

function Get-SignaturePlainText {
    [CmdletBinding()]
    param (
        [UserAccount]$user
    )

    Write-Verbose "Building Plaint Text"

    return @"
$(if($user.DisplayName){ $user.DisplayName })
$(if($user.JobTitle){ $user.JobTitle })
___________________________________________________________

$(if($user.Company){ $($user.Company) })
$(if($user.StreetAddress){ "$($user.StreetAddress), " })
$(if($user.City){ "$($user.City), " })
$(if($user.State){ "$($user.State), " })
$(if($user.PostalCode){ $user.PostalCode })

$(if($($user.TelephoneNumber)){"T: $($user.TelephoneNumber)"})
$(if($($user.Mobile)){"M: $($user.Mobile)"})
$(if($user.EmailAddress){"E: $($user.EmailAddress)"})
$(if($($user.WebPage)){"W: $($user.WebPage)"})
"@
}

# [ CSS ]

$style = 
@"
<style>
p, table, a, span { 
    font-family: Arial, Helvetica, sans-serif;
    font-size:  12pt;
}
</style>
"@

# [ Execution ]

$officeVersion = Get-OfficeVersion
$user = Get-UserDetails

# If the folder does not exist create it
if (-not (Test-Path -Path $folderLocation)) {
    try {
        Write-Verbose "Creating Director: $folderLocation"
        New-Item -ItemType directory -Path $folderLocation
    } catch {
        Write-Error "Error: Unable to create the signatures folder. Details: $($_.Exception.Message)"
        exit
    }
}

$htmlSignature = Get-SignatureHTML -user $user
$plainTextSignature = Get-SignaturePlainText -user $user

# Saving HTML and plain Text signatures
try {
    Write-Verbose "Saving HTML signature"
    $style + $htmlSignature | Out-File -FilePath "$file.htm" -Encoding $Encoding

    Write-Verbose "Saving txt signature"
    $plainTextSignature | out-file "$file.txt" -encoding $Encoding
} catch {
    Write-Error "Error: Unable to save signature file. Details: $($_.Exception.Message)"
    exit 1
}

Set-RegistryEntries -Version $officeVersion -SignatureName $fileName -ForceSignatureOnNewEmails $ForceSignatureOnNewEmails -ForceSignatureOnReplies $ForceSignatureOnReplies