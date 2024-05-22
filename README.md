[![Contributors][contributors-shield]][contributors-url] 
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]

# Automated Outlook Signature Script
The script retrieves user details from Active Directory, generates a new signature file, and sets it as the default Outlook signature. It ensures that any changes in user details, such as job title, are reflected in the signature during the next logon.

## Table of Contents
- [Features](#features)
- [Usage](#usage)
- [Parameters](#parameters)
- [Installation with Group Policy](#installation-with-group-policy)
- [Known Issues and Caveats](#known-issues-and-caveats)
- [Contributing](#contributing)
- [License](#license)
- [Acknowledgements](#acknowledgements)

## Features
- Supports multiple versions of Outlook.
- Retrieves user details from Active Directory.
- Generates HTML and plain text signatures.
- Sets registry keys to configure the default Outlook signature.
- Supports verbose output for detailed execution.

## Usage
I recommend using the script in Group Policy as a logon script. During the user's logon process, the script runs in the background, retrieves the necessary user details, generates a new signature file, and replaces the existing one.

For detailed instructions on configuring logon scripts with Group Policy, refer to this article: [Configuring Logon PowerShell Scripts with Group Policy - 4Sysops](https://4sysops.com/archives/configuring-logon-powershell-scripts-with-group-policy/). 

## Dynamic User Details

The script retrieves various user details from Active Directory to dynamically generate the signature. The following properties from the `UserAccount` class can be used within the signature:

- **Name**: Full name of the user.
- **DistinguishedName**: Distinguished name in Active Directory.
- **UserPrincipalName**: User's principal name (UPN).
- **DisplayName**: Display name of the user.
- **GivenName**: Given (first) name.
- **Initials**: User's initials.
- **Surname**: User's surname (last name).
- **Description**: Description of the user.
- **JobTitle**: Job title of the user.
- **Department**: Department the user belongs to.
- **Company**: Company name.
- **EmailAddress**: Email address of the user.
- **StreetAddress**: Street address.
- **City**: City of the user.
- **State**: State or province.
- **PostalCode**: Postal code.
- **Country**: Country or region.
- **TelephoneNumber**: Primary telephone number.
- **Mobile**: Mobile phone number.
- **Pager**: Pager number.
- **Fax**: Fax number.
- **HomePhoneNumber**: Home phone number.
- **OtherHomePhoneNumber**: Additional home phone number.
- **HomeFax**: Home fax number.
- **OtherFax**: Additional fax number.
- **IPPhone**: IP phone number.
- **OtherIPPhone**: Additional IP phone number.
- **WebPage**: Web page URL.
- **ExtensionAttribute1**: Custom extension attribute 1.
- **ExtensionAttribute2**: Custom extension attribute 2.
- **ExtensionAttribute3**: Custom extension attribute 3.
- **ExtensionAttribute4**: Custom extension attribute 4.
- **ExtensionAttribute5**: Custom extension attribute 5.
- **ExtensionAttribute6**: Custom extension attribute 6.
- **ExtensionAttribute7**: Custom extension attribute 7.
- **ExtensionAttribute8**: Custom extension attribute 8.
- **ExtensionAttribute9**: Custom extension attribute 9.
- **ExtensionAttribute10**: Custom extension attribute 10.
- **ExtensionAttribute11**: Custom extension attribute 11.
- **ExtensionAttribute12**: Custom extension attribute 12.
- **ExtensionAttribute13**: Custom extension attribute 13.
- **ExtensionAttribute14**: Custom extension attribute 14.
- **ExtensionAttribute15**: Custom extension attribute 15.

These details can be customised and included in the signature template to ensure that each user's signature is personalised and up-to-date with the latest information from Active Directory.

## Parameters
- **Verbose**: Enables verbose output.
- **CompanyName**: Overwrite / specifies the company name to use in the signature.
- **Website**: Overwrite / specifies the company website to include in the signature.

### Example
```powershell
# Normal Execution
.\AutomatedOutlookSignature.ps1 
```

### Example
```powershell
# Normal Execution with alternative encoding. Default is Unicode.
.\AutomatedOutlookSignature.ps1 -encoding ascii
```

### Example
```powershell
# Execution with verbose output
.\AutomatedOutlookSignature.ps1 -Verbose
```

### Example
```powershell
# Execution but statically forcing the company name and website URL
.\AutomatedOutlookSignature.ps1 -CompanyName "YourCompany" -Website "www.yourcompany.com"
```

## Installation with Group Policy
1. Download the latest version of the script from the [GitHub repository](https://github.com/CaptainQwerty/AutomatedOutlookSignature).
2. Edit the script and update the value of `$logo` to your publicly available logo image.
3. Place the script in a location accessible to your users, such as a network share or GPO script folder.
4. Configure Group Policy to run the script at logon. Refer to the provided [Configuring Logon PowerShell Scripts with Group Policy - 4Sysops](https://4sysops.com/archives/configuring-logon-powershell-scripts-with-group-policy/) article for guidance. You can also watch this [YouTube Video Guide](https://www.youtube.com/watch?v=rt9y02iBoPE).

## Known Issues and Caveats
Below is a list of any known issues and caveats.

- Outlook will not have an signature on the very first launch but will after that

## Contributing
Contributions are welcome! Please visit the [GitHub repository](https://github.com/CaptainQwerty/AutomatedOutlookSignature) to open an issue or submit a pull request.

## License
This project is licensed under the MIT License. See the [LICENSE](https://github.com/CaptainQwerty/AutomatedOutlookSignature/blob/master/LICENSE) file for details.

## Acknowledgements
- [EduGeek Forum](http://www.edugeek.net/forums/scripts/205976-outlook-email-signature-automation-ad-attributes.html#post1760284) for discussions and feedback.

## Changelog

### [5.0.1] - In Development
- Minor spelling mistakes fixed.
- Completed the detailed description.
- Added -ForceSignatureOnNewEmails and -ForceSignatureOnReplies launch parameters.
- Default encoding swapped out for UTF-8 when -Encoding launch parameter not specified.
- Neatened the HTML and TXT saving. 

### [5.0.0] - 20/05/2024
- New script layout utilising functions to enhance readability.
- Added support for verbose output.
- Introduced parameters for script execution. Website and CompanyName can be statically set as parameters.
- Refactored code to use a class for better structure and readability.
- Improved compatibility with more versions of Office.
- Removed group check example to reduce script run time.
- Added a parameter to easily switch encoding.

[contributors-shield]: https://img.shields.io/github/contributors/CaptainQwerty/AutomatedOutlookSignature.svg?style=for-the-badge
[contributors-url]: https://github.com/CaptainQwerty/AutomatedOutlookSignature/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/CaptainQwerty/AutomatedOutlookSignature.svg?style=for-the-badge
[forks-url]: https://github.com/CaptainQwerty/AutomatedOutlookSignature/network/members
[stars-shield]: https://img.shields.io/github/stars/CaptainQwerty/AutomatedOutlookSignature.svg?style=for-the-badge
[stars-url]: https://github.com/CaptainQwerty/AutomatedOutlookSignature/stargazers
[issues-shield]: https://img.shields.io/github/issues/CaptainQwerty/AutomatedOutlookSignature.svg?style=for-the-badge
[issues-url]: https://github.com/CaptainQwerty/AutomatedOutlookSignature/issues
[license-shield]: https://img.shields.io/github/license/CaptainQwerty/AutomatedOutlookSignature.svg?style=for-the-badge
[license-url]: https://github.com/CaptainQwerty/AutomatedOutlookSignature/blob/master/LICENSE