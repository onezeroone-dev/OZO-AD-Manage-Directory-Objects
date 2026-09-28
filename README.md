# OZO AD Manage Directory Objects Installation and Usage
## Description
This script creates Active Directory contacts, computers, groups, [empty\] group policies, organizational units, and users based on a JSON configuration file.

## Prerequisites
This script requires the _ActiveDirectory_, _GroupPolicy_, _ImportExcel_, _OZO_, _OZOAD_, _OZOFiles_,and _OZOLogger_ PowerShell modules. The ActiveDirectory and GroupPolicy modules are included with the [Remote Server Administration Tools](https://learn.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools) installation. The remaining modules are published to [PowerShell Gallery](https://learn.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-5.1). Ensure your system is configured for this repository then execute the following in an _Administrator_ PowerShell:

```powershell
Install-Module ImportExcel,OZO,OZOAD,OZOFiles,OZOLogger
```

## Installation
This script is published to [PowerShell Gallery](https://learn.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-5.1). Ensure your system is configured for this repository then execute the following in an _Administrator_ PowerShell:

```powershell
Install-Script ozo-ad-manage-directory-objects
```

## Usage
```powershell
ozo-ad-manage-directory-objects
    -Configuration <String>
    [-OutDir] <String>
```

## Parameters
|Parameter|Description|
|---------|-----------|
|`Configuration`|Path to the JSON configuration file. Defaults to `ozo-ad-manage-directory-objects.json` in the same directory as the script.|
|`OutDir`|Directory for the Excel results report. Defaults to the current directory.|

## JSON Configuration Definition
This script leverages the [One Zero One Unified AD JSON Schema](https://onezeroone.dev/ozo-unified-ad-json-schema/). The elements of the schema used by this script are as follows. Please also see [ozo-ad-manage-directory-objects-EXAMPLE.json](https://github.com/onezeroone-dev/OZO-AD-Manage-Directory-Objects/blob/main/ozo-ad-manage-directory-objects-EXAMPLE.json).

## Examples
```powershell
ozo-ad-manage-directory-objects -Configuration (Join-Path -Path $Env:USERPROFILE -ChildPath "Downloads\ozo-ad-manage-directory-objects.json")
```

## Logging
Messages as written to the Windows Event Viewer [_One Zero One_](https://github.com/onezeroone-dev/OZO-Windows-Event-Log-Provider-Setup/blob/main/README.md) provider when available. Otherwise, messages are written to the _Microsoft-Windows-PowerShell_ provider with event ID 4100.

## Licensing
This script is licensed under the [GNU General Public License (GPL) version 2.0](LICENSE).

## Notes
Run this script with a user that has rights to create comuters, contacts, groups, and users in the Active Directory domain.

## Acknowledgements
Special thanks to my employer, [Sonic Healthcare USA](https://sonichealthcareusa.com), who supports the growth of my PowerShell skillset and enables me to contribute portions of my work product to the PowerShell community.
