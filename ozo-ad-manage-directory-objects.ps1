#Requires -Modules ImportExcel,OZO,OZOLogger -Version 5.1

<#PSScriptInfo
    .VERSION 0.0.1
    .GUID 87eaa283-5e8b-4265-a025-ce1f5791491b
    .AUTHOR Andy Lievertz <alievertz@onezeroone.dev>
    .COMPANYNAME One Zero One
    .COPYRIGHT This script is released under the terms of the GNU General Public License ("GPL") version 2.0.
    .TAGS 
    .LICENSEURI https://github.com/onezeroone-dev/OZO-AD-Manage-Directory-Objects/blob/main/LICENSE
    .PROJECTURI https://github.com/onezeroone-dev/OZO-AD-Manage-Directory-Objects
    .ICONURI 
    .EXTERNALMODULEDEPENDENCIES
    .REQUIREDSCRIPTS 
    .EXTERNALSCRIPTDEPENDENCIES 
    .RELEASENOTES https://github.com/onezeroone-dev/OZO-AD-Manage-Directory-Objects/blob/main/CHANGELOG.md
#>

<# 
    .SYNOPSIS
    See description.
    .DESCRIPTION 
    Creates Active Directory users and groups based on a JSON configuration file.
    .PARAMETER Configuration
    Path to the JSON configuration file. Defaults to "ozo-ad-manage-directory-objects.json" in the same directory as the script.
    .PARAMETER OutDir
    Directory for the Excel report. Defaults to the current directory.
    .LINK
    https://github.com/onezeroone-dev/OZO-AD-Manage-Directory-Objects/blob/main/README.md
#>

# PARAMETERS
[CmdletBinding(SupportsShouldProcess = $true)] Param (
    [Parameter(Mandatory=$false,HelpMessage="Path to the JSON configuration file")][String]$Configuration = (Join-Path -Path $PSScriptRoot -ChildPath "ozo-ad-manage-directory-objects.json"),
    [Parameter(Mandatory=$false,HelpMessage="Path for the Excel report")][String]$OutDir = (Get-Location)
)

# CLASSES
Class OZOMain {
    # PROPERTIES: Booleans, Strings
    [Boolean] $Success   = $false
    [String]  $adDomain  = $null
    [String]  $jsonPath  = $null
    [String]  $outDir    = $null
    [String]  $excelPath = $null
    # PROPERTIES: PSCustomObjects
    [PSCustomObject] $Json   = $null
    [PSCustomObject] $Logger = $null
    # PROPERTIES: PSCustomObject Lists
    [System.Collections.Generic.List[PSCustomObject]] $adOUs       = @()
    [System.Collections.Generic.List[PSCustomObject]] $adContacts  = @()
    [System.Collections.Generic.List[PSCustomObject]] $adComputers = @()
    [System.Collections.Generic.List[PSCustomObject]] $adGroups    = @()
    [System.Collections.Generic.List[PSCustomObject]] $adGPOs      = @()
    [System.Collections.Generic.List[PSCustomObject]] $adUsers     = @()
    # METHODS: Constructor method
    OZOMain($Configuration,$OutDir) {
        # Set properties
        $this.Logger   = (New-OZOLogger)
        $this.jsonPath = $Configuration
        $this.outDir   = $OutDir
        # Log a process start message
        $this.Logger.Write("Starting process.","Information")
        # Determine if the configuation is valid
        If (($this.ValidateConfiguration() -And $this.ValidateEnvironment()) -eq $true) {
            # Iterate through the JSON OU objects
            ForEach ($adOU in $this.Json.ADOrganizationalUnits) {
                # Add an OZOADOrganizationalUnit object to the adOUs list
                $this.adOUs.Add(([OZOADOrganizationalUnit]::($adOU)))
            }
            # Iterate through the JSON Contacts objects
            ForEach ($adContact in $this.Json.ADContacts) {
                # Add an OZOADOContact object to the adContacts list
                $this.adContacts.Add(([OZOADContact]::new($adContact)))
            }
            # Iterate through the JSON Computer objects
            ForEach ($adComputer in $this.Json.ADComputers) {
                # Add an OZOADComputer object to the adComputers list
                $this.adComputers.Add(([OZOADComputer]::new($adComputer)))
            }
            # Iterate through the JSON Group objects
            ForEach ($adGroup in $this.Json.ADGroups) {
                # Add an OZOADGroup object to the adGroups list
                $this.adGroups.Add(([OZOADGroup]::new($adGroup)))
            }
            # Iterate through the JSON GPO objects
            ForEach ($adGPO in $this.Json.ADGroupPolicies) {
                # Add an OZOADGroupPolicyObject object to the adOUs list
                $this.adGPOs.Add(([OZOADGroupPolicyObject]::new($adOU)))
            }
            # Iterate through the JSON User objects
            ForEach($adUser in $this.Json.ADUsers) {
                # Add an OZOADUser object to the adUsers list
                $this.adUsers.Add(([OZOADUser]::new($adUser,($this.adDomain).Name)))
            }
            # Manage default objects
            $this.ManageDefaultObjects()
        }
        # Report
        $this.Report()
        # Log a process end message
        $this.Logger.Write(("Process complete."),"Information")
    }
    # METHODS: JSON validation method
    Hidden [Boolean] ValidateConfiguration() {
        # Control variable
        [Boolean] $Return = $true
        # Check that the jsonPath is valid
        If ((Test-Path -Path $this.jsonPath) -eq $true) {
            # Attempt to read the JSON
            Try {
                $this.Json = (Get-Content $this.jsonPath -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
                # Success
            } Catch {
                # Failure
                $this.Logger.Write(("Invalid JSON in " + $this.jsonPath + "."),"Error")
                $Return = $false
            }
        } Else {
            # JSON path is invalid
            $this.Logger.Write(("Could not read configuration file " + $this.jsonPath + "."),"Error")
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: Environment validation method
    Hidden [Boolean] ValidateEnvironment() {
        # Control variable
        [Boolean] $Return = $true
        # Determine if the outDir exists
        If ([Boolean](Test-Path -Path $this.outDir) -eq $true) {
            # Output directory exists; set the Excel path
            $this.excelPath = (Join-Path -Path $this.outDir -ChildPath ((Get-OZO8601Date -Time) + "-ozo-ad-manage-directory-objects.xlsx"))
        } Else {
            # Output directory does not exist; report
            $this.Logger.Write(("Output directory is invalid or inaccessible."),"Error")
            $Return = $false
        }
        # Try to get the domain information
        Try {
            $this.adDomain = Get-ADDomain -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Logger.Write("Failed to get AD domain information.")
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: Manage default objects method
    Hidden [Void] ManageDefaultObjects() {
        # Iterate over the objects in the Users container
        ForEach ($adObject in (Get-ADObject -Filter * -SearchBase $this.adDomain.UsersContainer)) {
            # Switch on object class
            Switch($adObject.ObjectClass) {
                "user" {
                    # Object class is user; move object to default users OU
                    Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $this.Json.ADDefaultObjectsOUs.Users
                }
                "group" {
                    # Object class is group; move object to default groups OU
                    Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $this.Json.ADDefaultObjectsOUs.Groups
                }
            }
        }
    }
    # METHODS: Report method
    Hidden [Void] Report() {
        
    }
}

Class OZOADContact {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADContact($adContact) {
        # Define parameters
        $Parameters = @{
            Name = ($adContact.FirstName + " " + $adContact.LastName)
            Type = "contact"
            Path = $adContact.Path
            DisplayName = ($adContact.FirstName + " " + $adContact.LastName)
            Mail = $adContact.EmailAddress
            TelephoneNumber = $adContact.TelephoneNumber
        }
        # Determine if the contact does not already exist
        If ([Boolean](Get-ADObject -Filter { Name -eq $adContact.Name -And ObjectClass -eq "contact" } -ErrorAction SilentlyContinue) -eq $false) {
            # Contact does not already exist; try to create it
            Try {
                New-ADObject @Parameters -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add(("Contact creation failed with " + $_))
                $this.Success = $false
            }
        } Else {
            # Contact exists; skip
            $this.Messages.Add("Contact already exists; skipping")
        }
    }
}

Class OZOADComputer {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADComputer($adComputer) {
        # Define parameters
        $Parameters = @{
            DisplayName = $adComputer.ComputerName
            Enabled = $true
            Name = $adComputer.ComputerName
            Path = $adComputer.Path
        }
        # Determine if the computer does not already exist
        If ([Boolean](Get-ADComputer -Identity $adComputer.Name) -eq $false) {
            # Computer object does not already exist; try to create it
            Try {
                New-ADComputer @Parameters -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add(("Computer creation failed with " + $_))
                $this.Success = $false
            }
        } Else {
            # Computer object already exists; skip
            $this.Messages.Add("Computer already exists; skipping")
        }
    }
}

Class OZOADGroup {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADGroup($adGroup) {
        # Define parameters
        $Parameters = @{}
        # Determine if the group does not already exist
        If ([Boolean](Get-ADGroup -Identity $adGroup.Name -ErrorAction SilentlyContinue) -eq $false) {
            # Group does not already exist; try to create it
            Try {
                New-ADGroup @Parameters -ErrorAction Stop
                # Success; iterate over parent groups
                ForEach ($Group in $adGroup.Groups) {
                    # Try to add group to parent group
                    Try {
                        Add-ADGroupMember -Identity $Group -Members $adGroup.Name -ErrorAction Stop
                        # Success
                    } Catch {
                        # Failure
                        $this.Messages.Add(("Adding to parent group " + $Group + " failed with " + $_))
                        $this.Success = $false
                    }
                }
            } Catch {
                # Failure
                $this.Messages.Add(("Group creation failed with " + $_))
                $this.Success = $false
            }
        } Else {
            # Group already exists; skip
            $this.Messages.Add("Group already exists; skipping")
        }
    }
}

Class OZOADGroupPolicyObject {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADGroupPolicyObject($adGPO) {
        # Determine if the GPO does not already exist
        If ([Boolean](Get-GPO -Name $adGPO.Name) -eq $false) {
            # GPO does not already exists; try to create it
            Try {
                New-GPO -Name $adGPO.Name -ErrorAction Stop
                # Success; iterate over the Links
                ForEach ($OU in $adGPO.Links) {
                    # Try to link GPO to OU
                    Try {
                        New-GPLink -Name $adGPO.Name -Target $OU -ErrorAction Stop
                        # Success
                    } Catch {
                        # Failure
                        $this.Messages.Add(("Linking GPO to " + $OU + " failed with " + $_))
                    }
                }
            } Catch {
                # Failure
                $this.Messages.Add(("Group Policy Object creation failed with " + $_))
                $this.Success = $false
            }
        } Else {
            # GPO does not already exist; skip
            $this.Messages.Add("Group Policy Object already exists; skipping")
        }
    }
}

Class OZOADOrganizationalUnit {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADOrganizationalUnit($adOU) {
        # Generate DN
        [String] $adOUDN = ("OU=" + $adOU.Name + "," + $adOU.Path)
        # Define parameters
        $Parameters = @{
            Name = $adOU.Name
            Path = $adOU.Path
        }
        # Determine if the OU does not already exist
        If ([Boolean](Get-ADOrganizationalUnit -Identity $adOUDN) -eq $false) {
            # OU does not already exist; try to create it
            Try {
                New-ADOrganizationalUnit @Parameters -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add(("Organizational Unit creation failed with " + $_))
                $this.Success = $false
            }
        } Else {
            # OU does not already exist; skip
            $this.Messages.Add("Organizational Unit already exists; skipping")
        }
    }
}

Class OZOADUser {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADUser($adUser,$adDomain) {
        # Define parameters
        $Parameters = @{
            Name = ($adUser.GivenName + " " + $adUser.Surname)
            SamAccountName = $adUser.SamAccountName
            UserPrincipalName = ($adUser.SamAccountName + "@" + $adDomain)
            AccountPassword = (ConvertTo-SecureString -AsPlainText -String $adUser.Password -Force)
            Enabled = $true
            Path = $adUser.Path
        }
        # Determine that the user does not already exist
        If ([Boolean](Get-ADUser -Identity $adUser.SamAccountName -ErrorAction SilentlyContinue) -eq $false) {
            # User does not already exist; try to create the user
            Try {
                New-ADUser @Parameters -ErrorAction Stop
                # Success; iterate through the groups
                ForEach ($adGroup in $adUser.Groups) {
                    # Try to create the group
                    Try {
                        Add-AdGroupMember -Identity $adGroup -Members $adUser.SamAccountName -ErrorAction Stop
                    } Catch {
                        # Failure
                        $this.Messages.Add(("Adding user to the " + $adGroup + " group failed with " + $_))
                    }
                }
            } Catch {
                # Failure
                $this.Messages.Add(("User creation failed with " + $_))
                $this.Success = $false
            }
        } Else {
            # User already exists; skip
            $this.Messages.Add("User already exists; skipping")
        }
    }
}

# Create a Main object
[OZOMain]::new($Configuration,$OutDir) | Out-Null
