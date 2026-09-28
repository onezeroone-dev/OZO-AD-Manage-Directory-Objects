#Requires -Modules ImportExcel,OZO,OZOLogger -Version 5.1

<#PSScriptInfo
    .VERSION 0.0.2
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
    # PROPERTIES: Strings
    [String] $jsonPath  = $null
    [String] $excelPath = $null
    # PROPERTIES: PSCustomObjects
    [PSCustomObject] $adDomain  = $null
    [PSCustomObject] $Json      = $null
    [PSCustomObject] $ozoLogger = $null
    # PROPERTIES: PSCustomObject Lists
    [System.Collections.Generic.List[PSCustomObject]] $adOUs       = @()
    [System.Collections.Generic.List[PSCustomObject]] $adComputers = @()
    [System.Collections.Generic.List[PSCustomObject]] $adContacts  = @()
    [System.Collections.Generic.List[PSCustomObject]] $adGroups    = @()
    [System.Collections.Generic.List[PSCustomObject]] $adGPOs      = @()
    [System.Collections.Generic.List[PSCustomObject]] $adUsers     = @()
    # METHODS: Constructor method
    OZOMain($Configuration,$OutDir) {
        # Create a logger object
        $this.ozoLogger = (New-OZOLogger)
        # Log a process start message
        $this.ozoLogger.Write("Starting process.","Information")
        # Determine if the configuation is valid
        If (($this.ValidateConfiguration($Configuration) -And $this.ValidateEnvironment($OutDir)) -eq $true) {
            # Iterate through the JSON OU objects
            ForEach ($adOU in $this.Json.ADOrganizationalUnits) {
                # Add an OZOADOrganizationalUnit object to the adOUs list
                $this.adOUs.Add(([OZOADOU]::new($adOU)))
            }
            # Iterate through the JSON Computer objects
            ForEach ($adComputer in $this.Json.ADComputers) {
                # Add an OZOADComputer object to the adComputers list
                $this.adComputers.Add(([OZOADComputer]::new($adComputer)))
            }
            # Iterate through the JSON Contacts objects
            ForEach ($adContact in $this.Json.ADContacts) {
                # Add an OZOADOContact object to the adContacts list
                $this.adContacts.Add(([OZOADContact]::new($adContact)))
            }
            # Iterate through the JSON Group objects
            ForEach ($adGroup in $this.Json.ADGroups) {
                # Add an OZOADGroup object to the adGroups list
                $this.adGroups.Add(([OZOADGroup]::new($adGroup)))
            }
            # Iterate through the JSON Group Policy objects
            ForEach ($adGPO in $this.Json.ADGroupPolicyObjects) {
                # Add an OZOADGroupPolicyObject object to the adGPOs list
                $this.adGPOs.Add(([OZOADGPO]::new($adGPO)))
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
        $this.ozoLogger.Write(("Process complete."),"Information")
    }
    # METHODS: JSON validation method
    Hidden [Boolean] ValidateConfiguration($Configuration) {
        # Control variable
        [Boolean] $Return = $true
        # Determine if the JSON path is valid
        If ([Boolean](Test-Path -Path $Configuration -ErrorAction SilentlyContinue) -eq $true) {
            # JSON path is valid; try to get the JSON content
            Try {
                $this.Json = (Get-Content $Configuration -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
                # Success
            } Catch {
                # Failure
                $this.ozoLogger.Write(("Invalid JSON in " + $Configuration + "."),"Error")
                $Return = $false
            }
        } Else {
            # JSON path is invalid
            $this.ozoLogger.Write(("Could not read configuration file " + $Configuration + "."),"Error")
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: Environment validation method
    Hidden [Boolean] ValidateEnvironment($OutDir) {
        # Control variable
        [Boolean] $Return = $true
        # Determine if the outDir exists
        If ([Boolean](Test-Path -Path $OutDir -ErrorAction SilentlyContinue) -eq $true) {
            # Output directory exists; set the Excel path
            $this.excelPath = (Join-Path -Path $OutDir -ChildPath ((Get-OZO8601Date -Time) + "-ozo-ad-manage-directory-objects.xlsx"))
        } Else {
            # Output directory does not exist; report
            $this.ozoLogger.Write(("Output directory is invalid or inaccessible."),"Error")
            $Return = $false
        }
        # Try to get the domain information
        Try {
            $this.adDomain = Get-ADDomain -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.ozoLogger.Write("Failed to get AD domain information.")
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
        # Determine that at least one object was processed
        If (($this.adComputers + $this.adContacts + $this.adGroups + $this.adGPOs + $this.adOUs + $this.adUsers).Count -gt 0) {
            # At least one object was processed; Produce Excel output
            $this.adComputers | Select-Object -Property @{Name="Computer Name";Expression={$_.Name}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Export-Excel -WorksheetName "Computers" -Path $this.excelPath
            $this.adContacts | Select-Object -Property @{Name="Contact Name";Expression={$_.Name}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Export-Excel -WorksheetName "Contacts" -Path $this.excelPath
            $this.adGroups | Select-Object -Property @{Name="Group Name";Expression={$_.Name}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Export-Excel -WorksheetName "Groups" -Path $this.excelPath
            $this.adGPOs | Select-Object -Property @{Name="GPO Name";Expression={$_.Name}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Export-Excel -WorksheetName "GPOs" -Path $this.excelPath
            $this.adOUs | Select-Object -Property @{Name="OU";Expression={$_.adOUDN}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Export-Excel -WorksheetName "OUs" -Path $this.excelPath
            $this.adUsers | Select-Object -Property @{Name="User Name";Expression={$_.Name}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Export-Excel -WorksheetName "Users" -Path $this.excelPath
            # Determine if session is interactive
            If ([Environment]::UserInteractive -eq $true) {
                # Session is interactive; produce output for the operator
                $this.adComputers | Select-Object -Property @{Name="Computer Name";Expression={$_.Name}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Format-Table | Out-Host
                $this.adContacts | Select-Object -Property @{Name="Contact Name";Expression={$_.Name}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Format-Table | Out-Host
                $this.adGroups | Select-Object -Property @{Name="Group Name";Expression={$_.Name}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Format-Table | Out-Host
                $this.adGPOs | Select-Object -Property @{Name="GPO Name";Expression={$_.Name}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Format-Table | Out-Host
                $this.adOUs | Select-Object -Property @{Name="OU";Expression={$_.adOUDN}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Format-Table | Out-Host
                $this.adUsers | Select-Object -Property @{Name="User Name";Expression={$_.Name}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Format-Table | Out-Host
            }
            # Determine if Excel was created
            If ([Boolean](Test-Path -Path $this.excelPath -ErrorAction SilentlyContinue) -eq $true) {
                # Excel was created
                $this.ozoLogger.Write(("For additional information, please see " + $this.excelPath + "."),"Information")
            } Else {
                $this.ozoLogger.Write("No Excel results report generated.","Warning")
            }
        } Else {
            # No objects were processed
            $this.ozoLogger.Write("No objects were processed.","Warning")
        }
    }
}

Class OZOADComputer {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: Strings
    [String] $Name = $null
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADComputer($adComputer) {
        # Set properties
        $this.Name = $adComputer.ComputerName
        # Define parameters
        $Parameters = @{
            DisplayName = $adComputer.ComputerName
            Enabled = $true
            Name = $adComputer.ComputerName
            Path = $adComputer.Path
        }
        # Determine if the computer does not already exist
        If ([Boolean](Get-ADComputer -Identity $adComputer.Name -ErrorAction SilentlyContinue) -eq $false) {
            # Computer object does not already exist; try to create it
            Try {
                New-ADComputer @Parameters -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add(("Computer creation failed with error " + $_))
                $this.Success = $false
            }
        } Else {
            # Computer object already exists; skip
            $this.Messages.Add("Computer already exists; skipping")
        }
    }
}

Class OZOADContact {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: Strings
    [String] $Name = $null
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADContact($adContact) {
        # Local variables
        $this.Name = ($adContact.FirstName + " " + $adContact.LastName)
        # Define parameters
        $Parameters = @{
            Name = $this.Name
            Type = "contact"
            Path = $adContact.Path
            OtherAttributes = @{
                DisplayName = $this.Name
                Mail = $adContact.EmailAddress
                TelephoneNumber = $adContact.TelephoneNumber
            }
        }
        # Determine if the contact does not already exist
        If ([Boolean](Get-ADObject -Filter { Name -eq $this.Name -And ObjectClass -eq "contact" } -ErrorAction SilentlyContinue) -eq $false) {
            # Contact does not already exist; try to create it
            Try {
                New-ADObject @Parameters -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add(("Contact creation failed with error " + $_))
                $this.Success = $false
            }
        } Else {
            # Contact exists; skip
            $this.Messages.Add("Contact already exists; skipping")
        }
    }
}

Class OZOADGroup {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: Strings
    [String] $Name = $null
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADGroup($adGroup) {
        # Set properties
        $this.Name = $adGroup.Name
        # Define parameters
        $Parameters = @{
            Name = $adGroup.Name
            Path = $adGroup.Path
            GroupScope = $adGroup.Scope
        }
        # Try to get the group
        Try {
            Get-ADGroup -Identity $adGroup.Name -ErrorAction Stop | Out-Null
            # Success
        } Catch {
            # Failure; try to create the group
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
                $this.Messages.Add(("Group creation failed with error " + $_))
                $this.Success = $false
            }
        }
    }
}

Class OZOADGPO {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: Strings
    [String] $Name = $null
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADGPO($adGPO) {
        # Set properties
        $this.Name = $adGPO.Name
        # Determine if the GPO does not already exist
        If ([Boolean](Get-GPO -Name $adGPO.Name -ErrorAction SilentlyContinue) -eq $false) {
            # GPO does not already exist; try to create it
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
                        $this.Messages.Add(("Linking GPO to " + $OU + " failed with error " + $_))
                    }
                }
            } Catch {
                # Failure
                $this.Messages.Add(("Group Policy Object creation failed with error " + $_))
                $this.Success = $false
            }
        } Else {
            # GPO already exists
            $this.Messages.Add("Group Policy Object already exists; skipping")
        }
    }
}

Class OZOADOU {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: Strings
    [String] $adOUDN = $null
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADOU($adOU) {
        # Generate DN
        [String] $this.adOUDN = ("OU=" + $adOU.Name + "," + $adOU.Path)
        # Define parameters
        $Parameters = @{
            Name = $adOU.Name
            Path = $adOU.Path
        }
        # Try to get the Organizational Unit
        Try {
            Get-ADOrganizationalUnit -Identity $this.adOUDN -ErrorAction Stop | Out-Null
            # Success            
        } Catch {
            # Failure; create the OU
            Try {
                New-ADOrganizationalUnit @Parameters -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add(("Organizational Unit creation failed with error " + $_))
                $this.Success = $false
            }
            
        }
    }
}

Class OZOADUser {
    # PROPERTIES: Booleans
    [Boolean] $Success = $true
    # PROPERTIES: Strings
    [String] $Name = $null
    # PROPERTIES: String lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OZOADUser($adUser,$adDomainName) {
        # Set properties
        $this.Name = ($adUser.GivenName + " " + $adUser.Surname)
        # Define parameters
        $Parameters = @{
            Name = $this.Name
            SamAccountName = $adUser.SamAccountName
            UserPrincipalName = ($adUser.SamAccountName + "@" + $adDomainName)
            AccountPassword = (ConvertTo-SecureString -AsPlainText -String $adUser.Password -Force)
            Enabled = $true
            Path = $adUser.Path
        }
        # Try to get the user
        Try {
            Get-ADUser -Identity $adUser.SamAccountName -ErrorAction Stop | Out-Null
            # Success
        } Catch {
            # Failure; try to create the user
            Try {
                New-ADUser @Parameters -ErrorAction Stop
                # Success; iterate through the groups
                ForEach ($adGroup in $adUser.Groups) {
                    # Try to create the group
                    Try {
                        Add-AdGroupMember -Identity $adGroup -Members $adUser.SamAccountName -ErrorAction Stop
                    } Catch {
                        # Failure
                        $this.Messages.Add(("Adding user to the " + $adGroup + " group failed with error " + $_))
                    }
                }
            } Catch {
                # Failure
                $this.Messages.Add(("User creation failed with error " + $_))
                $this.Success = $false
            }
        }
    }
}

# Create a Main object
[OZOMain]::new($Configuration,$OutDir) | Out-Null
