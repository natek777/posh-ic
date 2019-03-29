# License Export V1.4  LicenseExport_1.4-xlsx.ps1
# Made by Nathanael Kaseman, for universal running v1.1, fixed license 1.2, added more user data v1.3, Added Server A detection v1.4
# This script will export User Station and License information from the registry of a CIC server
# While no performance issues are expected, please always run this on the backup server. 

#Requires -RunAsAdministrator
#Requires -Version 3.0

#Dor issues
#You get an error “The WinRM client cannot process the request. Basic authentication is currently disabled in the client configuration. Change the client configuration and try the request again.”
#Open a PowerShell prompt as Administrator and run:
#winrm set winrm/config/client/auth '@{Basic="true"}'
#winrm set winrm/config/client '@{AllowUnencrypted="true"}'



[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

#Adds Modules from script path
Import-Module $PSScriptRoot\Modules\ImportExcel
#installs module if it did not load
if (-not (Get-Module -Name "ImportExcel")) {
    Write-Host("Module missing. We are running the following command to add it- Install-Module ImportExcel -Scope CurrentUser -Force")
    Install-Module ImportExcel -Scope CurrentUser -Force
}
# Uncommment for Interactive save location
$SaveFileDialog = New-Object windows.forms.savefiledialog 
# Uncommment for Interactive save location
#$SaveFileDialog.Filter = "Excel Workbook|*.xlsx" 
#$SaveFileDialog.ShowDialog()
# Comment out for Interactive filename
$SaveFileDialog.FileName="D:\Scripts\UserLicenseStationExport.xlsx"
# Debuging
# Write-Host $SaveFileDialog.FileName

$Export = Invoke-Command -ComputerName $env:computername -ScriptBlock {
    function Open-Registry ($ComputerName) {
        Try {
            #From Google
            #$Reg = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 0)
            #old remote - working
            #$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
            #This is local based, it also works
            $Reg = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Default)
            #$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
        }
        Catch {
            return "Unable to connect to the Registry Values! Is this a PureConnect Server?";
        }
        return $Reg
    }
    function Get-SiteName () {
        Try {
            $RegKey = $Reg.OpenSubKey("SOFTWARE\\Wow6432Node\\Interactive Intelligence\\EIC\\Directory Services\\Root\\")
            Try { $value = $RegKey.GetValue("SITE") -replace "^\\", ""}
            Catch { $value = "" };
            return $value;
        }
        Catch {
            return "Unable to connect to the Registry Values! Is this a PureConnect Server?";
        }
    }
    function Get-ServerAName () {
        Try {
            $RegKey = $Reg.OpenSubKey("SOFTWARE\\Wow6432Node\\Interactive Intelligence\\EIC\\Directory Services\\Root\\")
            Try { $value = $RegKey.GetValue("SERVER") -replace "^\\", ""}
            Catch { $value = "" };
            return $value;
        }
        Catch {
            return "Unable to connect to the Registry Values! Is this a PureConnect Server?";
        }
    }
    # Change CustomerSite to the Sitename
    $Reg = Open-Registry -ComputerName $env:computername
    $SiteName = Get-SiteName -Registry $Reg
    #Detect server A name with all the coresponding REG Keys(including attendant) 
    $ServerAName = Get-ServerAName -Registry $Reg
    # Changed UserReg to Production\AdminConfig\Users from Production\\Users
    $UserReg = "HKLM:\software\Wow6432Node\Interactive Intelligence\EIC\Directory Services\Root\" + $SiteName + "\Production\Users"
    $Users = Get-ChildItem $UserReg
    $UserLicReg = "HKLM:\software\Wow6432Node\Interactive Intelligence\EIC\Directory Services\Root\" + $SiteName + "\Production\AdminConfig\UserInfo\"
    $UsersLic = Get-ChildItem $UserLicReg
    # Use ServeraName to detect Site Name and the Server A
    $StationReg = "HKLM:\software\Wow6432Node\Interactive Intelligence\EIC\Directory Services\Root\"+$ServerAName+"\Workstations\"
    $Stations = Get-ChildItem $StationReg
    # Old Station Function no automatic detection needed
    #$StationReg = "HKLM:\software\Wow6432Node\Interactive Intelligence\EIC\Directory Services\Root\" + $SiteName + "\Production\AdminConfig\StationInfo\"
    #$Stations = Get-ChildItem $StationReg
    $Export = @{}
    $Export["Users"] = @()
    $Export["Stations"] = @()
    $StationLicenses = New-Object PSCustomObject
    $UserLicenses = New-Object PSCustomObject
    $TotalLicenses = New-Object PSCustomObject

    $StationLicenses | Add-Member -NotePropertyName LicenseType -NotePropertyValue "Station" -Force
    $UserLicenses | Add-Member -NotePropertyName LicenseType -NotePropertyValue "User" -Force
    $TotalLicenses | Add-Member -NotePropertyName LicenseType -NotePropertyValue "Total" -Force

    $Licenses = $UsersLic | ForEach-Object {
        $Username = $_.Name.Substring($_.Name.LastIndexOf("\") + 1, $_.Name.Length - $_.Name.LastIndexOf("\") - 1)
        (Get-ItemProperty ($UserLicReg + $Username) -ErrorAction SilentlyContinue).'Counted Licenses'
    } | Sort-Object -Unique
    $Licenses += $Stations | ForEach-Object {
        $StationName = $_.Name.Substring($_.Name.LastIndexOf("\") + 1, $_.Name.Length - $_.Name.LastIndexOf("\") - 1)
        (Get-ItemProperty ($StationReg + $StationName) -ErrorAction SilentlyContinue).'Counted Licenses'
    } | Sort-Object -Unique

    foreach ($LicenceType in $Licenses) {
        $UserLicenses | Add-Member -NotePropertyName $LicenceType -NotePropertyValue 0 -Force
        $StationLicenses | Add-Member -NotePropertyName $LicenceType -NotePropertyValue 0 -Force
        $TotalLicenses | Add-Member -NotePropertyName $LicenceType -NotePropertyValue 0 -Force
    }
# List all the Users and add the property and license values
    foreach ($User in $Users) {
        $Username = $User.Name.Substring($User.Name.LastIndexOf("\") + 1, $User.Name.Length - $User.Name.LastIndexOf("\") - 1)
        $UserLicense = Get-ItemProperty ($UserLicReg + $Username) -ErrorAction SilentlyContinue
        $UserInfo = % {
            $object = [ordered]@{
                # add more user items here if desired (Check the registry for the value mapping)
                UserName           = $Username
                FirstName          = ($User.GetValue("givenName") -join ", ")
                LastName           = ($User.GetValue("surname") -join ", ")
                DisplayName        = ($User.GetValue("displayName") -join ", ")
                NTDomainName       = ($User.GetValue("NT Domain User") -join ", ")
                ICPrivacy          = ($User.GetValue("Alias") -join ", ")
                ClientTemplate	   = ($User.GetValue("ClientTemplateConfiguration") -join ", ")                
                BusinessAddress    = ($User.GetValue('officeLocation') -join ", ")
		        Department         = ($User.GetValue("departmentName") -join ", ")
                Title              = ($User.GetValue("Workgroups") -join ", ")
                EmailAddress       = ($User.GetValue("emailAddress") -join ", ")
                DefaultWorkstation = ($User.GetValue("Default Workstation") -join ", ")
                Extension          = ($User.GetValue("Extension") -join ", ")
                OutboundANI        = ($User.GetValue("Outbound ANI") -join ", ")
                FaxCapability      = ($User.GetValue('Fax Capability') -join ", ")
                Role               = ($User.GetValue('Role') -join ", ")
                Workgroups         = ($User.GetValue("Workgroups") -join ", ")
                Skills             = ($User.GetValue('Skills') -join ", ")
                Utilization        = ($User.GetValue("Agent Media % Utilization") -join ", ")	
                AutoAnswer         = ($User.GetValue("Auto-Answer Call") -join ", ")
                AutoAnswerNonACD   = ($User.GetValue("Auto-Answer Non-ACD Call") -join ", ")
                MWIEnabled         = ($User.GetValue('MWI Enabled') -join ", ")
                MasterAdmin        = ($User.GetValue('Master Admin') -join ", ")
                PublishHandlers    = ($User.GetValue('PublishHandlers') -join ", ")
                ManageHandlers     = ($User.GetValue("ManageHandlers") -join ", ")
                DebugHandlers      = ($User.GetValue('Debug Handlers') -join ", ")
                RemoteControl      = ($User.GetValue('RemoteControl') -join ", ")

            }
            New-Object PSObject -Property $object
        }
        ForEach ($LicenseType in $Licenses) {
            if ($UserLicense.'Counted Licenses' -contains $LicenseType) {
                $UserInfo | Add-Member -NotePropertyName $LicenseType -NotePropertyValue "X" -Force
                $UserLicenses.$LicenseType += 1
                $TotalLicenses.$LicenseType += 1
            }
            else {
                $UserInfo | Add-Member -NotePropertyName $LicenseType -NotePropertyValue "" -Force
            }
        }
        $Export["Users"] += $UserInfo
    }
# List all the stations and add the names and license values
    foreach ($Station in $Stations) {
        $StationName = $Station.Name.Substring($Station.Name.LastIndexOf("\") + 1, $Station.Name.Length - $Station.Name.LastIndexOf("\") - 1)
        $StationLicense = Get-ItemProperty ($StationReg + $StationName) -ErrorAction SilentlyContinue
        $StationInfo = % {
            $object = [ordered]@{
                #Add more statin Items Here if needed (Check the registry for the value mapping)
                StationName  =  $StationName
                Extension    =  ($Station.GetValue("Extension") -join ", ")
                Line         =  ($Station.GetValue("Line") -join ", ")
                MacAddress   =  ($Station.GetValue("MAC Address") -join ", ")
                ManagedIPPhone =($Station.GetValue("Managed IP Phone") -join ", ")
                StationType  =  ($Station.GetValue("Station Type") -join ", ")
                Active       =  ($Station.GetValue("Active") -join ", ")
                Licenses     =  ($Station.GetValue("Counted Licenses") -join ", ")
                LicenseAlloc =  ($Station.GetValue("License Allocation Method") -join ", ")
            }
            New-Object PSObject -Property $object
        }
        ForEach ($LicenseType in $Licenses) {
            if ($StationLicense.'Counted Licenses' -contains $LicenseType) {
                $StationInfo | Add-Member -NotePropertyName $LicenseType -NotePropertyValue "X" -Force
                $StationLicenses.$LicenseType += 1
                $TotalLicenses.$LicenseType += 1
            }
            else {
                $StationInfo | Add-Member -NotePropertyName $LicenseType -NotePropertyValue "" -Force
            }
        }
        $Export["Stations"] += $StationInfo
    }
    $Export["Total"] = @()
    $Export["Total"] = [Array]$UserLicenses + [Array]$StationLicenses + [Array]$UserLicenses
    $Export
}
Try {
    $Export["Total"] | Export-Excel -Path $SaveFileDialog.FileName -WorkSheetname "Total" -BoldTopRow -AutoSize 
}
Catch {
    return "Error The totals do not exist! Is this a PureConnect Server?";
}
Try {
    $Export["Users"] | Export-Excel -Path $SaveFileDialog.FileName -WorkSheetname "Users" -BoldTopRow -AutoSize
}
Catch {
    return "Error The Users do not exist! Is this a PureConnect Server?";
}
Try {
    $Export["Stations"] | Export-Excel -Path $SaveFileDialog.FileName -WorkSheetname "Stations" -BoldTopRow -AutoSize
}
Catch {
    return "Error The Stations do not exist! Is this a PureConnect Server?";
}
