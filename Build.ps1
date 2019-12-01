$ModuleName     = "ExchangeTrackingGUI"
$ModuleGuid     = "7dbaff80-3574-4797-9e57-339e228a3eab"
$Year           = (Get-Date).Year
$ModuleVersion  = "1.0.0"
$ReleaseNotes  = "
* Initial release
"
$AllFunctions   = @( Get-ChildItem -Path $PSScriptRoot\source\*.ps1 -ErrorAction SilentlyContinue -Recurse | Sort-Object)
$BuildDir       = "C:\Projects"
$CombinedModule = "$BuildDir\Build\$moduleName\$ModuleName.psm1"
$ManifestFile   = "$BuildDir\Build\$moduleName\$ModuleName.psd1"

# Init module script
[string]$CombinedResources = ""

# Prepare Functions
Foreach ($Function in @($AllFunctions)) {
    Write-Output "Add Function: $Function"
    Try {
        $CombinedResources += (Get-Content $Function -Raw) + "`n"
    }
    Catch {
        throw $_
    }
}

# Prepare Manifest
$ManifestDefinition = @"
@{

# Script module or binary module file associated with this manifest.
RootModule = '$ModuleName.psm1'

# Version number of this module.
ModuleVersion = '$ModuleVersion'

# Supported PSEditions
CompatiblePSEditions = @('5.1')

# ID used to uniquely identify this module
GUID = '$ModuleGuid'

# Author of this module
Author = 'Pavel Andreev'

# Company or vendor of this module
CompanyName = ''

# Copyright statement for this module
Copyright = '(c) $Year Pavel Andreev. All rights reserved.'

# Description of the functionality provided by this module
Description = 'GUI for Get-MessageTrackingLog command'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '5.1'

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @('Get-MessageTrackingGUI')

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = @('Exchange')

        # A URL to the license for this module.
        LicenseUri = 'https://github.com/pvs043/ExchangeTrackingGUI/blob/master/LICENSE'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/pvs043/ExchangeTrackingGUI'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        ReleaseNotes = '$ReleaseNotes'
    } # End of PSData hashtable

} # End of PrivateData hashtable

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''
}
"@

# Create Build dir
If (Test-Path -Path "$BuildDir\Build") { Remove-Item -Path "$BuildDir\Build" -Recurse -Force}
$null = New-Item -ItemType Directory -Path "$BuildDir\Build\$ModuleName"

# Build module from sources
Set-Content -Path $CombinedModule -Value $CombinedResources
Set-Content -Path $ManifestFile -Value $ManifestDefinition

# Add artefacts
Copy-Item -Path "$PSScriptRoot\README.md" -Destination "$BuildDir\Build\$moduleName\Readme.md" -Force
Copy-Item -Path "$PSScriptRoot\LICENSE" -Destination "$BuildDir\Build\$moduleName\LICENSE" -Force
Copy-Item -Path "$PSScriptRoot\source\plus-24.png" -Destination "$BuildDir\Build\$moduleName\plus-24.png" -Force
Copy-Item -Path "$PSScriptRoot\source\minus-24.png" -Destination "$BuildDir\Build\$moduleName\minus-24.png" -Force
