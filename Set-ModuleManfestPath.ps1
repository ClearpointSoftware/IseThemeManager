
# must be run from PowerShell_ISE from location where IseThemeManager module is installed
$psdpath = Join-Path $PSScriptRoot 'IseThemeManager.psd1'
Write-Output "Attempting to update: $psdpath"
$ErrorActionPreference = 'Stop'
try {
    Update-ModuleManifest -Path $psdpath
    $message = 'Update Complete!'
}
catch {
    ($Error[0]).GetType().FullName
    ($Error[0]).Exception

    $message = @"

Could not update IseThemeManager.psd1

Please update the Path setting manually in IseThemeManager.psd1
to enable IseThemeManager to be discovered by Get-Module.

"@
}
$ErrorActionPreference = 'Continue'
Write-Output $message
