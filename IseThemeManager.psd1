@{
RootModule = 'IseThemeManager.psm1'
ModuleVersion = '2.1.1.0'
GUID = '021fa1ec-b7ca-4813-9a05-0c56fc52ea86'
Author = 'John Elliott'
CompanyName = 'Clearpoint Software'
Copyright = '(c) 2018 Clearpoint Software. All rights reserved'
Description = 'Manages PowerShell ISE color themes'
PowerShellVersion = '3.0'
PowerShellHostName = 'Windows PowerShell ISE Host'
FunctionsToExport = 'Get-IseTheme', 'Get-IseVersion'
FileList = 'IseThemeManager.psm1', 'IseThemeManager.config'
PrivateData = @{
    PSData = @{
        Tags = 'Editor', 'ISE', 'Themes', 'Color Themes'
        ProjectUri = 'https://github.com/ClearpointSoftware/IseThemeManager'
    } # End of PSData hashtable
} # End of PrivateData hashtable
}
