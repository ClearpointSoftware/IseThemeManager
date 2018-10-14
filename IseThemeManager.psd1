@{
RootModule = 'C:\Users\<username>\Documents\WindowsPowerShell\Modules\IseThemeManager\IseThemeManager.psm1'
ModuleVersion = '2.1.0.0'
GUID = 'b4e22cd4-c703-433e-ae4b-1d3732fc51ec'
Author = 'John Elliott'
CompanyName = 'Clearpoint Software'
Copyright = '(c) 2018 Clearpoint Software All rights reserved'
Description = 'Manages PowerShell ISE color themes'
PowerShellVersion = '3.0'
PowerShellHostName = 'Windows PowerShell ISE Host'
FunctionsToExport = ('Get-IseTheme','Rename-IseTheme','List-IseTheme')
FileList = 'IseThemeManager.psm1'
PrivateData = @{
    PSData = @{
        Tags = @('Editor','ISE','Themes','Color Themes')
        ProjectUri = 'https://github.com/ClearpointSoftware/IseThemeManager'
    } # End of PSData hashtable
} # End of PrivateData hashtable
}
