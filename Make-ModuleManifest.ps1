$username = 'Administrator' # <== edit here
$modulename = 'IseThemeManager'
$manifestparams = @{
    Path = $(Join-Path "C:\Users\$username\Documents\WindowsPowerShell\Modules\IseThemeManager" ($modulename + '.psd1'));
    RootModule = $(Join-Path 'C:\Users\admin\Documents\WindowsPowerShell\Modules\IseThemeManager' ($modulename + '.psm1'));
    ModuleVersion = '2.1.0.0';
    Guid = $(New-Guid);
    Author = 'John Elliott';
    CompanyName = 'Clearpoint Software';
    Copyright = '(c) 2018 Clearpoint Software All rights reserved';
    Description = 'Manages PowerShell ISE color themes'
    PowerShellVersion = '3.0';
    PowerShellHostName = 'Windows PowerShell ISE Host';
    FunctionsToExport = 'Get-IseTheme';
    FileList = ($modulename + '.psm1');
    ProjectUri = 'https://github.com/ClearpointSoftware/IseThemeManager'
    Tags = 'Editor','ISE','Themes','Color Themes'
}
#New-ModuleManifest @manifestparams -PassThru
New-ModuleManifest @manifestparams
$uncomment = Get-Content -Path $manifestparams.Path | ? { $_ -notmatch '^#.*|^\s+#\s.*' }
$uncomment | % { $_.TrimEnd() } | ? { $_ -match '\S'} | Set-Content $manifestparams.Path
