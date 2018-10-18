
Function Dec2Hex {
    Param ([int]$indx)
    $a = $psxml.StorableColorTheme.Values.Color[$indx].A
    $r = $psxml.StorableColorTheme.Values.Color[$indx].R
    $g = $psxml.StorableColorTheme.Values.Color[$indx].G
    $b = $psxml.StorableColorTheme.Values.Color[$indx].B
    ($a,$r,$g,$b) | % {
        $hex = [System.Convert]::ToString($_,16)
        if ($hex.Length -eq 1) { $hex = "0$hex" }
        $hexcolor = $hexcolor + $hex
    }
    Return "#$hexcolor"
}

Function Load-IseTheme {
    Param ([string]$theme)
    $script:psxml = [xml](Get-Content -Path $theme)
    $keyarr = @($psxml.StorableColorTheme.Keys | Select -ExpandProperty string)
    $indx = 0
    foreach ($xmlkey in $keyarr) {
        switch -regex ($xmlkey) {
            '(^TokenColors|^ConsoleTokenColors|^XmlTokenColors)' {
                $node = $xmlkey -split '\\'
                if ([regex]::IsMatch($node[1],'\D')) {
                    $psISE.Options.($node[0]).item($node[1]) = (Dec2Hex $indx)
                }
            }
            default {
                $psISE.Options.($keyarr[$indx]) = (Dec2Hex $indx)
            }
        }
        $indx++
    }
    $psISE.Options.FontName = ($psxml.ChildNodes.FontFamily.GetValue(1))
    $psISE.Options.FontSize = ($psxml.ChildNodes.FontSize.GetValue(1))
    Write-Host "IseTheme: $($psxml.ChildNodes.Name.GetValue(1))"
    $xmlcfg.configuration.appSettings.CurrentThemeName.value = $($psxml.ChildNodes.Name.GetValue(1))
    $xmlcfg.Save($configfile)
}

Function Get-IseTheme {
    [CmdletBinding(DefaultParameterSetName='Load')]
    Param (
        [Parameter(ParameterSetName='Load',Position=0)][string]$Load,
        [Parameter(ParameterSetName='List')][switch]$List,
        [Parameter(ParameterSetName='Rename',Position=1)][switch]$Rename,
        [Parameter(ParameterSetName='Rename',Position=2,Mandatory=$true)][string]$CurrentName,
        [Parameter(ParameterSetName='Rename',Position=3,Mandatory=$false)][string]$NewName = $CurrentName,
        [Parameter(ParameterSetName='LineNumbers')][switch]$LineNumbers,
        [Parameter(ParameterSetName='Outlining')][switch]$Outlining,
        [Parameter(ParameterSetName='Toolbar')][switch]$Toolbar
    )

    $LoadThemeName = $Load
    $CurrentThemeName = $CurrentName
    $NewThemeName = $NewName
    $filetype = '.StorableColorTheme.ps1xml'
    $ThemeFiles = @(gci -Path $ModuleRoot | ? { $_.Name -match ".*\$filetype" })

    if ($ThemeFiles.Count -gt 0) {
        if ($PSCmdlet.ParameterSetName -eq 'List') {
            $ThemeFiles | Sort-Object -Property Name | % {
                if (($_.Name -replace $filetype,'') -eq $xmlcfg.configuration.appSettings.CurrentThemeName.value) {
                    Write-Host "$($xmlcfg.configuration.appSettings.CurrentThemeName.value) *"
                }
                else {
                    Write-Host $($_.Name -replace $filetype,'')
                }
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'Rename') {
            foreach ($t in $ThemeFiles) {
                if ($t.Name -eq ($NewThemeName + $filetype)) {
                    $NewThemeFile = $t.FullName
                    $xmlthm = [xml](Get-Content $NewThemeFile)
                    $xmlthm.StorableColorTheme.Name = $NewThemeName.Trim()
                    $xmlthm.Save($NewThemeFile)
                    $RenameItemParam = @{
                        Path = $NewThemeFile;
                        NewName = Join-Path (Split-Path $NewThemeFile -Parent) ($NewThemeName + $filetype);
                        Force = $true;
                    }
                    Rename-Item @RenameItemParam
                    break
                }
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'Load') {
            $LoadThemeFile = Join-Path $ModuleRoot ($LoadThemeName + $filetype)
            if (-not([System.IO.File]::Exists($LoadThemeFile))) {
                $ThemeFiles = $ThemeFiles | Sort-Object -Property LastWriteTime -Descending
                $LoadThemeFile = ($ThemeFiles[0]).FullName
                $notfound = 'IseThemeManager: ' +
                    'Could not find requested theme. Defaulting to most recently modified.'
                Write-Host $notfound -ForegroundColor DarkBlue -BackgroundColor Gray
            }
            Load-IseTheme $LoadThemeFile
        }
    }
    else {
        $notfound = 'IseThemeManager: ' +
            'No StorableColorTheme files available in IseThemeManager Module folder.'
        Write-Host $notfound -ForegroundColor DarkBlue -BackgroundColor Gray
    }
}

# $ModuleRoot = 'E:\Documents\WindowsPowerShell\Modules\IseThemeManager'
$ModuleRoot = $PSScriptRoot
$configfile = Join-Path $ModuleRoot 'IseThemeManager.config'
$xmlcfg = [xml](Get-Content $configfile)
