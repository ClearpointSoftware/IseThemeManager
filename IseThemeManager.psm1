
Function Get-DefaultTheme {
    Param (
        [switch]$Name,
        [switch]$File
    )
    if ($ThemeFiles.Count -gt 0) {
        if ($Name) {
            $DefaultTheme = ($ThemeFiles[0]).Name -replace $filetype,''
        }
        elseif ($File) {
            $DefaultTheme =  ($ThemeFiles[0]).FullName
        }
    }
    Return $DefaultTheme
}

Function Set-ActiveTheme {
    Param ([string]$ThemeName)
    $xmlcfg.configuration.appSettings.ActiveThemeName.value = $ThemeName
    $xmlcfg.Save($configfile)
}

Function Dec2Hex {
    Param ([int]$indx)
    ($psxml.StorableColorTheme.Values.Color[$indx].A,
        $psxml.StorableColorTheme.Values.Color[$indx].R,
        $psxml.StorableColorTheme.Values.Color[$indx].G,
        $psxml.StorableColorTheme.Values.Color[$indx].B) | % {
            $hex = [System.Convert]::ToString($_,16)
            if ($hex.Length -eq 1) {
                $hex = "0$hex"
            }
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
    Set-ActiveTheme -ThemeName $($psxml.ChildNodes.Name.GetValue(1))
    Write-Host "IseTheme: $($psxml.ChildNodes.Name.GetValue(1))"
}

Function Get-IseTheme {
    [CmdletBinding(DefaultParameterSetName='Load')]
    Param (
        [Parameter(ParameterSetName='Load',Position=0)][string]$Load,
        [Parameter(ParameterSetName='List')][switch]$List,
        [Parameter(ParameterSetName='Rename',Position=1)][switch]$Rename,
        [Parameter(ParameterSetName='Rename',Position=2,Mandatory=$true)][string]$CurrentName,
        [Parameter(ParameterSetName='Rename',Position=3,Mandatory=$true)][string]$NewName,
        [Parameter(ParameterSetName='LineNumbers')][switch]$LineNumbers,
        [Parameter(ParameterSetName='Outlining')][switch]$Outlining,
        [Parameter(ParameterSetName='Toolbar')][switch]$Toolbar
    )

    $LoadThemeName = $Load
    $CurrentThemeName = $CurrentName
    $NewThemeName = $NewName

    $script:ThemeFiles = @(gci -Path $ModuleRoot | ? { $_.Name -match ".*\$filetype" }) |
        Sort-Object -Property LastWriteTime -Descending
    $script:xmlcfg = [xml](Get-Content $configfile)

    if ($ThemeFiles.Count -gt 0) {
        if ($PSCmdlet.ParameterSetName -eq 'List') {
            $ThemeFiles | Sort-Object -Property Name | % {
                if (($_.Name -replace $filetype,'') -eq $xmlcfg.configuration.appSettings.ActiveThemeName.value) {
                    Write-Host "$($xmlcfg.configuration.appSettings.ActiveThemeName.value) *"
                }
                else {
                    Write-Host $($_.Name -replace $filetype,'')
                }
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'Rename') {
            foreach ($t in $ThemeFiles) {
                if ($t.Name -eq ($CurrentThemeName + $filetype)) {
                    $CurrentThemeFile = $t.FullName
                    $xmlthm = [xml](Get-Content $CurrentThemeFile)
                    $xmlthm.StorableColorTheme.Name = $NewThemeName
                    $xmlthm.Save($CurrentThemeFile)
                    $RenameItemParam = @{
                        Path = $CurrentThemeFile;
                        NewName = Join-Path $ModuleRoot ($NewThemeName + $filetype);
                        Force = $true;
                    }
                    Rename-Item @RenameItemParam
                    if ($xmlcfg.configuration.appSettings.ActiveThemeName.value -eq $CurrentThemeName) {
                        Set-ActiveTheme -ThemeName $NewThemeName
                    } 
                    break
                }
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'Load') {
            $LoadThemeFile = Join-Path $ModuleRoot ($LoadThemeName + $filetype)
            if (-not([System.IO.File]::Exists($LoadThemeFile))) {
                $LoadThemeFile = (Get-DefaultTheme -File)
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

$ModuleRoot = $PSScriptRoot
$filetype = '.StorableColorTheme.ps1xml'
$configfile = Join-Path $ModuleRoot 'IseThemeManager.config'
