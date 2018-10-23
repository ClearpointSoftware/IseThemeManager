﻿

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

Function Convert-ARGB {
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

Function Get-ScriptVersion {
    [CmdletBinding(DefaultParameterSetName='DisplayName')]
    Param (
        [Parameter(ParameterSetName='DisplayName')][switch]$DisplayName,
        [Parameter(ParameterSetName='BaseName')][switch]$BaseName,
        [Parameter(ParameterSetName='Number')][switch]$Number
    )
    $rgx = '(?<version>\d+$)' # $Matches returns 0 if no version digit
    $ScriptName =  [System.IO.Path]::GetFileNameWithoutExtension($psISE.CurrentFile.FullPath)
    $result = $ScriptName -match $rgx  
    $ScriptVersion = ($Matches.version) 
    switch -Regex ($PSCmdlet.ParameterSetName) {
        'DisplayName' {
            $var = $ScriptName
        }
        'BaseName' {
            $var = $ScriptName -replace($ScriptVersion,'')
        }
        'Number' {
            $var = [int]$ScriptVersion
        }
    }
    Return $var
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
                    $psISE.Options.($node[0]).item($node[1]) = (Convert-ARGB $indx)
                }
            }
            default {
                $psISE.Options.($keyarr[$indx]) = (Convert-ARGB $indx)
            }
        }
        $indx++
    }
    $psISE.Options.FontName = ($psxml.ChildNodes.FontFamily.GetValue(1))
    $psISE.Options.FontSize = ($psxml.ChildNodes.FontSize.GetValue(1))
    Set-ActiveTheme -ThemeName "$($psxml.ChildNodes.Name.GetValue(1))"
    Write-Host "IseTheme: $($psxml.ChildNodes.Name.GetValue(1))"
}

Function Get-IseVersion {
    [CmdletBinding(DefaultParameterSetName='CurrentFilePath')]
    Param (
        [Parameter(ParameterSetName='CurrentFilePath')][switch]$CurrentFilePath,
        [Parameter(ParameterSetName='Revision')][switch]$Revision,
        [Parameter(ParameterSetName='Fork')][switch]$Fork,
        [Parameter(ParameterSetName='RepositoryPath',Position=0,Mandatory=$true)][switch]$RepositoryPath,
        [Parameter(ParameterSetName='RepositoryPath',Position=1,Mandatory=$false)][string]$SetPath
    )
    $CurrentFile = $psISE.CurrentFile
    $ext = [System.IO.Path]::GetExtension($psISE.CurrentFile.FullPath)
    switch -Regex ($PSCmdlet.ParameterSetName) {
        'CurrentFilePath' {
            Write-Host "$([System.Environment]::NewLine)$($CurrentFile.FullPath)"
        }
        'Revision|Fork' {
            if ($_ -match 'Revision') {
                $rev = (Get-ScriptVersion -Number) + 5
                $NewFileName = "$(Get-ScriptVersion -BaseName)$rev" + $ext
                $NewFileContent = $CurrentFile.Editor.Text
            }
            elseif ($_ -match 'Fork') {
                $NewFileName = "$(Get-ScriptVersion -DisplayName)F_0" + $ext
                $NewFileContent = $CurrentFile.Editor.SelectedText
            }
            if ($NewFileContent.length -gt 0) {
                $NewFile = $psISE.CurrentPowerShellTab.Files.Add()
                $NewFile.Editor.Text = $NewFileContent
                $NewFile.Editor.SetCaretPosition(1,1)
                $ErrorActionPreference = 'Stop'
                try {
                    $NewFilePath = Join-Path ($xmlcfg.configuration.appSettings.RepositoryPath.value) $NewFileName
                }
                catch {
                    $NewFilePath = $null
                    $message = ' Invalid Repository Path '
                    Write-Host $message -ForegroundColor DarkBlue -BackgroundColor DarkYellow
                }
                try {
                    if ([System.IO.File]::Exists($NewFilePath)) {
                        $message = " $NewFilePath already exists! " 
                        Write-Host $message -ForegroundColor DarkBlue -BackgroundColor DarkYellow
                    }
                    elseif ($NewFilePath) {
                        $NewFile.SaveAs($NewFilePath)
                        if ($_ -match 'Revision') {
                            $CurrentFile.Save()
                            $psISE.CurrentPowerShellTab.Files.Remove($CurrentFile,$true)
                        }
                    }
                }
                catch {
                    $message = ' Error: could not save to Repository  ' 
                    Write-Host $message -ForegroundColor White -BackgroundColor DarkMagenta
                }
                $ErrorActionPreference = 'Continue'
            }
            else {
                $message = ' Select the code you wish to fork '
                Write-Host $message -ForegroundColor DarkBlue -BackgroundColor DarkYellow
            }
        }
        'RepositoryPath' {
            if ($RepositoryPath -and (-not($SetPath))) {
                $message = "$([System.Environment]::NewLine)$($xmlcfg.configuration.appSettings.RepositoryPath.value)"
                Write-Output $message
            }
            elseif ($RepositoryPath -and (Test-Path $SetPath)) {
                $xmlcfg.configuration.appSettings.RepositoryPath.value = $SetPath
                $xmlcfg.Save($configfile)
            }
            else {
                $message = ' Invalid path. Please try again '
                Write-Host $message -ForegroundColor DarkBlue -BackgroundColor DarkYellow
            }
        }
        default { }
    }
}

Function Get-IseTheme {
    [CmdletBinding(DefaultParameterSetName='Load')]
    Param (
        [Parameter(ParameterSetName='Load',Position=0)][string]$Load,
        [Parameter(ParameterSetName='List')][switch]$List,
        [Parameter(ParameterSetName='Rename',Position=1,Mandatory=$true)][switch]$Rename,
        [Parameter(ParameterSetName='Rename',Position=2,Mandatory=$true)][string]$CurrentName,
        [Parameter(ParameterSetName='Rename',Position=3,Mandatory=$true)][string]$NewName
    )
    $script:ModuleRoot = Split-Path "$PSCommandPath"
    $script:filetype = '.StorableColorTheme.ps1xml'
    $script:configfile = Join-Path $ModuleRoot 'IseThemeManager.config'
    $script:ThemeFiles = @(gci -Path $ModuleRoot | ? { $_.Name -match ".*\$filetype" }) |
        Sort-Object -Property LastWriteTime -Descending
    $script:xmlcfg = [xml](Get-Content $configfile)

    $LoadThemeName = $Load
    $CurrentThemeName = $CurrentName
    $NewThemeName = $NewName

    if ($ThemeFiles.Count -gt 0) {
        switch ($PSCmdlet.ParameterSetName) {
            List {
                $psISE.CurrentPowerShellTab.ConsolePane.Clear()
                $ThemeFiles | Sort-Object -Property Name | % {
                    if (($_.Name -replace $filetype,'') -eq $xmlcfg.configuration.appSettings.ActiveThemeName.value) {
                        Write-Host "$($xmlcfg.configuration.appSettings.ActiveThemeName.value) *"
                    }
                    else {
                        Write-Host $($_.Name -replace $filetype,'')
                    }
                }
            }
            Rename {
                foreach ($t in $ThemeFiles) {
                    if ($t.Name -eq ($CurrentThemeName + $filetype)) {
                        $CurrentThemeFile = $t.FullName
                        $themexml = [xml](Get-Content $CurrentThemeFile)
                        $themexml.StorableColorTheme.Name = $NewThemeName
                        $themexml.Save($CurrentThemeFile)
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
            Load {
                $LoadThemeFile = Join-Path $ModuleRoot ($LoadThemeName + $filetype)
                if (-not([System.IO.File]::Exists($LoadThemeFile))) {
                    $LoadThemeFile = (Get-DefaultTheme -File)
                    $message = ' Could not find requested theme. Defaulting to most recently modified. '
                    Write-Host $message -ForegroundColor DarkBlue -BackgroundColor DarkYellow
                }
                Load-IseTheme $LoadThemeFile
            }
            default { }
        }
    }
    else {
        $message = ' IseThemeManager could not find any StorableColorTheme files '
        Write-Host $notfound -ForegroundColor DarkBlue -BackgroundColor DarkYellow
    }
}

