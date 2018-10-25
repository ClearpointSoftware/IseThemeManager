

Function Get-DefaultTheme {
    Param ([string]$descriptor)

    if ($ThemeFiles.Count -gt 0) {
        switch ($descriptor) {
            'Name' { $DefaultTheme = ($ThemeFiles[0]).Name -replace $filetype,'' }
            'FullPath' { $DefaultTheme =  ($ThemeFiles[0]).FullName }
        }
    }
    Return $DefaultTheme
}

Function Set-ActiveTheme {
    Param ([string]$ActiveTheme)

    $xmlcfg.configuration.appSettings.ActiveThemeName.value = $ActiveTheme
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
    Param ([string]$identifier)

    $rgx = '(?<version>\d+$)'
    $ScriptName =  [System.IO.Path]::GetFileNameWithoutExtension($psISE.CurrentFile.FullPath)
    $result = $ScriptName -match $rgx  
    $ScriptVersion = ($Matches.version)
    switch ($identifier) {
        DisplayName {
            $var = $ScriptName
        }
        BaseName {
            $var = $ScriptName.Substring(0, ($ScriptName.Length - $ScriptVersion.Length))
        }
        Number {
            $var = [int]$ScriptVersion
        }
        default { }
    }
    Return $var
}

Function Load-IseTheme {
    Param ([string]$theme)

    $script:psxml = [xml](Get-Content -Path $theme)
    $keyarr = @($psxml.StorableColorTheme.Keys | Select -ExpandProperty string)
    $indx = 0
    foreach ($xmlkey in $keyarr) {
        switch -Regex ($xmlkey) {
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
    [CmdletBinding(DefaultParameterSetName='CurrentFile')]
    Param (
        [Parameter(ParameterSetName='CurrentFile')][switch]$CurrentFile,
        [Parameter(ParameterSetName='Revision')][switch]$Revision,
        [Parameter(ParameterSetName='Fork')][switch]$Fork,
        [Parameter(ParameterSetName='RepositoryPath',Position=0,Mandatory=$true)][switch]$RepositoryPath,
        [Parameter(ParameterSetName='RepositoryPath',Position=1,Mandatory=$false)][string]$SetPath = $null
    )

    $ext = [System.IO.Path]::GetExtension($psISE.CurrentFile.FullPath)
    switch ($PSCmdlet.ParameterSetName) {
        'CurrentFile' {
            Write-Host "$([System.Environment]::NewLine)$($psISE.CurrentFile.FullPath)"
        }
        'Revision' {
            $OriginalFile = $psISE.CurrentFile
            $rev = (Get-ScriptVersion 'Number') + 5
            $NewFileName = "$(Get-ScriptVersion 'BaseName')$rev" + $ext
            $NewFilePath = Join-Path ($xmlcfg.configuration.appSettings.RepositoryPath.value) $NewFileName
            if ([System.IO.File]::Exists($NewFilePath)) {
                $message = " $NewFilePath already exists! " 
                Write-Host $message -ForegroundColor DarkBlue -BackgroundColor DarkYellow
            }
            else {
                $NewFileContent = $psISE.CurrentFile.Editor.Text
                if ($NewFileContent.length -gt 0) {
                    try {
                        $OriginalFile.Save()
                        $NewFile = $psISE.CurrentPowerShellTab.Files.Add()
                        $NewFile.Editor.Text = $NewFileContent
                        $NewFile.Editor.SetCaretPosition(1,1)
                        $NewFile.SaveAs($NewFilePath)
                        $psISE.CurrentPowerShellTab.Files.Remove($OriginalFile,$false)
                    }
                    catch {
                        $message = ' Repository save action failed ' 
                        Write-Host $message -ForegroundColor White -BackgroundColor DarkMagenta
                    }
                }
            }
        }
        'Fork' {
            $NewFileName = "$(Get-ScriptVersion 'DisplayName')F_0" + $ext
            $NewFilePath = Join-Path ($xmlcfg.configuration.appSettings.RepositoryPath.value) $NewFileName
            if ([System.IO.File]::Exists($NewFilePath)) {
                $message = " $NewFilePath already exists! " 
                Write-Host $message -ForegroundColor DarkBlue -BackgroundColor DarkYellow
            }
            else {
                $NewFileContent = $psISE.CurrentFile.Editor.SelectedText
                if ($NewFileContent.length -gt 0) {
                    $NewFile = $psISE.CurrentPowerShellTab.Files.Add()
                    $NewFile.Editor.Text = $NewFileContent
                    $NewFile.Editor.SetCaretPosition(1,1)
                    $ErrorActionPreference = 'Stop'
                    try {
                        $NewFile.SaveAs($NewFilePath)
                    }
                    catch {
                        $message = ' Repository save action failed ' 
                         Write-Host $message -ForegroundColor White -BackgroundColor DarkMagenta
                    }
                    $ErrorActionPreference = 'Continue'
                }
                else {
                    $message = ' Select the code you wish to fork '
                    Write-Host $message -ForegroundColor DarkBlue -BackgroundColor DarkYellow
                }
            }
        }
        'RepositoryPath' {
            if ([System.String]::IsNullOrEmpty($SetPath)) {
                $message = "$([System.Environment]::NewLine)$($xmlcfg.configuration.appSettings.RepositoryPath.value)"
                Write-Output $message
            }
            elseif (Test-Path -Path $SetPath -PathType Container) {
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
    [CmdletBinding(DefaultParameterSetName='List')]
    Param (
        [Parameter(ParameterSetName='Load')][string]$Load,
        [Parameter(ParameterSetName='Rename',Position=0,Mandatory=$true)][switch]$Rename,
        [Parameter(ParameterSetName='Rename',Position=1,Mandatory=$true)][string]$CurrentName,
        [Parameter(ParameterSetName='Rename',Position=2,Mandatory=$true)][string]$NewName
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
                    $themename = $_.Name -replace $filetype,''
                    if ($themename -eq $xmlcfg.configuration.appSettings.ActiveThemeName.value) {
                        $themename = "$themename *"
                    }
                    Write-Host $themename
                }
            }
            Load {
                $LoadThemeFile = Join-Path $ModuleRoot ($LoadThemeName + $filetype)
                if (-not([System.IO.File]::Exists($LoadThemeFile))) {
                    $LoadThemeFile = Get-DefaultTheme 'FullPath'
                    $message = ' Could not find requested theme. Defaulting to most recently modified. '
                    Write-Host $message -ForegroundColor DarkBlue -BackgroundColor DarkYellow
                }
                Load-IseTheme $LoadThemeFile
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
            default { }
        }
    }
    else {
        $message = ' IseThemeManager could not find any StorableColorTheme files '
        Write-Host $notfound -ForegroundColor DarkBlue -BackgroundColor DarkYellow
    }
}

