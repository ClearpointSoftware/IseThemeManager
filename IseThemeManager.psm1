


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
    Param ([string]$ActiveThemeFile)

    $ActiveThemeName = ((Split-Path $ActiveThemeFile -Leaf) -split '\.')[0]
    if ($psxml.StorableColorTheme.Name -ne $ActiveThemeName) {
        $psxml.StorableColorTheme.Name = $ActiveThemeName
        $psxml.Save($ActiveThemeFile)
    }
    $xmlcfg.configuration.appSettings.ActiveThemeName.value = $ActiveThemeName
    $xmlcfg.Save($configfile)
    Return $ActiveThemeName
}

Function Write-Message {
    Param ([string]$msg)

    if ($msg) {
        Write-Host $msg -ForegroundColor DarkBlue -BackgroundColor DarkYellow
    }
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
    [void]($ScriptName -match $rgx) 
    $ScriptVersion = $Matches.version
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
    if ($xmlcfg.configuration.appSettings.SelectSingleNode('Outlining')) {
        $psISE.Options.ShowOutlining = [bool]([byte]($xmlcfg.configuration.appSettings.Outlining.value))
    }
    Write-Message " $(Set-ActiveTheme $theme) "
}

Function Get-IseTheme {
    [CmdletBinding(DefaultParameterSetName='List')]
    Param (
        [Parameter(ParameterSetName='Load',Position=0)][string]$Load,
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

    $LoadThemeName = $Load.Substring(0,1).ToUpper() + $Load.Substring(1)
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
                    Write-Message ' Could not find requested theme. Defaulting to most recently modified '
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
                        try {
                            Rename-Item @RenameItemParam -ErrorAction Stop
                        }
                        catch [System.IO.IOException] {
                            $MoveItemParam = $RenameItemParam + @{'Destination' = $RenameItemParam.NewName}
                            $MoveItemParam.Remove('NewName') 
                            Move-Item @MoveItemParam -ErrorAction Stop
                        }
                        catch {
                            Write-Message " $($_.Exception.Message) "
                        }
                        if ($xmlcfg.configuration.appSettings.ActiveThemeName.value -eq $CurrentThemeName) {
                            Set-ActiveTheme $NewThemeName
                        }
                        break
                    }
                }
            }
            default { }
        }
    }
    else {
        Write-Message ' Could not find any StorableColorTheme files '
    }
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

    $script:xmlcfg = [xml](Get-Content $configfile)
    $ext = [System.IO.Path]::GetExtension($psISE.CurrentFile.FullPath)
    switch ($PSCmdlet.ParameterSetName) {
        'CurrentFile' {
            Write-Message " $($psISE.CurrentFile.FullPath) "
        }
        'Revision' {
            $OriginalFile = $psISE.CurrentFile
            $rev = (Get-ScriptVersion 'Number') + 5
            $NewFileName = "$(Get-ScriptVersion 'BaseName')$rev" + $ext
            $NewFilePath = Join-Path ($xmlcfg.configuration.appSettings.RepositoryPath.value) $NewFileName
            if ([System.IO.File]::Exists($NewFilePath)) {
               Write-Message " $NewFilePath already exists! " 
            }
            else {
                $NewFileContent = $psISE.CurrentFile.Editor.Text
                $ErrorActionPreference = 'Stop'
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
                        Write-Host " $($_.Exception.Message) "
                    }
                }
                $ErrorActionPreference = 'Continue'
            }
        }
        'Fork' {
            $NewFileName = "$(Get-ScriptVersion 'DisplayName')F_0" + $ext
            try {
                $NewFilePath = Join-Path ($xmlcfg.configuration.appSettings.RepositoryPath.value) $NewFileName
            }
            catch {
                Write-Message ' Invalid Repository path '
                Return
            }
            if ([System.IO.File]::Exists($NewFilePath)) {
                Write-Message " $NewFilePath already exists! " 
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
                        Write-Message " $($_.Exception.Message) "
                    }
                    $ErrorActionPreference = 'Continue'
                }
                else {
                    Write-Message ' Select the code you wish to fork '
                }
            }
        }
        'RepositoryPath' {
            if ([System.String]::IsNullOrEmpty($SetPath)) { 
                $RepoPath = [string]($xmlcfg.configuration.appSettings.RepositoryPath.value)
                if ([System.String]::IsNullOrEmpty($RepoPath)) {
                    $message = 'Repository Path not set'
                }
                else {
                    $message = $RepoPath
                    if (-not(Test-Path -Path $RepoPath -PathType Container)) {
                        $message = $message + ' (failed Test-Path)'
                    }
                }
                Write-Message " $message "
            }
            else {
                $xmlcfg.configuration.appSettings.RepositoryPath.value = $SetPath
                $xmlcfg.Save($configfile)
            }
        }
        default { }
    }
}


