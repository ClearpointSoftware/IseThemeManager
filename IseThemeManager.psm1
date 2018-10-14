
Function Load-IseTheme {
    Param([string]$theme)
    $xmldoc = [xml](Get-Content $theme)
    $keyarr = @($xmldoc.StorableColorTheme.Keys | Select -ExpandProperty string)
    $color = {
        Param ($indx)
        $a = $xmldoc.StorableColorTheme.Values.Color[$indx].A
        $r = $xmldoc.StorableColorTheme.Values.Color[$indx].R
        $g = $xmldoc.StorableColorTheme.Values.Color[$indx].G
        $b = $xmldoc.StorableColorTheme.Values.Color[$indx].B
        $hexcolor = $null
        foreach ($i in @($a,$r,$g,$b)) {
            $hex = [Convert]::ToString($i, 16).ToUpper()
            if ($hex.Length -eq 1) {
                $hex = "0$hex"
            }
            $hexcolor = $hexcolor + $hex
        }
        Return "#$hexcolor"
    }
    $indx = 0
    foreach ($xmlkey in $keyarr) {
        switch -regex ($xmlkey) {
            '(^TokenColors|^ConsoleTokenColors|^XmlTokenColors)' {
                $node = $xmlkey -split '\\'
                if ([regex]::IsMatch($node[1],'\D')) {
                    $psISE.Options.($node[0]).item($node[1]) = (& $color $indx)
                }
            }
            default {
                $psISE.Options.($keyarr[$indx]) = (& $color $indx)
            }
        } 
        $indx++
    }
    $psISE.Options.FontName = ($xmldoc.ChildNodes.FontFamily.GetValue(1))
    $psISE.Options.FontSize = ($xmldoc.ChildNodes.FontSize.GetValue(1))
    Write-Host "IseTheme: $($xmldoc.ChildNodes.Name.GetValue(1))"
    $setcontentparams = @{
        Path = Join-Path $PSScriptRoot 'CurrentThemeName.txt';
        Value = $xmldoc.ChildNodes.Name.GetValue(1);
        Force = $true;
    }
    Set-Content @setcontentparams
}

Function Get-IseTheme {
    [CmdletBinding(DefaultParameterSetName='Name')]
    Param (
        [Parameter(ParameterSetName='Name',Position=0)][string]$Name,
        [Parameter(ParameterSetName='List')][switch]$List,
        [Parameter(ParameterSetName='Rename',Position=1)][switch]$Rename,
        [Parameter(ParameterSetName='Rename',Position=2,Mandatory=$true)][string]$CurrentName,
        [Parameter(ParameterSetName='Rename',Position=3,Mandatory=$false)][string]$NewName = $CurrentName
    )

    $filetype = '.StorableColorTheme.ps1xml'
    $themes = @(gci -Path $PSScriptRoot | ? { $_.Name -match ".*\$filetype" })
    if ($themes.Count -gt 0) {
        if ($List) {
            $CurrentThemeName = Get-Content -Path (Join-Path $PSScriptRoot 'CurrentThemeName.txt')
            $themes | Sort-Object | % {
                if (($_.Name -replace $filetype,'') -eq $CurrentThemeName) {
                    Write-Host "$CurrentThemeName *"
                }
                else {
                    Write-Host $($_.Name -replace $filetype,'')
                }
            }
        }
        else {
            if ($PSCmdlet.ParameterSetName -eq 'Rename') {
                    $Name = $CurrentName
            }
            foreach ($t in ($themes | Sort-Object -Property LastWriteTime -Descending)) {
                if ($t.Name -eq ($Name + $filetype)) {
                    $themefilepath = $t.FullName
                    break
                }
            }
            if ($Rename -and $themefilepath) {
                $xml = [xml](Get-Content $themefilepath)
                $xml.StorableColorTheme.Name = $NewName.Trim()
                $xml.Save($themefilepath)
                $renameitemparams = @{
                    Path = $themefilepath;
                    NewName = Join-Path (Split-Path $themefilepath -Parent) ($NewName + $filetype);
                    Force = $true;
                }
                Rename-Item @renameitemparams
            }
            if ($PSCmdlet.ParameterSetName -eq 'Name') {
                if (-not $themefilepath) {
                    $themefilepath = $themes[0].FullName
                    if ($Name) {
                        $themenotfound = 'IseThemeManager: ' +
                            'Could not find requested theme. Defaulting to most recently modified.'
                        Write-Host $themenotfound -ForegroundColor DarkBlue -BackgroundColor Gray
                    }
                }
                Load-IseTheme $themefilepath
            }
        }
    }
    else {
        $filenotfound = 'IseThemeManager: No StorableColorTheme files available in Module folder.'
        Write-Host $filenotfound -ForegroundColor DarkBlue -BackgroundColor Gray
    }
}
