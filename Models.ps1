#region ClickOnce Registry
class ClickOnceRegistryKey {
    [string] $Key
}

class Component : ClickOnceRegistryKey {
    [string[]] $Dependencies 
    Component([string]$keyName) {
        $this.Key = $keyName
    }
}

class Implication : ClickOnceRegistryKey {
    [string]$Name
    [string]$Value
    Implication([string]$keyName) {
        $this.Key = $keyName
    }
}

class RegistryMarker {
    [Microsoft.Win32.RegistryKey] $Parent
    [string] $ItemName
    RegistryMarker([Microsoft.Win32.RegistryKey]$key, [string]$itemName) {
        $this.Parent = $key
        $this.ItemName = $itemName
    }
}

class Mark : ClickOnceRegistryKey {
    [string]$AppId
    [string]$Identity
    [Implication[]]$Implications
    Mark([string]$keyName) {
        $this.Key = $keyName
    }
}

class ClickOnceRegistry {
    [string]$ComponentsRegistryPath = "Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0\Components";
    [string]$MarksRegistryPath = "Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0\Marks";
    [Component[]] $Components
    [Mark[]] $Marks
    ClickOnceRegistry() {
        $this.ReadComponents()
        $this.ReadMarks()
    }
    [void] ReadComponents() {
        $this.Components = @()
        $comps = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($this.ComponentsRegistryPath)
        if ($null -eq $comps) { return }
        foreach ($keyName in $comps.GetSubKeyNames()) {
            $componentKey = $comps.OpenSubKey($keyName)
            if ($null -eq $componentKey) { continue }
            $component = [Component]::new($keyName)
            $this.Components += $component
            $component.Dependencies = @()
            foreach ($dependencyName in ($componentKey.GetSubKeyNames() | Where-Object { $_ -ne "Files" })) {
                $component.Dependencies += $dependencyName
            }
        }
    }   

    [void] ReadMarks() {
        $this.Marks = @()
        $mks = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($this.MarksRegistryPath)
        if ($null -eq $mks ) { return }
        foreach ($keyName in $mks.GetSubKeyNames()) {
            $markKey = $mks.OpenSubKey($keyName)
            if ($null -eq $markKey) { continue }
            $mark = [Mark]::new($keyName)
            $this.Marks += $mark
            [byte[]]$appId = $markKey.GetValue("appid")
            if ($null -ne $appId) {
                $mark.AppId = [System.Text.Encoding]::ASCII.GetString($appId)
            }
            [byte[]]$identity = $markKey.GetValue("identity")
            if ($null -ne $identity) {
                $mark.Identity = [System.Text.Encoding]::ASCII.GetString($identity)
            }
            $mark.Implications = @()
            $impls = $markKey.GetValueNames() | Where-Object { $_ -like "implication*" }
            foreach ($implicationName in $impls) {
                $implicationName = [string]$implicationName
                [byte[]]$implication = $markKey.GetValue($implicationName)
                if ($null -ne $implication) {
                    $impl = [Implication]::new($implicationName) 
                    $impl.Name = $implicationName.Substring(12)
                    $impl.Value = [System.Text.Encoding]::ASCII.GetString($implication)
                    $mark.Implications += $impl  
                }
            }            
        }
    }
}
#endregion

#region Uninstall Info Model
class UninstallInfo {
    static [string]$UninstallRegistryPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
    [string]$Key
    [string]$UninstallString
    [string]$ShortcutFolderName
    [string]$ShortcutSuiteName
    [string]$ShortcutFileName
    [string]$SupportShortcutFileName
    static [UninstallInfo] Find([string]$appName) {
        $uninstallKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey([UninstallInfo]::UninstallRegistryPath)
        if ($null -eq $uninstallKey) { return $null }
        
        foreach ($app in $uninstallKey.GetSubKeyNames()) {
            $appKey = $uninstallKey.OpenSubKey($app)
            if ($null -ne $appKey -and ([string]$appKey.GetValue("DisplayName")) -eq $appName) {
                return New-Object UninstallInfo -Property @{
                    Key                     = $app
                    UninstallString         = [string]$appKey.GetValue("UninstallString") 
                    ShortcutFolderName      = [string]$appKey.GetValue("ShortcutFolderName") 
                    ShortcutSuiteName       = [string]$appKey.GetValue("ShortcutSuiteName") 
                    ShortcutFileName        = [string]$appKey.GetValue("ShortcutFileName") 
                    SupportShortcutFileName = [string]$appKey.GetValue("SupportShortcutFileName") 
                }
            }           
        }
        return $null
    }    
    [string] GetPublicKeyToken() {
        $token = ($this.UninstallString.Split(",") | Where-Object { $_.Trim() -like "PublicKeyToken=*" } | Select-Object -First 1).Substring(16)
        if ($token.Length -ne 16) { throw [System.ArgumentException]::new() }
        return $token
    }
}
#endregion

#region UninstallActions
class CloseOpenApplication {
    hidden [UninstallInfo]$_uninstallInfo
    hidden [bool]$_applicationIsClosed = $true
    hidden [System.Diagnostics.Process]$_process


    CloseOpenApplication([UninstallInfo]$uninstallInfo) {
        $this._uninstallInfo = $uninstallInfo
    }

    [void] Prepare([string[]]$componentsToRemove) {
        $appsFolder = Join-Path ([Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)) -ChildPath "Apps\2.0"
        $publicKeyToken = $this._uninstallInfo.GetPublicKeyToken()       
        $this._process = Get-Process | Where-Object { $_.MainModule.FileName -match ($appsFolder -replace "\\", "\\") -and $_.MainModule.FileName -match $publicKeyToken } | Select-Object -First 1
        $this._applicationIsClosed = $null -eq $this._process
    }

    [void] PrintDebugInformation() {
        if ($this._applicationIsClosed) {
            Write-Host "Clickonce application is not currently active, continuing" -ForegroundColor green
            return
        } 
        
        Write-Host "ClickOnce '$($this._process.ProcessName)' application active! " -ForegroundColor red -NoNewline
        Write-Host "Closing..."
    }

    [void] Execute() {
        if ($this._applicationIsClosed) {
            return
        }

        $this._process | Stop-Process -Force 
        Start-Sleep -Milliseconds 500
    }
}

class RemoveFiles {
    hidden [string] $_clickOnceFolder
    hidden [string[]] $_foldersToRemove = @()
    hidden [string[]] $_filesToRemove = @()

    [void] Prepare([string[]] $componentsToRemove) {
        $this._clickOnceFolder = $this.FindClickOnceFolder()
        foreach ($directoryItem in (Get-ChildItem $this._clickOnceFolder -Directory)) {
            if ($directoryItem.Name -in $componentsToRemove) {
                $this._foldersToRemove += $directoryItem.FullName
            }
        }

        foreach ($fileItem in (Get-ChildItem $this._clickOnceFolder -File)) {
            if ($fileItem.Name -in $componentsToRemove) {
                $this._filesToRemove += $fileItem.FullName
            }
        }
    }
    [void] PrintDebugInformation() {
        if ([string]::IsNullOrEmpty($this._clickOnceFolder)) {
            throw [System.InvalidOperationException]::new("Call Prepare() first")
        }

        foreach ($folder in $this._foldersToRemove) {
            Write-Host "Delete Folder: " -ForegroundColor Red -NoNewline
            Write-Host "$folder"
        }

        foreach ($file in $this._filesToRemove) {
            Write-Host "Delete File: " -ForegroundColor Red -NoNewline
            Write-Host "$file"
        }
    }
    [void] Execute() {
        if ([string]::IsNullOrEmpty($this._clickOnceFolder) -or -not (Test-Path $this._clickOnceFolder)) {
            throw [System.InvalidOperationException]::new("Call Prepare() first")
        }
        foreach ($folder in $this._foldersToRemove) {
            try {
                Remove-Item -Path $folder -Recurse -Force
            }
            catch {
                Write-Warning "Failed to delete folder $folder. Error: $_"
            }
        }

        foreach ($file in $this._filesToRemove) {
            try {
                Remove-Item -Path $file -Force
            }
            catch {
                Write-Warning "Failed to delete file $file. Error: $_"
            }
        }
    }
    [string] FindClickOnceFolder() {
        $appsFolder = Join-Path ([Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)) -ChildPath "Apps\2.0"
        if (-not (Test-Path $appsFolder)) { throw [System.ArgumentException]::new("Could not find ClickOnce folder at: $appsFolder") }
        foreach ($subFolderItem in (Get-ChildItem -Path $appsFolder -Directory)) {
            if ($subFolderItem.Name.Length -eq 12) {
                foreach ($nestedSubFolderItem in (Get-ChildItem -path $subFolderItem.FullName -Directory)) {
                    if ($nestedSubFolderItem.Name.Length -eq 12) {
                        return $nestedSubFolderItem.FullName
                    }
                }
            } 
        }
        throw [System.ArgumentException]::new("Could not find ClickOnce folder")
    }
}

class RemoveRegistryKeys {
    [string] $PackageMetadataRegistryPath = "Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0\PackageMetadata"
    [string] $ApplicationRegistryPath = "Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0\StateManager\Applications"
    [string] $FamiliesRegistryPath = "Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0\StateManager\Families"
    [string] $VisibilityRegistryPath = "Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0\Visibility"

    hidden [ClickOnceRegistry]$_clickOnceRegistry
    hidden [UninstallInfo]$_uninstallInfo
    hidden [System.IDisposable[]]$_disposables = @()
    hidden [RegistryMarker[]]$_keysToRemove
    hidden [RegistryMarker[]]$_valuesToRemove

    RemoveRegistryKeys([ClickOnceRegistry]$registry, [UninstallInfo]$uninstallInfo) {
        $this._clickOnceRegistry = $registry
        $this._uninstallInfo = $uninstallInfo
    } 
    [void] Prepare([string[]] $componentsToRemove) {
        $this._keysToRemove = @()
        $this._valuesToRemove = @()

        $componentsKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($this._clickOnceRegistry.ComponentsRegistryPath, $true)
        $this._disposables += $componentsKey
        foreach ($component in $this._clickOnceRegistry.Components) {
            if ($component.Key -in $componentsToRemove) {
                $this._keysToRemove += [RegistryMarker]::new($componentsKey, $component.Key)
            }
        }

        $marksKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($this._clickOnceRegistry.MarksRegistryPath, $true)
        $this._disposables += $marksKey
        foreach ($mark in $this._clickOnceRegistry.Marks) {
            if ($mark.Key -in $componentsToRemove) {
                $this._keysToRemove += [RegistryMarker]::new($marksKey, $mark.Key)
            }
            else {
                $implications = $mark.Implications | Where-Object Name -in $componentsToRemove
                if ($implications.Count -gt 0) {
                    $markKey = $marksKey.OpenSubKey($mark.Key, $true)
                    $this._disposables += $markKey

                    foreach ($implication in $implications) {
                        $this._valuesToRemove += [RegistryMarker]::new($markKey, $implication.Key)
                    }
                }
            }
        }

        $token = $this._uninstallInfo.GetPublicKeyToken()

        $packageMetadata = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($this.PackageMetadataRegistryPath)
        foreach ($keyName in $packageMetadata.GetSubKeyNames()) {
            $keyPath = Join-Path $this.PackageMetadataRegistryPath  $keyName
            $this.DeleteMatchingSubKey($keyPath, $token)
        }

        $this.DeleteMatchingSubKey($this.ApplicationRegistryPath, $token)
        $this.DeleteMatchingSubKey($this.FamiliesRegistryPath, $token)
        $this.DeleteMatchingSubKey($this.VisibilityRegistryPath, $token)
    }
    
    [void] PrintDebugInformation() {
        if ($null -eq $this._keysToRemove) {
            throw [System.InvalidOperationException]::new("Please run Prepare() before printing debug information")
        }

        foreach ($key in $this._keysToRemove) {
            Write-Host "Delete key: " -ForegroundColor Red -NoNewline
            Write-Host "$($key.Parent) in $($key.ItemName)"
        }

        foreach ($value in $this._valuesToRemove) {
            Write-Host "Delete value:" -ForegroundColor red -NoNewline
            Write-Host "$($value.Parent) in $($value.ItemName)"
        }
    }

    [void] Execute() {
        if ($null -eq $this._keysToRemove) {
            throw [System.InvalidOperationException]::new("Please run Prepare() before executing uninstall step")
        }

        foreach ($key in $this._keysToRemove) {
            $key.Parent.DeleteSubKeyTree($key.ItemName)
        }
        foreach ($value in $this._valuesToRemove) {
            $value.Parent.DeleteValue($value.ItemName)
        }
    }

    hidden [void] DeleteMatchingSubKey([string]$registryPath, [string]$token) {
        $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($registryPath, $true)
        $this._disposables += $key
        foreach ($subKeyName in $key.GetSubKeyNames()) {
            if ($subKeyName -like "*$token*") {
                $this._keysToRemove += [RegistryMarker]::new($key, $subKeyName)
            }
        }
    }
}

class RemoveStartMenuEntry {
    hidden [UninstallInfo] $_uninstallInfo
    hidden [string[]]$_foldersToRemove
    hidden [string[]]$_filesToRemove

    RemoveStartMenuEntry([UninstallInfo]$uninstallInfo) {
        $this._uninstallInfo = $uninstallInfo
    }

    [void] Prepare([string[]] $componentsToRemove) {
        $this._filesToRemove = @()
        $this._foldersToRemove = @()

        $programsFolder = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Programs)
        $folder = Join-Path $programsFolder -ChildPath $this._uninstallInfo.ShortcutFolderName
        $suiteFolder = Join-Path $folder -ChildPath $this._uninstallInfo.ShortcutSuiteName
        $shortcut = Join-Path $suiteFolder -ChildPath "$($this._uninstallInfo.ShortcutFileName).appref-ms"
        $supportShortcut = Join-Path $suiteFolder -ChildPath "$($this._uninstallInfo.SupportShortcutFileName).url"
        $desktopFolder = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
        $desktopShortcut = Join-Path $desktopFolder -ChildPath "$($this._uninstallInfo.ShortcutFileName).appref-ms"
        if (Test-Path $shortcut) {
            $this._filesToRemove += $shortcut
        }
        if (Test-Path $supportShortcut) {
            $this._filesToRemove += $supportShortcut
        }
        if (Test-Path $desktopShortcut) {
            $this._filesToRemove += $desktopShortcut
        }

        if ((Test-Path $suiteFolder) -and ($this._filesToRemove | Where-Object { $_ -notin (Get-ChildItem $suiteFolder -file).FullName }).Count -eq 0) {
            $this._foldersToRemove += $suiteFolder

            $folderFolders = Get-ChildItem $folder -Directory
            $folderFiles = Get-ChildItem $folder -File

            if ($folderFolders.Count -eq 1 -and $folderFiles.Count -eq 0) {
                $this._foldersToRemove += $folder
            }
        }
    }
    
    [void] PrintDebugInformation() {
        if ($null -eq $this._foldersToRemove) {
            throw [System.InvalidOperationException]::new("Call Prepare() first")
        }

        foreach ($file in $this._filesToRemove) {
            Write-Host "Delete file: " -NoNewline -ForegroundColor red
            Write-Host $file
        }

        foreach ($folder in $this._foldersToRemove) {
            Write-Host "Delete folder: " -NoNewline -ForegroundColor red
            Write-Host $folder
        }

    }

    [void] Execute() {
        if ($null -eq $this._foldersToRemove) { throw [System.InvalidOperationException]::new("Call Prepare() first") }

        foreach ($file in $this._filesToRemove) {
            try {
                Remove-Item -Path $file -Force
            }
            catch {
                Write-Warning "Failed to delete file $file. Error: $_"
            }
        }

        foreach ($folder in $this._foldersToRemove) {
            try {
                Remove-Item -Path $folder -Force -Recurse
            }
            catch {
                Write-Warning "Failed to delete folder $folder. Error: $_"
            }
        }
    }
}

class RemoveUninstallEntry {
    hidden [UninstallInfo] $_uninstallInfo
    hidden [Microsoft.Win32.RegistryKey]$_uninstall

    RemoveUninstallEntry([UninstallInfo]$uninstallInfo) {
        $this._uninstallInfo = $uninstallInfo
    }
    [void] Prepare([string[]] $componentsToRemove) {
        $this._uninstall = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey([UninstallInfo]::UninstallRegistryPath, $true)
    }
    
    [void] PrintDebugInformation() {
        if ($null -eq $this._uninstall) {
            throw [System.InvalidOperationException]::new("Please call Prepare() first")
        }

        Write-Host "Remove uninstall information: " -NoNewline -ForegroundColor red
        Write-Host $this._uninstall.OpenSubKey($this._uninstallInfo.Key).Name
    }

    [void] Execute() {
        if ($null -eq $this._uninstall) {
            throw [System.InvalidOperationException]::new("Please call Prepare() first")
        }

        $this._uninstall.DeleteSubKey($this._uninstallInfo.Key)
    }
}
#endregion

#region UNINSTALLER
class Uninstaller {
    hidden [ClickOnceRegistry]$_registry
    Uninstaller([ClickOnceRegistry]$registry) {
        $this._registry = $registry
    }

    Uninstaller() {
        $this._registry = [ClickOnceRegistry]::new()
    }

    [void] Uninstall([UninstallInfo]$uninstallInfo) {
        $toRemove = $this.FindComponentsToRemove($uninstallInfo.GetPublicKeyToken())

        Write-Host "Components to remove:" -ForegroundColor red
        $toRemove | ForEach-Object { Write-Host "`t- $_" }


        $steps = @(
            [CloseOpenApplication]::new($uninstallInfo)
            [RemoveFiles]::new(),
            [RemoveStartMenuEntry]::new($uninstallInfo),
            [RemoveRegistryKeys]::new($this._registry, $uninstallInfo),
            [RemoveUninstallEntry]::new($uninstallInfo)
        )

        $steps | ForEach-Object {
            $_.Prepare($toRemove)
        }

        $steps | ForEach-Object {
            $_.PrintDebugInformation()
        }
        
        $steps | ForEach-Object {
            $_.Execute()
        }
    }

    [string[]] FindComponentsToRemove([string]$token) {
        $components = $this._registry.Components | Where-Object Key -like "*$token*"

        $toRemove = @()
        foreach ($component in $components) {
            $toRemove += $component.Key

            foreach ($dependency in $component.Dependencies) {
                if ($dependency -in $toRemove) { continue }
                if ($dependency -notin $this._registry.Components) { continue }

                $mark = $this._registry.Marks | Where-Object Key -eq $dependency
                if ($null -ne $mark -and ($mark.Implications | Where-Object Name -in $components.Key).Count -eq 0) {
                    continue
                }

                $toRemove += $dependency
            }
        }
        return $toRemove
    }
}

#endregion
# SIG # Begin signature block
# MIIfMgYJKoZIhvcNAQcCoIIfIzCCHx8CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDmbcPaFJJO2cD0
# 4BFw8d/6hDR3mLfNG3k/CYiJKxvZu6CCGU8wggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwggYRMIIE+aADAgECAhNLAANK+NuR4l0sB1+2AAEAA0r4MA0G
# CSqGSIb3DQEBCwUAMEYxEjAQBgoJkiaJk/IsZAEZFgJiZTEYMBYGCgmSJomT8ixk
# ARkWCGF0c2dyb2VwMRYwFAYDVQQDEw1BdHNDZXJ0U3J2MDAyMB4XDTI1MDcxNTEz
# MDU0M1oXDTI4MDcxNTEzMDU0M1owXDELMAkGA1UEBhMCQkUxETAPBgNVBAgTCEZs
# YW5kZXJzMRIwEAYDVQQHEwlNZXJlbGJla2UxEjAQBgNVBAoTCUFUUyBHcm9lcDES
# MBAGA1UEAxMJQVRTIEdyb2VwMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
# AQEAn5Lkqf9bBQWeILRwzcTsHfCwqIJ4ngFxjak5QBylaq6qclst3QZ4OI8qimfu
# fOITWY8oZyA/zsJ8goDKGVnu9bC/h7Df3+elEOBljxDhiZ4U1lzNKZdDBcuPpScK
# jwQZOk5O2teeglklC4P+DF51nNL1flY7ZiEqpgxd9aHUm28QzBfw5yGLn9M/j84s
# WtyJh/B7hf7zce7c6LBYC6h6q/9kwD8d/8vxPkSr+pi4pDC36kMTpAiqZYB7nKxt
# SUnLCnGiPavNxLV7B06CYq+jbqq1RuSZiJBaqkAZBEdftprzhIzhgtFJde6F4WBZ
# kBeIXxjTrmuRslstR6s+vVUoDQIDAQABo4IC4DCCAtwwOwYJKwYBBAGCNxUHBC4w
# LAYkKwYBBAGCNxUIhaOTfqK3WYbZmwuC84pzhsphHIe3zkGF2upwAgFnAgEIMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAA
# MBsGCSsGAQQBgjcVCgQOMAwwCgYIKwYBBQUHAwMwTwYJKwYBBAGCNxkCBEIwQKA+
# BgorBgEEAYI3GQIBoDAELlMtMS01LTIxLTIwMjU0MjkyNjUtNDEzMDI3MzIyLTE0
# MTcwMDEzMzMtNDU2MzMwHQYDVR0OBBYEFIHHpkR3LhkocoJJtTy62xCsevAFMB8G
# A1UdIwQYMBaAFLLJJhXrXffN6To4qag1kvKeba41MIH5BgNVHR8EgfEwge4wgeug
# geiggeWGgbRsZGFwOi8vL0NOPUF0c0NlcnRTcnYwMDIsQ049dkFUUzAwMSxDTj1D
# RFAsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29u
# ZmlndXJhdGlvbixEQz1hdHNncm9lcCxEQz1iZT9jZXJ0aWZpY2F0ZVJldm9jYXRp
# b25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9uUG9pbnSGLGh0
# dHA6Ly9jZHAuYXRzZ3JvZXAuYmUvY3JsL0F0c0NlcnRTcnYwMDIuY3JsMIG/Bggr
# BgEFBQcBAQSBsjCBrzCBrAYIKwYBBQUHMAKGgZ9sZGFwOi8vL0NOPUF0c0NlcnRT
# cnYwMDIsQ049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZp
# Y2VzLENOPUNvbmZpZ3VyYXRpb24sREM9YXRzZ3JvZXAsREM9YmU/Y0FDZXJ0aWZp
# Y2F0ZT9iYXNlP29iamVjdENsYXNzPWNlcnRpZmljYXRpb25BdXRob3JpdHkwDQYJ
# KoZIhvcNAQELBQADggEBAGO86qpdChLa6MfcDLpKHKthUNqU8uqpaA0pGjVQBYjL
# 99aMmjMLdnYYx/pzvIfIEhtsjh1M+fR70V1c+HE5jBiiAWckz87hEoi0GdGAL/44
# 22/hsiXj4SPgYQF6fwT1t3XOM/REwnhYsaqf2l0jmplA34r993/fFOPn5HZVjWyc
# mrsHMqKD+kQf6hLO3jSL0NGBA9iXArw1NwERdALDieJSxrzgqiM9e6LPVGwj60+z
# DKEY4U9aD5nWi0LLDMe0ot3BOJswGw2yyezhDCBsXCDAjbY9S0RE939itPXu4fVP
# v9RRCkMiujDaXBJ/q2gz3mrXBXTI2oVEIlEfAYDBIWUwgga0MIIEnKADAgECAhAN
# x6xXBf8hmS5AQyIMOkmGMA0GCSqGSIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUw
# EwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20x
# ITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yNTA1MDcwMDAw
# MDBaFw0zODAxMTQyMzU5NTlaMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQC0eDHTCphBcr48RsAcrHXbo0ZodLRRF51NrY0NlLWZloMs
# VO1DahGPNRcybEKq+RuwOnPhof6pvF4uGjwjqNjfEvUi6wuim5bap+0lgloM2zX4
# kftn5B1IpYzTqpyFQ/4Bt0mAxAHeHYNnQxqXmRinvuNgxVBdJkf77S2uPoCj7GH8
# BLuxBG5AvftBdsOECS1UkxBvMgEdgkFiDNYiOTx4OtiFcMSkqTtF2hfQz3zQSku2
# Ws3IfDReb6e3mmdglTcaarps0wjUjsZvkgFkriK9tUKJm/s80FiocSk1VYLZlDwF
# t+cVFBURJg6zMUjZa/zbCclF83bRVFLeGkuAhHiGPMvSGmhgaTzVyhYn4p0+8y9o
# HRaQT/aofEnS5xLrfxnGpTXiUOeSLsJygoLPp66bkDX1ZlAeSpQl92QOMeRxykvq
# 6gbylsXQskBBBnGy3tW/AMOMCZIVNSaz7BX8VtYGqLt9MmeOreGPRdtBx3yGOP+r
# x3rKWDEJlIqLXvJWnY0v5ydPpOjL6s36czwzsucuoKs7Yk/ehb//Wx+5kMqIMRvU
# BDx6z1ev+7psNOdgJMoiwOrUG2ZdSoQbU2rMkpLiQ6bGRinZbI4OLu9BMIFm1UUl
# 9VnePs6BaaeEWvjJSjNm2qA+sdFUeEY0qVjPKOWug/G6X5uAiynM7Bu2ayBjUwID
# AQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU729TSunk
# Bnx6yuKQVvYv1Ensy04wHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08w
# DgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEB
# BGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsG
# AQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVz
# dGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBABfO+xaAHP4H
# PRF2cTC9vgvItTSmf83Qh8WIGjB/T8ObXAZz8OjuhUxjaaFdleMM0lBryPTQM2qE
# JPe36zwbSI/mS83afsl3YTj+IQhQE7jU/kXjjytJgnn0hvrV6hqWGd3rLAUt6vJy
# 9lMDPjTLxLgXf9r5nWMQwr8Myb9rEVKChHyfpzee5kH0F8HABBgr0UdqirZ7bowe
# 9Vj2AIMD8liyrukZ2iA/wdG2th9y1IsA0QF8dTXqvcnTmpfeQh35k5zOCPmSNq1U
# H410ANVko43+Cdmu4y81hjajV/gxdEkMx1NKU4uHQcKfZxAvBAKqMVuqte69M9J6
# A47OvgRaPs+2ykgcGV00TYr2Lr3ty9qIijanrUR3anzEwlvzZiiyfTPjLbnFRsjs
# Yg39OlV8cipDoq7+qNNjqFzeGxcytL5TTLL4ZaoBdqbhOhZ3ZRDUphPvSRmMThi0
# vw9vODRzW6AxnJll38F0cuJG7uEBYTptMSbhdhGQDpOXgpIUsWTjd6xpR6oaQf/D
# Jbg3s6KCLPAlZ66RzIg9sC+NJpud/v4+7RWsWCiKi9EOLLHfMR2ZyJ/+xhCx9yHb
# xtl5TPau1j/1MIDpMPx0LckTetiSuEtQvLsNz3Qbp7wGWqbIiOWCnb5WqxL3/BAP
# vIXKUjPSxyZsq8WhbaM2tszWkPZPubdcMIIG7TCCBNWgAwIBAgIQCoDvGEuN8QWC
# 0cR2p5V0aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMO
# RGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGlt
# ZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMB4XDTI1MDYwNDAwMDAw
# MFoXDTM2MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBUaW1l
# c3RhbXAgUmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3zBlCMGMyqJnfFNZx+wvA
# 69HFTBdwbHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8TchTySA2R4QKpVD7dvNZh6w
# W2R6kSu9RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWjFDYOzDi8SOhPUWlLnh00
# Cll8pjrUcCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2UoyrN0ijtUDVHRXdmncOOM
# A3CoB/iUSROUINDT98oksouTMYFOnHoRh6+86Ltc5zjPKHW5KqCvpSduSwhwUmot
# uQhcg9tw2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KSuNLoZLc1Hf2JNMVL4Q1O
# pbybpMe46YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7wJNdoRORVbPR1VVnDuSeH
# VZlc4seAO+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vWdoUoHLWnqWU3dCCyFG1r
# oSrgHjSHlq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOgrY7rlRyTlaCCfw7aSURO
# wnu7zER6EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K096V1hE0yZIXe+giAwW0
# 0aHzrDchIc2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCfgPf8+3mnAgMBAAGjggGV
# MIIBkTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zyMe39/dfzkXFjGVBDz2GM
# 6DAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMC
# B4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXQYIKwYBBQUHMAKG
# UWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSg
# UqBQhk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRU
# aW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAI
# BgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQBlKq3xHCcE
# ua5gQezRCESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZD9gBq9fNaNmFj6Eh8/Ym
# RDfxT7C0k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/ML9lFfim8/9yJmZSe2F8
# AQ/UdKFOtj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu+WUqW4daIqToXFE/JQ/E
# ABgfZXLWU0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4obEMnxYOX8VBRKe1uNnzQ
# VTeLni2nHkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2hECZpqyU1d0IbX6Wq8/gV
# utDojBIFeRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasnM9AWcIQfVjnzrvwiCZ85
# EE8LUkqRhoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol/DJgddJ35XTxfUlQ+8Hg
# gt8l2Yv7roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgYxQbV1S3CrWqZzBt1R9xJ
# gKf47CdxVRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3ocCVccAvlKV9jEnstrniLv
# UxxVZE/rptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcBZU8atufk+EMF/cWuiC7P
# OGT75qaL6vdCvHlshtjdNXOCIUjsarfNZzGCBTkwggU1AgEBMF0wRjESMBAGCgmS
# JomT8ixkARkWAmJlMRgwFgYKCZImiZPyLGQBGRYIYXRzZ3JvZXAxFjAUBgNVBAMT
# DUF0c0NlcnRTcnYwMDICE0sAA0r425HiXSwHX7YAAQADSvgwDQYJYIZIAWUDBAIB
# BQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYK
# KwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG
# 9w0BCQQxIgQgNnc3ttMspFMCumoctunUJhyOeMrXAHMmoIjy1ZkuYekwDQYJKoZI
# hvcNAQEBBQAEggEAPK2FeLkU4ecMfAO60CCuC7cxvFtZKEX8fExxhzEIIfmjtF+v
# 7oTkGNjOHFRw/EfSsJ4l1B1ZfIRHrz1LgZlbUEgqqueT3EHnuym/WzVVcFmoCKQ4
# xEuSNbfWIaNYy3e8YLFWk/gBFsM5+qGscb2LeDW7kW3oIxwPXJmr1ZziJGMRyN9m
# /izw7fReZ14bv8Yb99ajp3wj0zrZezy8bxFe5U9PSlj7I1Qa0TfhKrA/E5ylJ7Dn
# 9Smg/73luUGBfFBmwYGGf50L8jEq0QheesAwJdpSaFxfZnd0sTWCvPPbsUodcVYv
# GD8zh/S25QnWORzk24dXLY1DTQ8qxRhFnePBIaGCAyYwggMiBgkqhkiG9w0BCQYx
# ggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcg
# UlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZI
# AWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJ
# BTEPFw0yNjAzMTgxMDA2MDRaMC8GCSqGSIb3DQEJBDEiBCBKSlD71kOskuWUfh2D
# yba9rJ7dSen8KguFEph9CKUW+jANBgkqhkiG9w0BAQEFAASCAgCd+emvQVPDhTwd
# qqQCOHnAHl+XM04V8NmMO6aPgYHalmOaMjxw0xlxbjWvTjD2YUvwH4Qp2TydvopH
# v2ArXSl+DEZnQ2ZDkDmrofv8/5mHWgKDOHmISm1RaVOrN371saSlNCOML3D6cpqd
# pjPUWjiFlym0eD9loW/pqpvC1W46lL6vov1kgN/TsZ9eW8TaQ5AL8e+xOtboc3ZT
# hygv59WyLyu8ymwEKbsZi0DLbDoM0zRXWeLs4smj+kSE8bPzXrE1zwxslbtGt5pG
# 7XYWZUhQxyr096dcvqgBVhVJVFY9jBRkBSfzSS+X4cAXSlkqacQwvhm8z6t2Ei/L
# HbrFbTKUdjZ5YL4PgoAUUVtUr2gPIAWyatp4FYHf+Xt/304m2yYv+jHc9STqM3hK
# BakcRzQvHapFDZXpgx4X8QGObcsxEgRRCZQwuKJkvSdFvpfhRb4rTKnsYObhrq0K
# vB/7X1kKZ9CRZqyeN8IPIUk9KUl7KJao2TWdTCsDA6WjOGqGOHXbzjLHoka6mkhP
# 2nPdZ1GFYO88+QRa59RuNx55U9zy6mAYVrk5NG4O2Yl+47TNNP7za/BBVaLkkiyn
# lF0l5CsmAeNmV8WG8Fup2GYMuyFo58IqOiubGfMPdzAA4EN0bMfIypoXtckGaJMF
# z/ghXNymdm5PuVqnPRayZNPz0iEolA==
# SIG # End signature block
