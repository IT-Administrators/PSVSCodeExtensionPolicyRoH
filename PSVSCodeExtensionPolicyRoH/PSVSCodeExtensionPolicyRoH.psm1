function Set-VSCodeExtensionPolicy {
<#
.SYNOPSIS
    Manages Visual Studio Code extension allow‑lists by adding, removing, or denying entries

.DESCRIPTION
    Manage VSCode extensions on windows via Registry for system installations of VSCode or 
    a config file for user installations.
    In systemcontext the extensions have to be explicitely enabled, all other extensions are 
    denied when the "AllowedExtensions" key is set.

    For example (only applies to system installations of VSCode):
    
    Allow extensions of a publisher and deny all other. 
    {"IT-Administrators": true}

    Allow specific extension and deny all other.
    {"ms-python.python":true}

.PARAMETER AddAllowed 
    Allowed extensions or publishers. Requires full qualified extension name.

.PARAMETER DenyAllowed
    Speicifically deny extensions or publishers.

.PARAMETER RemoveAllowed
    Remove exntensions or publisher from list.

.PARAMETER UserContext
    Configure extensions in user context.

.PARAMETER SystemContext
    Configure extensions for all users on the current device. Requires admin privileges. Fallback to user context
    if VSCode is installed in user profile.

.PARAMETER RemoveUnapprovedExtensions
    Remove all extensions that are not in the approved list.

.PARAMETER EnableAutoUpdate
    Enable auto update for extensions.

.PARAMETER AutoCheckUpdates
    Enable automatic extension update check.

.EXAMPLE
    Set allowed extensions for the current user.

    Set-VSCodeExtensionPolicy -AddAllowed "microsoft","IT-Administrators","ms-vscode" -UserContext -Verbose

.EXAMPLE
    Allow specific extension for all users on the current system. All other extensions are denied.

    Set-VSCodeExtensionPolicy -AddAllowed "ms-python.python" -SystemContext -Verbose

.NOTES
    Compatible with Windows PowerShell 5.1.

.LINK
    https://github.com/IT-Administrators/PSVSCodeExtensionPolicyRoH
#>

    [CmdletBinding(DefaultParameterSetName = "User")]
    param(
        [Parameter(ParameterSetName = "User", 
        Mandatory = $true,
        HelpMessage = "Context to be configured. UserContext applies to the current user.")]
        [switch] $UserContext,

        [Parameter(ParameterSetName = "System", 
        Mandatory = $true,
        HelpMessage = "Context to be configured. SystemContext applies to all users of the current system.")]
        [switch] $SystemContext,
        
        [Parameter(
        HelpMessage = "Allowed extensions (requires full qualified extension name) or publishers.")]
        [string[]] $AddAllowed, 
        
        [Parameter(
        HelpMessage = "Denied extensions (requires full qualified extension name) or publishers.")]
        [string[]] $DenyAllowed,
        
        [Parameter(
        HelpMessage = "Remove extensions (requires full qualified extension name) or publishers.")]
        [string[]] $RemoveAllowed,

        [Parameter(
        HelpMessage = "Remove all extensions that are currently isntalled and not in allow list.")]
        [switch] $RemoveUnapprovedExtensions,
        
        [Parameter(
        ParameterSetName = "User",
        HelpMessage = "Enable auto update for extensions.")]
        [bool] $EnableAutoUpdate = $false,
        
        [Parameter(
        ParameterSetName = "User",
        HelpMessage = "Enable automatic update check.")]
        [bool] $AutoCheckUpdates = $false
    )

    BEGIN {

        function ConvertTo-Hashtable {
            param([Parameter(ValueFromPipeline)] $InputObject)

            if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
                $hash = @{}
                foreach ($prop in $InputObject.PSObject.Properties) {
                    $hash[$prop.Name] = ConvertTo-Hashtable $prop.Value
                }
                return $hash
            }

            if ($InputObject -is [System.Collections.IDictionary]) {
                $hash = @{}
                foreach ($key in $InputObject.Keys) {
                    $hash[$key] = ConvertTo-Hashtable $InputObject[$key]
                }
                return $hash
            }

            if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
                return $InputObject | ForEach-Object { ConvertTo-Hashtable $_ }
            }

            return $InputObject
        }

        function Optimize-List {
            param([string[]] $Items)

            $Clean = @()
            foreach ($i in $Items) {
                if ([string]::IsNullOrWhiteSpace($i)) { continue }
                $Clean += $i.Trim().TrimEnd('*').TrimEnd('.')
            }
            return $Clean
        }
    }

    PROCESS {

        # Normalize input
        $AddAllowed    = Optimize-List $AddAllowed
        $DenyAllowed   = Optimize-List $DenyAllowed
        $RemoveAllowed = Optimize-List $RemoveAllowed

        #
        # Load existing allow-map depending on context
        #
        if ($UserContext) {

            $UserSettingsPath = "$env:APPDATA\Code\User\settings.json"

            if (-not (Test-Path $UserSettingsPath)) {
                New-Item -ItemType File -Path $UserSettingsPath -Force | Out-Null
            }

            try {
                $Raw = Get-Content -Path $UserSettingsPath -Raw
                $CleanRaw = $Raw -replace '(?m)^\s*//.*$' -replace '/\*.*?\*/', ''

                $ExistingSettings = if ($CleanRaw.Trim()) {
                    $CleanRaw | ConvertFrom-Json
                } else { @{} }
            }
            catch {
                Write-Warning "settings.json contains invalid JSON. Rebuilding required keys."
                $ExistingSettings = @{}
            }

            $ExistingSettings = ConvertTo-Hashtable $ExistingSettings
            if (-not ($ExistingSettings -is [hashtable])) { $ExistingSettings = @{} }

            if ($ExistingSettings.ContainsKey("extensions.allowed")) {
                $AllowMap = ConvertTo-Hashtable $ExistingSettings["extensions.allowed"]
            }
            else {
                $AllowMap = @{}
            }
        }

        if ($SystemContext) {

            $PolicyPath = "Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\VSCode"
            if (-not (Test-Path $PolicyPath)) {
                New-Item -Path $PolicyPath -Force | Out-Null
            }

            $existingJson = (Get-ItemProperty -Path $PolicyPath -ErrorAction SilentlyContinue)."AllowedExtensions"

            if ($existingJson) {
                try {
                    $obj = $existingJson | ConvertFrom-Json
                    $AllowMap = @{}
                    foreach ($prop in $obj.PSObject.Properties) {
                        $AllowMap[$prop.Name] = $prop.Value
                    }
                }
                catch {
                    Write-Warning "Invalid JSON in system policy. Rebuilding allow-list."
                    $AllowMap = @{}
                }
            }
            else {
                $AllowMap = @{}
            }
        }

        #
        # Apply allow additions (set to $true)
        #
        foreach ($item in $AddAllowed) {
            $AllowMap[$item] = $true
        }

        #
        # Apply deny (set to $false always)
        #
        foreach ($item in $DenyAllowed) {
            $AllowMap[$item] = $false
        }

        #
        # Apply removals (remove key entirely)
        #
        foreach ($item in $RemoveAllowed) {
            if ($AllowMap.ContainsKey($item)) {
                $AllowMap.Remove($item) | Out-Null
            }
        }

        #
        # Remove unapproved extensions (optional)
        #
        if ($RemoveUnapprovedExtensions) {
            $PossiblePaths = @(
                "$env:LOCALAPPDATA\Programs\Microsoft VS Code\bin\code.cmd",
                "$env:ProgramFiles\Microsoft VS Code\bin\code.cmd",
                "$env:ProgramFiles(x86)\Microsoft VS Code\bin\code.cmd"
            )

            $CodeCLI = $PossiblePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

            if ($CodeCLI) {
                $Installed = & $CodeCLI --list-extensions

                foreach ($ext in $Installed) {
                    $IsAllowed = $false

                    if ($AllowMap.ContainsKey($ext) -and $AllowMap[$ext] -eq $true) {
                        $IsAllowed = $true
                    }
                    else {
                        foreach ($key in $AllowMap.Keys) {
                            if ($AllowMap[$key] -eq $true -and
                                $ext.StartsWith("$key.", [System.StringComparison]::OrdinalIgnoreCase)) {
                                $IsAllowed = $true
                                break
                            }
                        }
                    }

                    if (-not $IsAllowed) {
                        Write-Output "Removing unapproved extension: $ext"
                        & $CodeCLI --uninstall-extension $ext --force | Out-Null
                    }
                }
            }
        }

        #
        # Write back USER CONTEXT
        #
        if ($UserContext) {
            $ExistingSettings["extensions.allowed"] = $AllowMap
            $ExistingSettings["extensions.autoCheckUpdates"] = $AutoCheckUpdates
            $ExistingSettings["extensions.autoUpdate"] = $EnableAutoUpdate
            $ExistingSettings["extensions.ignoreRecommendations"] = $true

            $SettingsJson = $ExistingSettings | ConvertTo-Json -Depth 10
            $SettingsJson | Out-File -FilePath $UserSettingsPath -Encoding UTF8 -Force

            Write-Verbose "User settings.json updated."
        }

        #
        # Write back SYSTEM CONTEXT
        #
        if ($SystemContext) {
            $PolicyPath = "Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\VSCode"
            $PolicyValue = $AllowMap | ConvertTo-Json -Depth 5 -Compress

            Set-ItemProperty -Path $PolicyPath -Name "AllowedExtensions" -Value $PolicyValue -Type String
            Write-Verbose "Machine-wide VS Code policy updated."
        }
    }

    END {
        Write-Output "VS Code extension policy processing completed."
    }
}
