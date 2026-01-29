function Set-VSCodeExtensionPolicy {
<#
.SYNOPSIS
    Enforces a strict allow‑list of Visual Studio Code extensions.

.DESCRIPTION
    Applies VS Code extension allow‑list policies either per-user or system-wide.
    Safely merges settings.json, strips JSONC comments, and preserves user settings.

    If VSCode is installed per user. The <SystemContext> parameter falls back to user context.

.PARAMETER AllowedExtensions
    Allowed extensions. Requires full qualified extension name.

.PARAMETER AllowedPublishers
    Allowed publishers.

.PARAMETER UserContext
    Configure extensions in user context.

.PARAMETER SystemContext
    Configure extensions for all users on the current device. Requires admin privileges. Fallback to user context
    if VSCode is installed in user profile.

.PARAMETER ForceSystemContext
    Force systemcontext config even if VSCode is installed in user context. Requires admin privileges.

.PARAMETER RemoveUnapprovedExtensions
    Remove all extensions that are not in the approved list.

.PARAMETER EnableAutoUpdate
    Enable auto update for extensions.

.PARAMETER AutoCheckUpdates
    Enable automatic extension update check.

.EXAMPLE
    Set allowed extensions for the current user.

    Set-VSCodeExtensionPolicy -AllowedPublishers "microsoft","IT-Administrators","ms-vscode" -UserContext -Verbose
.NOTES
    Compatible with Windows PowerShell 5.1.

.LINK
    https://github.com/IT-Administrators/PSVSCodeExtensionPolicyRoH
#>

    [CmdletBinding(DefaultParameterSetName = "User")]
    param(
        [Parameter(
        HelpMessage = "Allowed extensions (full qualified name) for example: ms-vscode.powershell")]
        [string[]] $AllowedExtensions,

        [Parameter(
        HelpMessage = "Allowed publishers for example: 'ms-vscode','microsoft'")]
        [string[]] $AllowedPublishers,

        [Parameter(
        ParameterSetName = "User",
        HelpMessage = "Configure extensions for current user.")]
        [switch] $UserContext,

        [Parameter(
        ParameterSetName = "System",
        HelpMessage = "Configure extensions for all users on current device. Requires Admin priviliges.")]
        [switch] $SystemContext,

        [Parameter(
        HelpMessage = "Enforce system configuration even if VSCode is installed in user profile. Requires Admin priviliges.")]
        [switch] $ForceSystemContext,

        [Parameter(
        HelpMessage = "Remove all unapproved extensions.")]
        [switch] $RemoveUnapprovedExtensions,

        [Parameter(
        HelpMessage = "Enable extensions auto updates. Default = false.")]
        [bool] $EnableAutoUpdate = $false,

        [Parameter(
        HelpMessage = "Enable extension update check. Default = false.")]
        [bool] $AutoCheckUpdates = $false
    )

    BEGIN {

        function ConvertTo-Hashtable {
            param([Parameter(ValueFromPipeline)] $InputObject)

            if ($InputObject -is [System.Collections.IDictionary]) {
                $hash = @{}
                foreach ($key in $InputObject.Keys) {
                    $hash[$key] = ConvertTo-Hashtable $InputObject[$key]
                }
                return $hash
            }
            elseif ($InputObject -is [System.Collections.IEnumerable] -and 
                    -not ($InputObject -is [string])) {
                return $InputObject | ForEach-Object { ConvertTo-Hashtable $_ }
            }
            else {
                return $InputObject
            }
        }

        function Normalize-AllowList {
            param(
                [string[]] $Extensions,
                [string[]] $Publishers
            )

            $CleanExtensions = @()
            $CleanPublishers = @()

            foreach ($ext in $Extensions) {
                if ([string]::IsNullOrWhiteSpace($ext)) { continue }
                $CleanExtensions += $ext.Trim()
            }

            foreach ($pub in $Publishers) {
                if ([string]::IsNullOrWhiteSpace($pub)) { continue }

                $p = $pub.Trim()

                if ($p -match '\*$') { $p = $p.TrimEnd('*') }
                if ($p.EndsWith('.')) { $p = $p.TrimEnd('.') }

                $CleanPublishers += $p
            }

            return @{
                Extensions = $CleanExtensions
                Publishers = $CleanPublishers
            }
        }

        # Detect system-wide VS Code installation
        $SystemInstallPaths = @(
            "$env:ProgramFiles\Microsoft VS Code\Code.exe",
            "$env:ProgramFiles(x86)\Microsoft VS Code\Code.exe"
        )

        $SystemVSCodeExists = $SystemInstallPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    }

    PROCESS {

        # ForceSystemContext overrides UserContext
        if ($ForceSystemContext) {
            Write-Warning "ForceSystemContext enabled. Applying system-wide policy regardless of installation location."
            $SystemContext = $true
            $UserContext = $false
        }

        # Normalize allow lists
        $Normalized = Normalize-AllowList -Extensions $AllowedExtensions -Publishers $AllowedPublishers
        $CleanExtensions = $Normalized.Extensions
        $CleanPublishers = $Normalized.Publishers

        # Build allowed-object
        $AllowedObject = @{}
        foreach ($pub in $CleanPublishers) { $AllowedObject[$pub] = $true }
        foreach ($ext in $CleanExtensions) { $AllowedObject[$ext] = $true }

        # Remove unapproved extensions
        if ($RemoveUnapprovedExtensions) {
            $PossiblePaths = @(
                "$env:LOCALAPPDATA\Programs\Microsoft VS Code\bin\code.cmd",
                "$env:ProgramFiles\Microsoft VS Code\bin\code.cmd",
                "$env:ProgramFiles(x86)\Microsoft VS Code\bin\code.cmd"
            )

            $Script:CodeCLI = $PossiblePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

            if ($Script:CodeCLI) {
                $Installed = & $Script:CodeCLI --list-extensions

                foreach ($ext in $Installed) {
                    $IsAllowed = $false

                    if ($ext -in $CleanExtensions) { $IsAllowed = $true }

                    foreach ($pub in $CleanPublishers) {
                        if ($ext.StartsWith("$pub.", [System.StringComparison]::OrdinalIgnoreCase)) {
                            $IsAllowed = $true
                            break
                        }
                    }

                    if (-not $IsAllowed) {
                        Write-Output "Removing unapproved extension: $ext"
                        & $Script:CodeCLI --uninstall-extension $ext --force | Out-Null
                    }
                }
            }
        }

        # User context enforcement
        if ($UserContext) {
            $UserSettingsPath = "$env:APPDATA\Code\User\settings.json"

            if (-not (Test-Path $UserSettingsPath)) {
                New-Item -ItemType File -Path $UserSettingsPath -Force | Out-Null
            }

            # Load existing settings
            $ExistingSettings = @{}
            try {
                $Raw = Get-Content -Path $UserSettingsPath -Raw

                # Strip JSONC comments
                $CleanRaw = $Raw -replace '(?m)^\s*//.*$' -replace '/\*.*?\*/', ''

                if ($CleanRaw.Trim()) {
                    $ExistingSettings = $CleanRaw | ConvertFrom-Json -ErrorAction Stop
                }
            }
            catch {
                Write-Warning "settings.json contains invalid JSON. Rebuilding required keys."
                $ExistingSettings = @{}
            }

            $ExistingSettings = ConvertTo-Hashtable $ExistingSettings
            if (-not ($ExistingSettings -is [hashtable])) { $ExistingSettings = @{} }

            # Apply enforced settings
            $ExistingSettings["extensions.allowed"] = $AllowedObject
            $ExistingSettings["extensions.autoCheckUpdates"] = $AutoCheckUpdates
            $ExistingSettings["extensions.autoUpdate"] = $EnableAutoUpdate
            $ExistingSettings["extensions.ignoreRecommendations"] = $true

            # Write merged settings
            $SettingsJson = $ExistingSettings | ConvertTo-Json -Depth 10
            $SettingsJson | Out-File -FilePath $UserSettingsPath -Encoding UTF8 -Force

            Write-Output "User settings.json merged and updated."
        }

        # System context enforcement
        if ($SystemContext) {
            $PolicyPath = "HKLM:\Software\Policies\Microsoft\VSCode"
            New-Item -Path $PolicyPath -Force | Out-Null

            # JSON string for extensions.allowed
            $PolicyValue = $AllowedObject | ConvertTo-Json -Depth 5 -Compress

            $EnableAutoUpdateValue      = if ($EnableAutoUpdate) { 1 } else { 0 }
            $EnableAutoCheckUpdateValue = if ($AutoCheckUpdates) { 1 } else { 0 }

            Set-ItemProperty -Path $PolicyPath -Name "extensions.allowed"          -Value $PolicyValue -Type String
            Set-ItemProperty -Path $PolicyPath -Name "extensions.autoUpdate"       -Value $EnableAutoUpdateValue -Type DWord
            Set-ItemProperty -Path $PolicyPath -Name "extensions.autoCheckUpdates" -Value $EnableAutoCheckUpdateValue -Type DWord
            Set-ItemProperty -Path $PolicyPath -Name "extensions.gallery.enabled"  -Value 1 -Type DWord

            Write-Output "Machine-wide VS Code policy applied."
        }
    }

    END {
        Write-Output "VS Code extension enforcement completed."
    }
}
