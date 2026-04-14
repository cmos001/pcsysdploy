[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:LogPath = 'X:\Windows\Temp\AutoDeploy.log'
$script:MountedIsoPath = $null

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    try {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    }
    catch {
    }
}

function Invoke-NativeCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList
    )

    Write-Log ("Executing: {0} {1}" -f $FilePath, ($ArgumentList -join ' '))
    $process = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -Wait -PassThru -NoNewWindow
    if ($process.ExitCode -ne 0) {
        throw ("Command failed with exit code {0}: {1}" -f $process.ExitCode, $FilePath)
    }
}

function Find-ProfilePath {
    $drives = Get-PSDrive -PSProvider FileSystem | Sort-Object Root
    foreach ($drive in $drives) {
        $candidate = Join-Path $drive.Root 'sysdeploy\profile'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    throw 'Unable to find the first valid <Drive>:\sysdeploy\profile file.'
}

function Resolve-RelativePath {
    param(
        [string]$BaseDirectory,
        [string]$PathValue
    )

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $null
    }

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }

    return Join-Path $BaseDirectory $PathValue
}

function Get-FirmwareMode {
    try {
        Invoke-NativeCommand -FilePath 'wpeutil.exe' -ArgumentList @('UpdateBootInfo')
    }
    catch {
    }

    $firmwareType = $null
    try {
        $firmwareType = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name 'PEFirmwareType' -ErrorAction Stop).PEFirmwareType
    }
    catch {
    }

    switch ($firmwareType) {
        1 { return 'BIOS' }
        2 { return 'UEFI' }
        default { throw 'Unable to determine firmware mode from WinPE.' }
    }
}

function Assert-InstallMode {
    param([object]$Profile)

    $installMode = [string]$Profile.deploy.install_mode
    if ([string]::IsNullOrWhiteSpace($installMode)) {
        throw 'deploy.install_mode is required in profile.'
    }

    if ($installMode.ToLowerInvariant() -ne 'clean') {
        throw 'Only deploy.install_mode=clean is supported. WinPE does not support in-place upgrade or no-loss setup flows.'
    }
}

function Assert-TargetSettings {
    param([object]$Profile)

    $target = $Profile.deploy.target
    if (-not $target) {
        throw 'deploy.target is required in profile.'
    }

    if ([string]::IsNullOrWhiteSpace([string]$target.existing_windows_path)) {
        throw 'deploy.target.existing_windows_path is required.'
    }

    if ([string]::IsNullOrWhiteSpace([string]$target.existing_system_path)) {
        throw 'deploy.target.existing_system_path is required.'
    }
}

function Resolve-PackagePath {
    param(
        [object]$Profile,
        [string]$SysDeployRoot
    )

    $configured = Resolve-RelativePath -BaseDirectory $SysDeployRoot -PathValue ([string]$Profile.deploy.package.path)
    $candidates = @()
    if ($configured) {
        $candidates += $configured
    }

    $candidates += @(
        (Join-Path $SysDeployRoot 'windows.wim'),
        (Join-Path $SysDeployRoot 'windows.esd'),
        (Join-Path $SysDeployRoot 'windows.iso')
    )

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    throw 'No deployment package found. Put windows.wim, windows.esd, or windows.iso under sysdeploy\.'
}

function Resolve-ImageSource {
    param(
        [string]$PackagePath,
        [int]$ImageIndex
    )

    $extension = [System.IO.Path]::GetExtension($PackagePath).ToLowerInvariant()
    switch ($extension) {
        '.wim' {
            return [pscustomobject]@{
                ImagePath  = $PackagePath
                ImageIndex = $ImageIndex
                IsIso      = $false
            }
        }
        '.esd' {
            return [pscustomobject]@{
                ImagePath  = $PackagePath
                ImageIndex = $ImageIndex
                IsIso      = $false
            }
        }
        '.iso' {
            $image = Mount-DiskImage -ImagePath $PackagePath -PassThru
            $script:MountedIsoPath = $PackagePath
            $volume = $image | Get-Volume | Select-Object -First 1
            if (-not $volume) {
                throw "Mounted ISO has no accessible volume: $PackagePath"
            }

            $root = $volume.DriveLetter + ':\'
            $wim = Join-Path $root 'sources\install.wim'
            $esd = Join-Path $root 'sources\install.esd'

            if (Test-Path -LiteralPath $wim -PathType Leaf) {
                return [pscustomobject]@{
                    ImagePath  = $wim
                    ImageIndex = $ImageIndex
                    IsIso      = $true
                }
            }

            if (Test-Path -LiteralPath $esd -PathType Leaf) {
                return [pscustomobject]@{
                    ImagePath  = $esd
                    ImageIndex = $ImageIndex
                    IsIso      = $true
                }
            }

            throw "No install.wim or install.esd found inside ISO: $PackagePath"
        }
        default {
            throw "Unsupported package type: $extension"
        }
    }
}

function Resolve-ExistingVolumes {
    param([object]$Target)

    return [pscustomobject]@{
        WindowsVolume = ([string]$Target.existing_windows_path).TrimEnd('\') + '\'
        SystemVolume  = ([string]$Target.existing_system_path).TrimEnd('\') + '\'
    }
}

function Format-TargetWindowsVolume {
    param([string]$WindowsVolume)

    $driveRoot = $WindowsVolume.TrimEnd('\')
    $driveLetter = $driveRoot.TrimEnd(':')

    if (-not (Test-Path -LiteralPath $WindowsVolume)) {
        throw "Target Windows volume not accessible: $WindowsVolume"
    }

    Write-Log "Formatting target Windows volume only: $driveRoot"
    Invoke-NativeCommand -FilePath 'format.com' -ArgumentList @(
        $driveRoot
        '/FS:NTFS'
        '/Q'
        '/Y'
        '/V:Windows'
    )
}

function Apply-WindowsImage {
    param(
        [string]$ImagePath,
        [int]$ImageIndex,
        [string]$WindowsVolume
    )

    Invoke-NativeCommand -FilePath 'dism.exe' -ArgumentList @(
        '/Apply-Image'
        "/ImageFile:$ImagePath"
        "/Index:$ImageIndex"
        "/ApplyDir:$WindowsVolume"
        '/CheckIntegrity'
    )

    $windowsDir = Join-Path $WindowsVolume 'Windows'
    if (-not (Test-Path -LiteralPath $windowsDir -PathType Container)) {
        throw "Windows directory not found after apply: $windowsDir"
    }

    return $windowsDir
}

function Inject-Drivers {
    param(
        [object]$Profile,
        [string]$SysDeployRoot,
        [string]$WindowsVolume
    )

    if (-not [bool]$Profile.deploy.driver_injection.enabled) {
        Write-Log 'Driver injection disabled.'
        return
    }

    $driverPath = Resolve-RelativePath -BaseDirectory $SysDeployRoot -PathValue ([string]$Profile.deploy.driver_injection.path)
    if (-not (Test-Path -LiteralPath $driverPath -PathType Container)) {
        throw "Driver directory not found: $driverPath"
    }

    Invoke-NativeCommand -FilePath 'dism.exe' -ArgumentList @(
        "/Image:$WindowsVolume"
        '/Add-Driver'
        "/Driver:$driverPath"
        '/Recurse'
    )
}

function Get-InstalledWindowsCandidates {
    $drives = Get-PSDrive -PSProvider FileSystem | Sort-Object Root
    $results = foreach ($drive in $drives) {
        $windowsDir = Join-Path $drive.Root 'Windows'
        $bootMgrPath = Join-Path $drive.Root 'EFI\Microsoft\Boot\bootmgfw.efi'

        if (Test-Path -LiteralPath $windowsDir -PathType Container) {
            [pscustomobject]@{
                Root = $drive.Root
                WindowsDirectory = $windowsDir
                BootManagerPath = if (Test-Path -LiteralPath $bootMgrPath -PathType Leaf) { $bootMgrPath } else { $null }
            }
        }
    }

    return @($results)
}

function Test-EfiCompatibility {
    param(
        [string]$FirmwareMode,
        [string]$SystemVolume
    )

    if ($FirmwareMode -ne 'UEFI') {
        Write-Log 'Firmware mode is BIOS; EFI compatibility check skipped.'
        return
    }

    $sharedBootManager = Join-Path $SystemVolume 'EFI\Microsoft\Boot\bootmgfw.efi'
    $otherWindows = Get-InstalledWindowsCandidates

    if (-not (Test-Path -LiteralPath $sharedBootManager -PathType Leaf)) {
        Write-Log -Level WARN -Message "No existing EFI boot manager found at $sharedBootManager. BCDBoot will create or refresh it."
        return
    }

    $bootVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($sharedBootManager).FileVersion
    Write-Log "Existing shared EFI boot manager version: $bootVersion"

    if ($otherWindows.Count -gt 1) {
        Write-Log -Level WARN -Message 'Multiple Windows installations were detected. Windows 10 and Windows 11 commonly share a single Windows Boot Manager on the ESP.'
        Write-Log -Level WARN -Message 'The supported servicing path is to let BCDBoot refresh the shared boot environment files from the deployed Windows image, rather than manually dropping arbitrary bootx64.efi or bootmgfw.efi files onto the ESP.'
    }
}

function Write-BootFiles {
    param(
        [string]$WindowsDirectory,
        [string]$SystemVolume,
        [string]$FirmwareMode,
        [object]$Profile
    )

    $preserveBootMenuOrder = $true
    $preserveFirmwareOrder = $true
    $addFirmwareEntryLast = $false
    $useBootEx = $false

    if ($Profile.deploy.PSObject.Properties.Name -contains 'boot_files') {
        $bootFiles = $Profile.deploy.boot_files
        if ($bootFiles) {
            if ($bootFiles.PSObject.Properties.Name -contains 'preserve_boot_menu_order') {
                $preserveBootMenuOrder = [bool]$bootFiles.preserve_boot_menu_order
            }
            if ($bootFiles.PSObject.Properties.Name -contains 'preserve_firmware_order') {
                $preserveFirmwareOrder = [bool]$bootFiles.preserve_firmware_order
            }
            if ($bootFiles.PSObject.Properties.Name -contains 'add_firmware_entry_last') {
                $addFirmwareEntryLast = [bool]$bootFiles.add_firmware_entry_last
            }
            if ($bootFiles.PSObject.Properties.Name -contains 'use_bootex') {
                $useBootEx = [bool]$bootFiles.use_bootex
            }
        }
    }

    if ($preserveFirmwareOrder -and $addFirmwareEntryLast) {
        throw 'boot_files.preserve_firmware_order and boot_files.add_firmware_entry_last cannot both be true.'
    }

    $arguments = @(
        $WindowsDirectory
        '/s'
        $SystemVolume.TrimEnd('\')
        '/f'
        $FirmwareMode
    )

    if ($preserveBootMenuOrder) {
        $arguments += '/d'
    }

    if ($preserveFirmwareOrder) {
        $arguments += '/p'
    }
    elseif ($addFirmwareEntryLast) {
        $arguments += '/addlast'
    }

    if ($useBootEx) {
        $arguments += '/bootex'
    }

    Write-Log 'Adding or repairing the Windows boot entry with BCDBoot using Microsoft-recommended preservation flags for existing boot order.'
    Invoke-NativeCommand -FilePath 'bcdboot.exe' -ArgumentList $arguments
}

function Dismount-ResolvedIso {
    if ($script:MountedIsoPath) {
        try {
            Dismount-DiskImage -ImagePath $script:MountedIsoPath | Out-Null
        }
        catch {
            Write-Log -Level WARN -Message "Failed to dismount ISO: $script:MountedIsoPath"
        }
    }
}

try {
    $profilePath = Find-ProfilePath
    $sysDeployRoot = Split-Path -Parent $profilePath
    Write-Log "Using profile: $profilePath"

    $profile = Get-Content -LiteralPath $profilePath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-InstallMode -Profile $profile
    Assert-TargetSettings -Profile $profile

    $firmwareMode = Get-FirmwareMode
    Write-Log "Detected firmware mode: $firmwareMode"

    $packagePath = Resolve-PackagePath -Profile $profile -SysDeployRoot $sysDeployRoot
    Write-Log "Deployment package: $packagePath"

    $imageIndex = 1
    if ($profile.deploy.package.image_index) {
        $imageIndex = [int]$profile.deploy.package.image_index
    }

    $imageSource = Resolve-ImageSource -PackagePath $packagePath -ImageIndex $imageIndex
    $volumes = Resolve-ExistingVolumes -Target $profile.deploy.target

    Test-EfiCompatibility -FirmwareMode $firmwareMode -SystemVolume $volumes.SystemVolume
    Format-TargetWindowsVolume -WindowsVolume $volumes.WindowsVolume
    $windowsDirectory = Apply-WindowsImage -ImagePath $imageSource.ImagePath -ImageIndex $imageSource.ImageIndex -WindowsVolume $volumes.WindowsVolume
    Inject-Drivers -Profile $profile -SysDeployRoot $sysDeployRoot -WindowsVolume $volumes.WindowsVolume
    Write-BootFiles -WindowsDirectory $windowsDirectory -SystemVolume $volumes.SystemVolume -FirmwareMode $firmwareMode -Profile $profile

    Write-Log 'Deployment completed successfully.'
    if ($profile.deploy.restart_after_deploy -ne $false) {
        Invoke-NativeCommand -FilePath 'wpeutil.exe' -ArgumentList @('reboot')
    }
}
catch {
    Write-Log -Level ERROR -Message $_.Exception.Message
    exit 1
}
finally {
    Dismount-ResolvedIso
}
