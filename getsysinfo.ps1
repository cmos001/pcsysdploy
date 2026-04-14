[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-FirmwareMode {
    try {
        $firmwareType = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name 'PEFirmwareType' -ErrorAction SilentlyContinue).PEFirmwareType
        switch ($firmwareType) {
            1 { return 'BIOS' }
            2 { return 'UEFI' }
        }
    }
    catch {
    }

    try {
        if (Confirm-SecureBootUEFI -ErrorAction Stop) {
            return 'UEFI'
        }

        return 'UEFI'
    }
    catch {
        return 'BIOS'
    }
}

function Get-SystemDriveLetter {
    $systemDrive = (Get-CimInstance -ClassName Win32_OperatingSystem).SystemDrive
    return $systemDrive.TrimEnd('\')
}

function Get-ExistingProfilePath {
    $drives = Get-PSDrive -PSProvider FileSystem | Sort-Object Root
    foreach ($drive in $drives) {
        $candidate = Join-Path $drive.Root 'sysdeploy\profile'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    return $null
}

function Get-LargestNonSystemDriveRoot {
    param([string]$SystemDrive)

    $candidates = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" |
        Where-Object { $_.DeviceID -ne $SystemDrive } |
        Sort-Object Size -Descending

    $selected = $candidates | Select-Object -First 1
    if (-not $selected) {
        throw 'Unable to find a non-system fixed drive for sysdeploy\profile creation.'
    }

    return ($selected.DeviceID + '\')
}

function Get-ProfilePath {
    $systemDrive = Get-SystemDriveLetter
    $existing = Get-ExistingProfilePath
    if ($existing) {
        return $existing
    }

    $root = Get-LargestNonSystemDriveRoot -SystemDrive $systemDrive
    $directory = Join-Path $root 'sysdeploy'
    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }

    return (Join-Path $directory 'profile')
}

function Get-DiskInventory {
    $partitions = Get-CimInstance -ClassName Win32_DiskPartition
    $logicalDisks = Get-CimInstance -ClassName Win32_LogicalDisk
    $diskDrives = Get-CimInstance -ClassName Win32_DiskDrive

    $partitionToDrive = @{}
    foreach ($partition in $partitions) {
        $associations = @(Get-CimAssociatedInstance -InputObject $partition -Association Win32_LogicalDiskToPartition -ErrorAction SilentlyContinue)
        $partitionToDrive[$partition.DeviceID] = @($associations | ForEach-Object { $_.DeviceID })
    }

    $result = foreach ($disk in $diskDrives) {
        $diskPartitions = @($partitions | Where-Object { $_.DiskIndex -eq $disk.Index })
        $letters = foreach ($partition in $diskPartitions) {
            if ($partitionToDrive.ContainsKey($partition.DeviceID)) {
                $partitionToDrive[$partition.DeviceID]
            }
        }

        [pscustomobject]@{
            DiskNumber = [int]$disk.Index
            Model = $disk.Model
            InterfaceType = $disk.InterfaceType
            SizeBytes = [int64]$disk.Size
            Partitions = @($letters | Sort-Object -Unique)
        }
    }

    return @($result)
}

function Get-VolumeInventory {
    $volumes = Get-CimInstance -ClassName Win32_LogicalDisk | Sort-Object DeviceID
    return @(
        foreach ($volume in $volumes) {
            [pscustomobject]@{
                DriveLetter = $volume.DeviceID
                DriveType = [int]$volume.DriveType
                FileSystem = $volume.FileSystem
                SizeBytes = if ($null -ne $volume.Size) { [int64]$volume.Size } else { $null }
                FreeBytes = if ($null -ne $volume.FreeSpace) { [int64]$volume.FreeSpace } else { $null }
                VolumeName = $volume.VolumeName
            }
        }
    )
}

function Get-DefaultProfileObject {
    return [ordered]@{
        schema_version = 1
        profile_id = 'default'
        collected_utc = $null
        host = [ordered]@{
            computer_name = $null
            os_caption = $null
            os_version = $null
            os_build = $null
            architecture = $null
            firmware = $null
            system_drive = $null
            windows_directory = $null
        }
        inventory = [ordered]@{
            disks = @()
            volumes = @()
        }
        deploy = [ordered]@{
            install_mode = 'clean'
            package = [ordered]@{
                path = '.\windows.wim'
                image_index = 1
                official_iso_url = 'https://www.microsoft.com/software-download/windows11'
                official_esd_url = 'https://www.microsoft.com/software-download/windows10iso'
                official_wim_note = 'Place your extracted or captured install.wim at sysdeploy\windows.wim.'
            }
            target = [ordered]@{
                existing_windows_path = 'W:\'
                existing_system_path = 'S:\'
            }
            driver_injection = [ordered]@{
                enabled = $false
                path = '.\drivers'
            }
            boot_files = [ordered]@{
                preserve_boot_menu_order = $true
                preserve_firmware_order = $true
                add_firmware_entry_last = $false
                use_bootex = $false
            }
            restart_after_deploy = $true
        }
    }
}

function Merge-ProfileValues {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Current,

        [Parameter(Mandatory = $true)]
        [hashtable]$Template
    )

    foreach ($key in $Template.Keys) {
        if (-not $Current.ContainsKey($key)) {
            $Current[$key] = $Template[$key]
            continue
        }

        if ($Current[$key] -is [hashtable] -and $Template[$key] -is [hashtable]) {
            Merge-ProfileValues -Current $Current[$key] -Template $Template[$key]
        }
    }

    return $Current
}

function ConvertTo-Hashtable {
    param([object]$Object)

    if ($null -eq $Object) {
        return $null
    }

    if ($Object -is [hashtable]) {
        $copy = @{}
        foreach ($key in $Object.Keys) {
            $copy[$key] = ConvertTo-Hashtable -Object $Object[$key]
        }

        return $copy
    }

    if ($Object -is [System.Collections.IDictionary]) {
        $copy = @{}
        foreach ($key in $Object.Keys) {
            $copy[$key] = ConvertTo-Hashtable -Object $Object[$key]
        }

        return $copy
    }

    if ($Object -is [System.Collections.IEnumerable] -and -not ($Object -is [string])) {
        $items = @()
        foreach ($item in $Object) {
            $items += ,(ConvertTo-Hashtable -Object $item)
        }

        return $items
    }

    if ($Object.PSObject -and $Object.PSObject.Properties.Count -gt 0) {
        $copy = @{}
        foreach ($property in $Object.PSObject.Properties) {
            $copy[$property.Name] = ConvertTo-Hashtable -Object $property.Value
        }

        return $copy
    }

    return $Object
}

$profilePath = Get-ProfilePath
$defaultProfile = Get-DefaultProfileObject

if (Test-Path -LiteralPath $profilePath -PathType Leaf) {
    $existingProfile = ConvertTo-Hashtable -Object (Get-Content -LiteralPath $profilePath -Raw -Encoding UTF8 | ConvertFrom-Json)
    $profile = Merge-ProfileValues -Current $existingProfile -Template $defaultProfile
}
else {
    $profile = $defaultProfile
}

$os = Get-CimInstance -ClassName Win32_OperatingSystem
$computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem

$profile.collected_utc = (Get-Date).ToUniversalTime().ToString('o')
$profile.host.computer_name = $env:COMPUTERNAME
$profile.host.os_caption = $os.Caption
$profile.host.os_version = $os.Version
$profile.host.os_build = $os.BuildNumber
$profile.host.architecture = $computerSystem.SystemType
$profile.host.firmware = Get-FirmwareMode
$profile.host.system_drive = $os.SystemDrive
$profile.host.windows_directory = $env:WINDIR
$profile.inventory.disks = @(Get-DiskInventory)
$profile.inventory.volumes = @(Get-VolumeInventory)

$profileDirectory = Split-Path -Parent $profilePath
if ($profileDirectory -and -not (Test-Path -LiteralPath $profileDirectory)) {
    New-Item -Path $profileDirectory -ItemType Directory -Force | Out-Null
}

$profile | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $profilePath -Encoding UTF8
Write-Host "Profile updated: $profilePath"
