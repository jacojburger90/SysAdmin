<#
.SYNOPSIS
    Provisions VHDX disks for a SQL Server VM on Hyper-V host.

.DESCRIPTION
    Creates structured VHDX files for SQL Server VMs following the standard
    Server drive layout:
        C:\ - OS       (SSD-backed, from template)
        D:\ - Data     (SSD-backed, fixed, 4K sector)
        E:\ - Documents (SAS-backed, fixed, 4K sector)
        L:\ - Logs     (SAS-backed, fixed, 4K sector)
        T:\ - TempDB   (SSD-backed, fixed, 4K sector)

    Output naming convention: <ServerName>_<DriveLetter>.vhdx
    Example: Server-Name_C.vhdx

.PARAMETER ServerName
    Target VM name (e.g. Server-Name). Used in VHDX naming.

.PARAMETER SizesGB
    Hashtable overriding default disk sizes in GB.
    Keys: C, D, E, L, T
    Example: @{ D = 200; T = 50 }

.PARAMETER TemplateDisk
    Full path to the OS template VHDX. Copied and renamed to <ServerName>_C.vhdx.

.PARAMETER SSDBase
    Base cluster storage path for SSD-backed disks.

.PARAMETER SASBase
    Base cluster storage path for SAS-backed disks.

.PARAMETER LogPath
    Directory for the CSV log file. Defaults to C:\Temp.

.PARAMETER WhatIf
    Dry-run mode. Shows what would be created without making changes.

.EXAMPLE
    .\New-SQLServerDisks.ps1 -ServerName "Server-Name"

.EXAMPLE
    .\New-SQLServerDisks.ps1 -ServerName "Server-Name" -SizesGB @{ D = 200; T = 50 } -WhatIf

.NOTES
    Version     : 2.0
    Author      : Jacobus J. Burger
    Requires    : Hyper-V PowerShell module, Run as Administrator
    Sector size : 4096 logical/physical (SQL Server best practice for VHDX)
    Template    : Must be a sysprepped VHDX located on SSD storage
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[A-Za-z0-9\-]{3,}$')]
    [string]$ServerName,

    [Parameter(Mandatory = $false)]
    [hashtable]$SizesGB = @{},

    [Parameter(Mandatory = $false)]
    [string]$TemplateDisk = "C:\ClusterStorage\Volume2\virtual_disks\template\db-template_C.vhdx",

    [Parameter(Mandatory = $false)]
    [string]$SSDBase = "C:\ClusterStorage\Volume2\virtual_disks",

    [Parameter(Mandatory = $false)]
    [string]$SASBase = "C:\ClusterStorage\Volume1\virtual_disks",

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Temp"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region -- Logging --------------------------------------------------------------
$LogFile = Join-Path $LogPath "DiskProvision_$($ServerName)_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
$Script:Log = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-Log {
    param (
        [string]$Step,
        [ValidateSet("INFO","SUCCESS","WARN","FAILED")]
        [string]$Status,
        [string]$Message
    )
    $entry = [PSCustomObject]@{
        Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Step      = $Step
        Status    = $Status
        Message   = $Message
    }
    $Script:Log.Add($entry)
    $colour = switch ($Status) {
        "SUCCESS" { "Green"  }
        "WARN"    { "Yellow" }
        "FAILED"  { "Red"    }
        default   { "Cyan"   }
    }
    Write-Host "[$Status] $Step -- $Message" -ForegroundColor $colour
}
#endregion

#region -- Pre-flight -----------------------------------------------------------
function Test-Prerequisites {
    Write-Log "Pre-flight" "INFO" "Checking prerequisites"

    # Must run elevated
    $currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Script must run as Administrator."
    }

    # Hyper-V module
    if (-not (Get-Module -ListAvailable -Name Hyper-V)) {
        throw "Hyper-V PowerShell module not available. Install RSAT-Hyper-V-Tools."
    }
    Import-Module Hyper-V -ErrorAction Stop
    Write-Log "Pre-flight" "SUCCESS" "Hyper-V module loaded"

    # Template disk exists
    if (-not (Test-Path $TemplateDisk)) {
        throw "Template disk not found: $TemplateDisk"
    }
    Write-Log "Pre-flight" "SUCCESS" "Template disk found: $TemplateDisk"

    # Log directory
    if (-not (Test-Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
    }
}
#endregion

#region -- Disk Definitions -----------------------------------------------------
#
# Drive layout:
#   C  OS       SSD  -- template copy (no New-VHD, just Copy + Rename)
#   D  Data     SSD  -- SQL data files
#   E  Documents SAS  -- application / archive data
#   L  Logs     SAS  -- SQL transaction logs (sequential write, SAS acceptable)
#   T  TempDB   SSD  -- TempDB (high IOPS, must be SSD)
#
# Sizes are defaults; override via -SizesGB hashtable parameter.
# All new VHDs: Fixed provisioning, 4096b logical/physical sectors (SQL best practice).
#
function Get-DiskDefinitions {
    # Defaults
    $defaults = @{
        C = 100   # OS  -- set to your template size
        D = 100   # Data
        E = 100   # Documents
        L = 100   # Logs
        T = 100   # TempDB
    }

    # Merge caller overrides
    foreach ($key in $SizesGB.Keys) {
        $k = $key.ToUpper()
        if ($defaults.ContainsKey($k)) {
            $defaults[$k] = $SizesGB[$key]
            Write-Log "Disk Definitions" "INFO" "Override: Drive $k -> $($SizesGB[$key]) GB"
        } else {
            Write-Log "Disk Definitions" "WARN" "Unknown drive letter override ignored: $key"
        }
    }

    $ssd = Join-Path $SSDBase $ServerName
    $sas = Join-Path $SASBase $ServerName

    return @(
        [PSCustomObject]@{
            Drive       = "C"
            Label       = "OS"
            StorageTier = "SSD"
            Folder      = $ssd
            IsTemplate  = $true
            SizeGB      = $defaults["C"]
        },
        [PSCustomObject]@{
            Drive       = "D"
            Label       = "Data"
            StorageTier = "SSD"
            Folder      = $ssd
            IsTemplate  = $false
            SizeGB      = $defaults["D"]
        },
        [PSCustomObject]@{
            Drive       = "E"
            Label       = "Documents"
            StorageTier = "SAS"
            Folder      = $sas
            IsTemplate  = $false
            SizeGB      = $defaults["E"]
        },
        [PSCustomObject]@{
            Drive       = "L"
            Label       = "Logs"
            StorageTier = "SAS"
            Folder      = $sas
            IsTemplate  = $false
            SizeGB      = $defaults["L"]
        },
        [PSCustomObject]@{
            Drive       = "T"
            Label       = "TempDB"
            StorageTier = "SSD"
            Folder      = $ssd
            IsTemplate  = $false
            SizeGB      = $defaults["T"]
        }
    )
}
#endregion

#region -- Disk Creation --------------------------------------------------------
function New-ServerDisks {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [PSCustomObject[]]$Disks
    )

    # Collect unique folders and create them
    $folders = $Disks | Select-Object -ExpandProperty Folder -Unique
    foreach ($folder in $folders) {
        if ($PSCmdlet.ShouldProcess($folder, "Create directory")) {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
            Write-Log "Create Folder" "SUCCESS" "$folder"
        } else {
            Write-Log "Create Folder" "INFO" "[WhatIf] Would create: $folder"
        }
    }

    foreach ($disk in $Disks) {
        $vhdxName = "$($ServerName)_$($disk.Drive).vhdx"
        $vhdxPath = Join-Path $disk.Folder $vhdxName

        # Guard: never overwrite an existing VHDX regardless of how the function is called
        if (Test-Path $vhdxPath) {
            Write-Log "Disk $($disk.Drive)" "WARN" "SKIPPED - already exists: $vhdxPath"
            continue
        }

        try {
            if ($disk.IsTemplate) {
                # -- C: Copy sysprep template (once only - guard above ensures
                #       this block is never reached if destination exists) -----
                $templateItem = Get-Item -Path $TemplateDisk -ErrorAction Stop
                if ($templateItem.Length -eq 0) {
                    throw "Template disk is 0 bytes: $TemplateDisk"
                }

                if ($PSCmdlet.ShouldProcess($vhdxPath, "Copy sysprep template ($([math]::Round($templateItem.Length / 1GB, 2)) GB)")) {
                    Copy-Item -Path $TemplateDisk -Destination $vhdxPath
                    $copiedItem = Get-Item -Path $vhdxPath -ErrorAction Stop
                    if ($copiedItem.Length -ne $templateItem.Length) {
                        throw "Copy size mismatch. Source: $($templateItem.Length) bytes, Destination: $($copiedItem.Length) bytes"
                    }
                    Write-Log "Disk C (OS)" "SUCCESS" "Sysprep template copied -> $vhdxPath [$([math]::Round($copiedItem.Length / 1GB, 2)) GB]"
                } else {
                    Write-Log "Disk C (OS)" "INFO" "[WhatIf] Would copy sysprep template ($([math]::Round($templateItem.Length / 1GB, 2)) GB) -> $vhdxPath"
                }
            } else {
                # -- D/E/L/T: Create fixed VHDX ------------------------------
                $sizeBytes = [int64]$disk.SizeGB * 1GB
                if ($PSCmdlet.ShouldProcess($vhdxPath, "New-VHD $($disk.SizeGB)GB Fixed")) {
                    New-VHD -Path                    $vhdxPath `
                            -SizeBytes               $sizeBytes `
                            -Fixed `
                            -LogicalSectorSizeBytes  4096 `
                            -PhysicalSectorSizeBytes 4096 `
                            -ErrorAction             Stop | Out-Null
                    Write-Log "Disk $($disk.Drive) ($($disk.Label))" "SUCCESS" `
                        "$($disk.SizeGB)GB Fixed [$($disk.StorageTier)] -> $vhdxPath"
                } else {
                    Write-Log "Disk $($disk.Drive) ($($disk.Label))" "INFO" `
                        "[WhatIf] Would create $($disk.SizeGB)GB Fixed -> $vhdxPath"
                }
            }
        } catch {
            Write-Log "Disk $($disk.Drive) ($($disk.Label))" "FAILED" $_.Exception.Message
        }
    }
}
#endregion

#region Summary Report
function Show-Summary {
    param ([PSCustomObject[]]$Disks)

    Write-Host "`n--- Disk Provision Summary -----------------------------------" -ForegroundColor Cyan
    Write-Host ("{0,-6} {1,-12} {2,-8} {3,-6} {4}" -f "Drive","Label","Tier","GB","Path") -ForegroundColor White
    Write-Host ("{0,-6} {1,-12} {2,-8} {3,-6} {4}" -f "-----","----------","--------","--","------------------------------") -ForegroundColor DarkGray

    foreach ($disk in $Disks) {
        $vhdxPath = Join-Path $disk.Folder "$($ServerName)_$($disk.Drive).vhdx"
        Write-Host ("{0,-6} {1,-12} {2,-8} {3,-6} {4}" -f `
            "$($disk.Drive):\", $disk.Label, $disk.StorageTier, $disk.SizeGB, $vhdxPath)
    }
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor DarkGray
}
#endregion

#region -- Main -----------------------------------------------------------------
try {
    Write-Host "`n=== Hyper-V Disk Provisioning -- $ServerName ===`n" -ForegroundColor Cyan

    Test-Prerequisites

    $disks = Get-DiskDefinitions
    Show-Summary -Disks $disks
    New-ServerDisks -Disks $disks

    # Export log
    $Script:Log | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8
    Write-Log "Log Export" "SUCCESS" "Log written to $LogFile"

    $failed = @($Script:Log | Where-Object { $_.Status -eq "FAILED" })
    if ($failed.Count -gt 0) {
        Write-Host "`n[$($failed.Count) FAILURES -- review log: $LogFile]" -ForegroundColor Red
        exit 1
    } else {
        Write-Host "`n[All disks provisioned successfully]`n" -ForegroundColor Green
        exit 0
    }
} catch {
    Write-Log "FATAL" "FAILED" $_.Exception.Message
    Write-Host "`nFATAL: $($_.Exception.Message)" -ForegroundColor Red
    if ($Script:Log.Count -gt 0) {
        $Script:Log | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8
        Write-Host "Partial log saved to: $LogFile" -ForegroundColor Yellow
    }
    exit 1
}
#endregion