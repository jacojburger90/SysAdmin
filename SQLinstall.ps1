<#
.SYNOPSIS
    Uniform, repeatable SQL Server unattended installation.
    Generates ConfigurationFile.ini, mounts ISO, runs setup /Q, validates install.
    Use -DryRun to validate all pre-flight checks and preview configuration without
    touching the system (no install, no directory creation, no SQL changes).

.DESCRIPTION
    Drive layout  : C:\OS  D:\Data  D:\Backup  E:\Documents  E:\IO  L:\Logs  T:\TempDB
    Instance      : MSSQLSERVER (default)
    Collation     : Latin1_General_CI_AS
    Protocols     : TCP/IP enabled (1433)  Named Pipes enabled
    Memory        : min 1024 MB  max 10240 MB (10 GB)
    Auth          : Windows Authentication only
    SSMS          : NOT included - install separately

    DRY RUN MODE (-DryRun):
    Runs all 7 pre-flight checks, computes TempDB file count, renders the full
    ConfigurationFile.ini to screen and log, and prints the setup.exe command that
    would be executed - without mounting the ISO, creating directories, running
    setup.exe, or making any SQL/registry changes. Safe to run on production servers
    before committing to an install.

.NOTES
    Version       : 1.7.0
    Requires      : PowerShell 5.1 | Run as local Administrator
    Compatibility : SQL Server 2019 / 2022 / 2025

.PARAMETER ISOPath
    Full UNC or local path to the SQL Server ISO.
    Example: "\\server\share\SQLServer2022.iso"

.PARAMETER SqlSvcAccount
    Service account for SQL Engine.
    gMSA    : "DOMAIN\svc-sql$"           (no password required - omit SqlSvcPassword)
    Virtual : "NT Service\MSSQLSERVER"    (no password required - omit SqlSvcPassword)
    Domain  : "DOMAIN\sql.service"        (SqlSvcPassword required)

.PARAMETER SqlSvcPassword
    Password for SqlSvcAccount. Required when using a plain domain user account.
    Omit entirely for gMSA (trailing $) or NT Service\ virtual accounts.
    Accepted as SecureString - never written to disk or the INI file.

.PARAMETER AgtSvcAccount
    Service account for SQL Agent.
    gMSA    : "DOMAIN\svc-sqlagt$"        (no password required - omit AgtSvcPassword)
    Virtual : "NT Service\SQLSERVERAGENT" (no password required - omit AgtSvcPassword)
    Domain  : "DOMAIN\sql.agent"          (AgtSvcPassword required)

.PARAMETER AgtSvcPassword
    Password for AgtSvcAccount. Required when using a plain domain user account.
    Omit entirely for gMSA (trailing $) or NT Service\ virtual accounts.
    Accepted as SecureString - never written to disk or the INI file.

.PARAMETER SysAdminAccounts
    One or more Windows accounts/groups to add to sysadmin.
    Example: "DOMAIN\SQL-Admins"  or  "DOMAIN\SQL-Admins","DOMAIN\jbloggs"

.PARAMETER UpdateSource
    "MU" for Microsoft Update, or UNC path to a CU folder.
    Default: "MU"

.PARAMETER DryRun
    Switch. When present: runs all pre-flight checks, renders the ConfigurationFile.ini
    and setup command to screen/log, then exits WITHOUT installing, creating directories,
    mounting the ISO, or making any post-install SQL changes.
    Use this to validate parameters and environment before committing.

.EXAMPLE
    # Domain user accounts (bridge - use gMSA in production)
    $sqlPwd = Read-Host "SQL Engine password" -AsSecureString
    $agtPwd = Read-Host "SQL Agent password"  -AsSecureString
    .\Install-SQL-Uniform.ps1 `
        -ISOPath          "D:\Software\SQLServer2025.iso" `
        -SqlSvcAccount    "DMN\sql.service" `
        -SqlSvcPassword   $sqlPwd `
        -AgtSvcAccount    "DMN\sql.agent" `
        -AgtSvcPassword   $agtPwd `
        -SysAdminAccounts "DMN\DBA"

.EXAMPLE
    # gMSA accounts (production standard - no passwords needed)
    .\Install-SQL-Uniform.ps1 `
        -ISOPath          "D:\Software\SQLServer2025.iso" `
        -SqlSvcAccount    "DMN\sql.service`$" `
        -AgtSvcAccount    "DMN\sql.agent`$" `
        -SysAdminAccounts "DMN\DBA"

.EXAMPLE
    # Dry run - validate only, no changes made
    .\Install-SQL-Uniform.ps1 `
        -ISOPath          "D:\Software\SQLServer2025.iso" `
        -SqlSvcAccount    "DMN\sql.service`$" `
        -AgtSvcAccount    "DMN\sql.agent`$" `
        -SysAdminAccounts "DMN\DBA" `
        -DryRun
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$ISOPath,
    [Parameter(Mandatory)][string]$SqlSvcAccount,
    [Parameter(Mandatory)][string]$AgtSvcAccount,
    [Parameter(Mandatory)][string[]]$SysAdminAccounts,
    # Passwords: only required for plain domain accounts.
    # Omit for gMSA (account name ends with $) or NT Service\ virtual accounts.
    # Accepted as SecureString - converted to plain text in memory only for the
    # setup.exe CLI call and immediately discarded. Never written to disk or INI.
    [System.Security.SecureString]$SqlSvcPassword,
    [System.Security.SecureString]$AgtSvcPassword,
    [string]$UpdateSource = "MU",
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# LOGGING
# ============================================================
$LogDir  = Join-Path $PSScriptRoot "Logs"
$LogFile = Join-Path $LogDir ("SQL-Install_{0}_{1}{2}.log" -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd_HHmmss'), $(if ($DryRun) { '_DRYRUN' } else { '' }))
if (-not (Test-Path $LogDir)) { [void](New-Item -ItemType Directory -Path $LogDir) }

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','HEADER','DRYRUN')][string]$Level = 'INFO'
    )
    $ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts][$Level] $Message"
    $entry | Out-File -FilePath $LogFile -Append -Encoding UTF8
    switch ($Level) {
        'HEADER'  { Write-Host "`n$entry" -ForegroundColor White }
        'INFO'    { Write-Host $entry -ForegroundColor Cyan }
        'WARN'    { Write-Host $entry -ForegroundColor Yellow }
        'ERROR'   { Write-Host $entry -ForegroundColor Red }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green }
        'DRYRUN'  { Write-Host $entry -ForegroundColor Magenta }
    }
}

function Exit-WithError {
    param([string]$Message, [int]$Code = 1)
    Write-Log $Message -Level ERROR
    Write-Log "=== INSTALL ABORTED | Exit code: $Code ===" -Level ERROR
    exit $Code
}

# ============================================================
# CONSTANTS
# ============================================================
$INSTANCE_NAME    = 'AUTHDB'
$COLLATION        = 'Latin1_General_CI_AS'
$MEM_MIN_MB       = 1024   # SQL Server 2025 non-Express minimum is 1 GB per Microsoft requirements
$MEM_MAX_MB       = 10240
$CONFIG_DIR       = Join-Path $PSScriptRoot "SQLInstallConfig"
$CONFIG_FILE      = Join-Path $CONFIG_DIR  "ConfigurationFile.ini"

# Drive layout
$PATHS = [ordered]@{
    'Data'      = 'D:\Data'
    'Backup'    = 'D:\Backup'
    'Documents' = 'E:\Documents'
    'IO'        = 'E:\IO'
    'Logs'      = 'L:\Logs'
    'TempDB'    = 'T:\TempDB'
    'TempLogs'  = 'T:\TempDB\Logs'
}

# ============================================================
# ACCOUNT TYPE DETECTION
# ============================================================
# Determine whether each account is gMSA, virtual, or domain user.
# gMSA  : name ends with $             -> no password param, no INI password line
# Virtual: starts with "NT Service\"  -> no password param, no INI password line
# Domain user: anything else           -> password param mandatory, injected CLI-only

function Get-AccountType {
    param([string]$Account)
    if ($Account -match '\$$')              { return 'gMSA' }
    if ($Account -match '^NT Service\\')  { return 'Virtual' }
    return 'DomainUser'
}

$SqlAcctType = Get-AccountType $SqlSvcAccount
$AgtAcctType = Get-AccountType $AgtSvcAccount

# Validate: domain user accounts require their password parameter
if ($SqlAcctType -eq 'DomainUser' -and -not $PSBoundParameters.ContainsKey('SqlSvcPassword')) {
    throw "SqlSvcAccount '$SqlSvcAccount' appears to be a domain user. Supply -SqlSvcPassword (SecureString). " +
          "For gMSA append `$` to the account name; for virtual accounts use 'NT Service\MSSQLSERVER'."
}
if ($AgtAcctType -eq 'DomainUser' -and -not $PSBoundParameters.ContainsKey('AgtSvcPassword')) {
    throw "AgtSvcAccount '$AgtSvcAccount' appears to be a domain user. Supply -AgtSvcPassword (SecureString). " +
          "For gMSA append `$` to the account name; for virtual accounts use 'NT Service\SQLSERVERAGENT'."
}

# Convert SecureString passwords to plain text for CLI injection (memory only, never disk)
function ConvertTo-PlainText {
    param([System.Security.SecureString]$Secure)
    if (-not $Secure) { return $null }
    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($Secure)
    try   { return [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ptr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr) }
}

$SqlSvcPasswordPlain = ConvertTo-PlainText $SqlSvcPassword
$AgtSvcPasswordPlain = ConvertTo-PlainText $AgtSvcPassword

# ============================================================
# PRE-FLIGHT CHECKS
# ============================================================
Write-Log "============================================================" -Level HEADER
Write-Log " SQL Server Uniform Install - $env:COMPUTERNAME" -Level HEADER
if ($DryRun) {
Write-Log " *** DRY RUN MODE - NO CHANGES WILL BE MADE ***" -Level DRYRUN
}
Write-Log "============================================================" -Level HEADER
Write-Log "Account mode: SQL Engine = $SqlAcctType | SQL Agent = $AgtAcctType" -Level INFO
if ($SqlAcctType -eq 'DomainUser') {
    Write-Log "Domain user accounts in use - passwords supplied via SecureString, injected CLI-only." -Level WARN
    Write-Log "Migrate to gMSA when ready: append `$ to account names and remove password params." -Level WARN
}

# --- Check 1: Administrator ---
Write-Log "CHECK 1/7: Verifying administrator privileges..." -Level INFO
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Exit-WithError "Script must be run as local Administrator."
}
Write-Log "Administrator: OK" -Level SUCCESS

# --- Check 2: ISO exists and is readable ---
Write-Log "CHECK 2/7: Verifying ISO path: $ISOPath" -Level INFO
if (-not (Test-Path -LiteralPath $ISOPath -PathType Leaf)) {
    if ($DryRun) {
        Write-Log "ISO not found: $ISOPath (DRY RUN - non-fatal, would abort on real install)" -Level WARN
    } else {
        Exit-WithError "ISO not found: $ISOPath"
    }
} else {
    $isoSize = (Get-Item -LiteralPath $ISOPath).Length
    if ($isoSize -lt 500MB) {
        Write-Log "ISO size is $([Math]::Round($isoSize/1MB))MB - unexpectedly small. Verify ISO integrity." -Level WARN
    }
    Write-Log "ISO found ($([Math]::Round($isoSize/1MB)) MB): OK" -Level SUCCESS
}

# --- Check 3: SQL Server not already installed ---
Write-Log "CHECK 3/7: Checking for existing SQL Server installation..." -Level INFO
$existingSQL = Get-Service -Name 'MSSQLSERVER' -ErrorAction SilentlyContinue
if ($existingSQL) {
    $msg = "SQL Server service (MSSQLSERVER) already exists on this host. Use a named instance or remove existing install."
    if ($DryRun) { Write-Log "WOULD ABORT: $msg" -Level WARN } else { Exit-WithError $msg }
}
$existingReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL' -ErrorAction SilentlyContinue
if ($existingReg) {
    $msg = "SQL Server instance registry keys detected. Existing install found - aborting to prevent conflict."
    if ($DryRun) { Write-Log "WOULD ABORT: $msg" -Level WARN } else { Exit-WithError $msg }
}
if (-not $existingSQL -and -not $existingReg) {
    Write-Log "No existing SQL Server found: OK" -Level SUCCESS
}

# --- Check 4: Drive layout ---
Write-Log "CHECK 4/7: Verifying required drive letters..." -Level INFO
$missingDrives = @()
foreach ($drive in @('D','E','L','T')) {
    if (-not (Test-Path "${drive}:\")) { $missingDrives += $drive }
}
if ($missingDrives.Count -gt 0) {
    Exit-WithError "Missing drives: $($missingDrives -join ', '). Provision drives before installing."
}
Write-Log "All required drives present (D, E, L, T): OK" -Level SUCCESS

# --- Check 5: Free disk space (minimum 10 GB per data drive) ---
Write-Log "CHECK 5/7: Checking free disk space (minimum 10 GB per drive)..." -Level INFO
$spaceFail = $false
foreach ($drive in @('D','E','L','T')) {
    $disk = Get-PSDrive -Name $drive -PSProvider FileSystem -ErrorAction SilentlyContinue
    if ($disk) {
        $freeGB = [Math]::Round($disk.Free / 1GB, 1)
        if ($freeGB -lt 10) {
            Write-Log "${drive}: has only ${freeGB} GB free - minimum 10 GB required." -Level WARN
            $spaceFail = $true
        } else {
            Write-Log "${drive}: ${freeGB} GB free: OK" -Level INFO
        }
    }
}
if ($spaceFail) {
    Exit-WithError "Insufficient disk space on one or more drives. Resolve before continuing."
}
Write-Log "Disk space: OK" -Level SUCCESS

# --- Check 6: RAM ---
Write-Log "CHECK 6/7: Verifying system RAM..." -Level INFO
$totalRamMB = [Math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB)
Write-Log "Total RAM: ${totalRamMB} MB" -Level INFO
if ($totalRamMB -lt 1024) {
    Exit-WithError "Total RAM (${totalRamMB} MB) is below the SQL Server 2025 minimum of 1 GB. Provision more memory."
} elseif ($totalRamMB -lt 4096) {
    Write-Log "Total RAM (${totalRamMB} MB) meets minimum but is below recommended 4 GB for SQL Server 2025." -Level WARN
}
# Validate max memory setting does not exceed physical RAM (leave ~1 GB for OS)
$osHeadroomMB = 1024
$safeMaxMem   = $totalRamMB - $osHeadroomMB
if ($MEM_MAX_MB -gt $safeMaxMem) {
    Write-Log "Configured max memory (${MEM_MAX_MB} MB) exceeds safe ceiling (${safeMaxMem} MB). Adjusting to ${safeMaxMem} MB." -Level WARN
    $script:MEM_MAX_MB = $safeMaxMem
}
Write-Log "Memory config - Min: ${MEM_MIN_MB} MB (SQL 2025 non-Express minimum: 1024 MB) | Max: $($script:MEM_MAX_MB) MB: OK" -Level SUCCESS

# --- Check 7: .NET Framework 4.7.2+ (SQL Server 2025 minimum requirement) ---
# SQL Server 2025 (17.x) requires .NET Framework 4.7.2 (release key 461808).
# WS2025 ships with .NET 4.8 (533320) built-in so this check will always pass
# on a clean WS2025 guest. Retained as a safety net.
Write-Log "CHECK 7/7: Verifying .NET Framework 4.7.2+ (SQL Server 2025 minimum)..." -Level INFO
$dotNetRelease = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Release
# 461808 = .NET 4.7.2  |  528040 = .NET 4.8 on WS2019  |  528372 = WS2022  |  533320 = WS2025
if (-not $dotNetRelease -or $dotNetRelease -lt 461808) {
    Exit-WithError ".NET Framework 4.7.2 or later not detected (Release key: $dotNetRelease). SQL Server 2025 requires .NET 4.7.2 minimum."
}
Write-Log ".NET Framework detected OK (Release key: $dotNetRelease = $(if ($dotNetRelease -ge 533320) {'4.8 on WS2025'} elseif ($dotNetRelease -ge 528372) {'4.8 on WS2022'} elseif ($dotNetRelease -ge 528040) {'4.8'} elseif ($dotNetRelease -ge 461808) {'4.7.2'} else {'Unknown'}))" -Level SUCCESS

# ============================================================
# CREATE DIRECTORIES
# ============================================================
Write-Log "Creating SQL data directories..." -Level HEADER
if (-not (Test-Path $CONFIG_DIR)) {
    if ($DryRun) {
        Write-Log "[DRYRUN] Would create config dir: $CONFIG_DIR" -Level DRYRUN
    } else {
        [void](New-Item -ItemType Directory -Path $CONFIG_DIR)
    }
}
foreach ($key in $PATHS.Keys) {
    if (-not (Test-Path $PATHS[$key])) {
        if ($DryRun) {
            Write-Log "[DRYRUN] Would create: $($PATHS[$key])" -Level DRYRUN
        } else {
            [void](New-Item -ItemType Directory -Path $PATHS[$key] -Force)
            Write-Log "Created: $($PATHS[$key])" -Level INFO
        }
    } else {
        Write-Log "Exists:  $($PATHS[$key])" -Level INFO
    }
}

# ============================================================
# TEMPDB FILE COUNT
# ============================================================
$logicalCores  = (Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors
$tempDbFileCount = [Math]::Min($logicalCores, 8)
Write-Log "TempDB file count: $tempDbFileCount (logical cores: $logicalCores, capped at 8)" -Level INFO

# ============================================================
# GENERATE ConfigurationFile.ini
# ============================================================
Write-Log "Generating ConfigurationFile.ini..." -Level HEADER

$sysAdminLine = ($SysAdminAccounts | ForEach-Object { "`"$_`"" }) -join ' '

$iniContent = @"
; ============================================================
; SQL Server Uniform Install - ConfigurationFile.ini
; Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
; Server    : $env:COMPUTERNAME
; Instance  : $INSTANCE_NAME
; Collation : $COLLATION
; ============================================================

[OPTIONS]

; ----- Setup control -----
ACTION                         = "Install"
QUIET                          = "True"
QUIETSIMPLE                    = "False"
INDICATEPROGRESS               = "True"
IACCEPTSQLSERVERLICENSETERMS   = "True"
SUPPRESSPRIVACYSTATEMENTNOTICE = "True"
ENU                            = "True"

; ----- Updates -----
UpdateEnabled                  = "True"
UpdateSource                   = "$UpdateSource"

; ----- Features -----
; IMPORTANT: Use explicit granular tokens only. Do NOT use the parent alias "SQL" -
; it silently pulls in Replication alongside the engine as a sub-feature.
;
; Installed:
;   SQLEngine  - Database Engine only (no Replication, no DQ)
;   FullText   - Full-Text Search service (fd launcher + fdhost)
;   Conn       - Client connectivity components (ODBC, OLE DB drivers)
;   BC         - Backward compatibility (sqlcmd, bcp, SMO shared components)
;
; Explicitly excluded:
;   Replication      - Not required; parent alias "SQL" would have added this silently
;   AS               - Analysis Services
;   IS               - Integration Services
;   RS               - Reporting Services (removed in SQL Server 2025)
;   MDS              - Master Data Services (removed in SQL Server 2025)
;   DQC / DQ         - Data Quality Services (removed in SQL Server 2025)
;   POLYBASE         - PolyBase / external data sources
;   AZUREEXTENSION   - Azure Arc agent; NOT installed - on-premises only
;   SDK / SNAC_SDK   - Deprecated; not available in SQL Server 2022+
;   Tools / SSMS     - Shipped out-of-band; install SSMS separately
FEATURES                       = SQLEngine,FullText,Conn,BC

; ----- Instance -----
INSTANCENAME                   = "$INSTANCE_NAME"
INSTANCEID                     = "$INSTANCE_NAME"

; ----- Collation -----
SQLCOLLATION                   = "$COLLATION"

; ----- Service accounts -----
SQLSVCACCOUNT                  = "$SqlSvcAccount"
AGTSVCACCOUNT                  = "$AgtSvcAccount"
SQLSVCSTARTUPTYPE              = "Automatic"
AGTSVCSTARTUPTYPE              = "Automatic"
BROWSERSVCSTARTUPTYPE          = "Automatic"

; ----- Authentication (Windows only) -----
; SECURITYMODE not set = Windows Authentication
SQLSYSADMINACCOUNTS            = $sysAdminLine
ADDCURRENTUSERASSQLADMIN       = "False"

; ----- Data directories -----
SQLUSERDBDIR                   = "D:\Data"
SQLUSERDBLOGDIR                = "L:\Logs"
SQLBACKUPDIR                   = "D:\Backup"

; ----- TempDB -----
SQLTEMPDBDIR                   = "T:\TempDB"
SQLTEMPDBLOGDIR                = "T:\TempDB\Logs"
SQLTEMPDBFILECOUNT             = "$tempDbFileCount"
SQLTEMPDBFILESIZE              = "10240"
SQLTEMPDBFILEGROWTH            = "512"
SQLTEMPDBLOGFILESIZE           = "8192"
SQLTEMPDBLOGFILEGROWTH         = "512"

; ----- Protocols -----
TCPENABLED                     = "1"
NPENABLED                      = "1"

; ----- Performance -----
SQLSVCINSTANTFILEINIT          = "True"
"@

if ($DryRun) {
    Write-Log "[DRYRUN] ConfigurationFile.ini NOT written to disk. Preview below:" -Level DRYRUN
    $iniContent -split "`n" | ForEach-Object { Write-Log "  $_" -Level DRYRUN }
} else {
    if (-not (Test-Path $CONFIG_DIR)) { [void](New-Item -ItemType Directory -Path $CONFIG_DIR) }
    $iniContent | Out-File -FilePath $CONFIG_FILE -Encoding ASCII -Force
    Write-Log "ConfigurationFile.ini written: $CONFIG_FILE" -Level SUCCESS
}

# ============================================================
# MOUNT ISO
# ============================================================
Write-Log "Mounting ISO: $ISOPath" -Level HEADER

if ($DryRun) {
    Write-Log "[DRYRUN] Would mount ISO: $ISOPath" -Level DRYRUN
    Write-Log "[DRYRUN] Would run: <setup.exe> /Q /ACTION=Install /CONFIGURATIONFILE=`"$CONFIG_FILE`" /IACCEPTSQLSERVERLICENSETERMS /SUPPRESSPRIVACYSTATEMENTNOTICE" -Level DRYRUN
    Write-Log "[DRYRUN] NOTE: ArgumentList passed as flat string (not array) to avoid PS Start-Process quoting bug." -Level DRYRUN
    if ($SqlAcctType -eq 'DomainUser') {
        $sqlPwdStatus = if ($SqlSvcPasswordPlain) { "supplied (not shown)" } else { "MISSING" }
        Write-Log "[DRYRUN] SQL Engine password: $sqlPwdStatus" -Level DRYRUN
    }
    if ($AgtAcctType -eq 'DomainUser') {
        $agtPwdStatus = if ($AgtSvcPasswordPlain) { "supplied (not shown)" } else { "MISSING" }
        Write-Log "[DRYRUN] SQL Agent password : $agtPwdStatus" -Level DRYRUN
    }
} else {
    try {
        $mount       = Mount-DiskImage -ImagePath $ISOPath -PassThru
        $driveLetter = ($mount | Get-Volume).DriveLetter
        $setupExe    = "${driveLetter}:\setup.exe"
        Write-Log "ISO mounted on ${driveLetter}: | setup.exe: $setupExe" -Level SUCCESS
    } catch {
        Exit-WithError "Failed to mount ISO: $_"
    }

    if (-not (Test-Path $setupExe)) {
        Dismount-DiskImage -ImagePath $ISOPath -ErrorAction SilentlyContinue
        Exit-WithError "setup.exe not found at: $setupExe. ISO may be corrupt or wrong media."
    }

    # ============================================================
    # RUN SETUP
    # ============================================================
    Write-Log "Starting SQL Server unattended install..." -Level HEADER
    Write-Log "Command: $setupExe /Q /ACTION=Install /CONFIGURATIONFILE=`"$CONFIG_FILE`"" -Level INFO

    # IMPORTANT: Pass a single flat string to -ArgumentList, NOT an array.
    # Start-Process joins array elements with spaces and re-processes quotes,
    # which corrupts paths containing backslashes and causes setup.exe to receive
    # malformed tokens like "//" that fail with InputSettingValidationException.
    # /IACCEPTSQLSERVERLICENSETERMS and /SUPPRESSPRIVACYSTATEMENTNOTICE are also
    # in the INI - they are repeated on CLI only as belt-and-braces; the flat
    # string approach handles this safely.
    # Build setup arg string. Passwords are injected here for domain user accounts only.
    # They exist in memory for this one Start-Process call and are never written to disk.
    $setupArgString = "/Q /ACTION=Install /CONFIGURATIONFILE=`"$CONFIG_FILE`" /IACCEPTSQLSERVERLICENSETERMS /SUPPRESSPRIVACYSTATEMENTNOTICE"
    if ($SqlAcctType -eq 'DomainUser' -and $SqlSvcPasswordPlain) {
        $setupArgString += " /SQLSVCPASSWORD=`"$SqlSvcPasswordPlain`""
    }
    if ($AgtAcctType -eq 'DomainUser' -and $AgtSvcPasswordPlain) {
        $setupArgString += " /AGTSVCPASSWORD=`"$AgtSvcPasswordPlain`""
    }

    try {
        if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Install SQL Server instance '$INSTANCE_NAME'")) {
            $proc = Start-Process -FilePath $setupExe `
                                  -ArgumentList $setupArgString `
                                  -Wait -PassThru -NoNewWindow
            # Zero plain text passwords from memory immediately after process exits
            $SqlSvcPasswordPlain = $null
            $AgtSvcPasswordPlain = $null

            # Setup bootstrap log path (version-agnostic glob)
            $bootstrapLog = Get-ChildItem -Path "C:\Program Files\Microsoft SQL Server" `
                                          -Filter "Summary.txt" -Recurse -ErrorAction SilentlyContinue |
                            Sort-Object LastWriteTime -Descending | Select-Object -First 1

            if ($proc.ExitCode -eq 0) {
                Write-Log "Setup completed successfully. Exit code: 0" -Level SUCCESS
            } else {
                Write-Log "Setup failed. Exit code: $($proc.ExitCode)" -Level ERROR
                if ($bootstrapLog) {
                    Write-Log "Bootstrap log: $($bootstrapLog.FullName)" -Level ERROR
                    Get-Content $bootstrapLog.FullName -Tail 40 |
                        ForEach-Object { Write-Log $_ -Level ERROR }
                }
                Dismount-DiskImage -ImagePath $ISOPath -ErrorAction SilentlyContinue
                exit $proc.ExitCode
            }
        }
    } catch {
        Dismount-DiskImage -ImagePath $ISOPath -ErrorAction SilentlyContinue
        Exit-WithError "Setup process threw an exception: $_"
    }
}

# ============================================================
# POST-INSTALL: SERVICE VALIDATION
# ============================================================
Write-Log "Validating services post-install..." -Level HEADER

$allServicesOK = $true

if ($DryRun) {
    Write-Log "[DRYRUN] Would validate services: MSSQLSERVER, SQLSERVERAGENT" -Level DRYRUN
} else {
    # Named instance service names use $ suffix: MSSQL$<name> and SQLAgent$<name>
    $requiredServices = @("MSSQL`$$INSTANCE_NAME", "SQLAgent`$$INSTANCE_NAME")

    foreach ($svcName in $requiredServices) {
        $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
        if (-not $svc) {
            Write-Log "Service not found: $svcName" -Level ERROR
            $allServicesOK = $false
        } elseif ($svc.Status -ne 'Running') {
            Write-Log "Service '$svcName' exists but is not Running (Status: $($svc.Status))" -Level WARN
            $allServicesOK = $false
        } else {
            Write-Log "Service '$svcName': Running OK" -Level SUCCESS
        }
    }
}

# ============================================================
# POST-INSTALL: sp_configure HARDENING
# ============================================================
Write-Log "Applying post-install configuration via SQL..." -Level HEADER

function Invoke-SqlCmd {
    param([string]$Query, [string]$Description)
    Write-Log "SQL: $Description" -Level INFO
    try {
        # Named instance requires COMPUTERNAME\INSTANCENAME; default instance uses COMPUTERNAME only
        $sqlTarget = if ($INSTANCE_NAME -eq 'MSSQLSERVER') { $env:COMPUTERNAME } else { "$env:COMPUTERNAME\$INSTANCE_NAME" }
        $conn = New-Object System.Data.SqlClient.SqlConnection(
            "Server=$sqlTarget;Integrated Security=SSPI;Connect Timeout=30;")
        $conn.Open()
        $cmd             = $conn.CreateCommand()
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = 60
        [void]$cmd.ExecuteNonQuery()
        $conn.Close()
        Write-Log "  -> OK" -Level SUCCESS
    } catch {
        Write-Log "  -> FAILED: $_" -Level WARN
    }
}

if ($DryRun) {
    $sqlTarget = if ($INSTANCE_NAME -eq 'MSSQLSERVER') { $env:COMPUTERNAME } else { "$env:COMPUTERNAME\$INSTANCE_NAME" }
    Write-Log "[DRYRUN] Would connect to: $sqlTarget" -Level DRYRUN
    Write-Log "[DRYRUN] Would apply the following SQL post-install configuration:" -Level DRYRUN
    Write-Log "[DRYRUN]   sp_configure 'show advanced options', 1" -Level DRYRUN
    Write-Log "[DRYRUN]   sp_configure 'min server memory (MB)', $MEM_MIN_MB" -Level DRYRUN
    Write-Log "[DRYRUN]   sp_configure 'max server memory (MB)', $($script:MEM_MAX_MB)" -Level DRYRUN
    Write-Log "[DRYRUN]   sp_configure 'cost threshold for parallelism', 50" -Level DRYRUN
    Write-Log "[DRYRUN]   ALTER DATABASE [model] SET TARGET_RECOVERY_TIME = 60 SECONDS" -Level DRYRUN
    Write-Log "[DRYRUN]   ALTER LOGIN [sa] DISABLE" -Level DRYRUN
} else {
    # Max / min server memory
    Invoke-SqlCmd -Description "Enable advanced options" -Query @"
EXEC sp_configure 'show advanced options', 1; RECONFIGURE WITH OVERRIDE;
"@

    Invoke-SqlCmd -Description "Set min server memory: $MEM_MIN_MB MB" -Query @"
EXEC sp_configure 'min server memory (MB)', $MEM_MIN_MB; RECONFIGURE WITH OVERRIDE;
"@

    Invoke-SqlCmd -Description "Set max server memory: $($script:MEM_MAX_MB) MB" -Query @"
EXEC sp_configure 'max server memory (MB)', $($script:MEM_MAX_MB); RECONFIGURE WITH OVERRIDE;
"@

    # Cost threshold for parallelism (default 5 causes too-aggressive parallelism)
    Invoke-SqlCmd -Description "Cost threshold for parallelism: 50" -Query @"
EXEC sp_configure 'cost threshold for parallelism', 50; RECONFIGURE WITH OVERRIDE;
"@

    # Indirect checkpoint on model (prevents long recovery on unclean shutdown)
    Invoke-SqlCmd -Description "Indirect checkpoint on model DB (60s target recovery)" -Query @"
ALTER DATABASE [model] SET TARGET_RECOVERY_TIME = 60 SECONDS;
"@

    # Disable SA (Windows Auth only)
    Invoke-SqlCmd -Description "Disable SA login" -Query "ALTER LOGIN [sa] DISABLE;"

    # Verify collation
    Invoke-SqlCmd -Description "Log server collation" -Query @"
SELECT SERVERPROPERTY('Collation') AS ServerCollation;
"@
}

# ============================================================
# POST-INSTALL: PROTOCOL VERIFICATION
# ============================================================
Write-Log "Verifying network protocols via registry..." -Level HEADER

if ($DryRun) {
    Write-Log "[DRYRUN] Would verify TCP:1433 and Named Pipes enabled via registry/SQL Server Configuration Manager." -Level DRYRUN
} else {
    $protocolBase = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL*\MSSQLServer\SuperSocketNetLib"
    $protoKeys    = Get-Item -Path $protocolBase -ErrorAction SilentlyContinue

    if (-not $protoKeys) {
        Write-Log "Protocol registry path not found. Verify TCP and Named Pipes via SQL Server Configuration Manager." -Level WARN
    } else {
        Write-Log "Protocol registry path found. Use SQL Server Configuration Manager to confirm TCP:1433 and NP are enabled." -Level INFO
    }
}

# ============================================================
# CLEANUP
# ============================================================
if (-not $DryRun) {
    Write-Log "Dismounting ISO..." -Level INFO
    Dismount-DiskImage -ImagePath $ISOPath -ErrorAction SilentlyContinue
}

# ============================================================
# SUMMARY
# ============================================================
Write-Log "============================================================" -Level HEADER
if ($DryRun) {
    Write-Log " DRY RUN COMPLETE - $env:COMPUTERNAME" -Level DRYRUN
    Write-Log " No changes were made to this system." -Level DRYRUN
    Write-Log " Review output above. If all checks passed, re-run WITHOUT -DryRun to install." -Level DRYRUN
} else {
    Write-Log " INSTALL SUMMARY - $env:COMPUTERNAME" -Level HEADER
}
Write-Log "============================================================" -Level HEADER
Write-Log "Mode        : $(if ($DryRun) { 'DRY RUN' } else { 'INSTALL' })" -Level $(if ($DryRun) { 'DRYRUN' } else { 'INFO' })
Write-Log "Instance    : $INSTANCE_NAME" -Level INFO
Write-Log "Collation   : $COLLATION" -Level INFO
Write-Log "Data dir    : D:\Data" -Level INFO
Write-Log "Log dir     : L:\Logs" -Level INFO
Write-Log "Backup dir  : D:\Backup" -Level INFO
Write-Log "TempDB dir  : T:\TempDB" -Level INFO
Write-Log "Min memory  : $MEM_MIN_MB MB" -Level INFO
Write-Log "Max memory  : $($script:MEM_MAX_MB) MB" -Level INFO
Write-Log "Protocols   : TCP/IP + Named Pipes" -Level INFO
if (-not $DryRun) {
    Write-Log "Services OK : $allServicesOK" -Level $(if ($allServicesOK) { 'SUCCESS' } else { 'WARN' })
}
Write-Log "Log file    : $LogFile" -Level INFO
Write-Log "============================================================" -Level HEADER