[CmdletBinding()]
param(
    [ValidateRange(1, 300)]
    [int]$SampleSeconds = 120,

    [ValidateRange(1, 100)]
    [int]$Top = 20,

    [ValidateRange(1, 10)]
    [int]$TimelineIntervalSec = 1,

    [switch]$Quick,

    [switch]$PruneOldRuns,

    [ValidateRange(1, 3650)]
    [int]$RetentionDays = 2,

    [switch]$ArchiveOldRuns,

    [switch]$IncludeServices,

    [switch]$EmitJson,

    [string]$JsonPath = "",

    [string]$OutDir = "",

    [switch]$NoClear
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Section {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [ConsoleColor]$Color = [ConsoleColor]::Cyan
    )

    Write-Host "`n=== $Title ===" -ForegroundColor $Color
}

function Write-KeyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Key,

        [Parameter(Mandatory)]
        [string]$Value,

        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    Write-Host ("- {0}: {1}" -f $Key, $Value) -ForegroundColor $Color
}

function Write-Stage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-Host ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor DarkCyan
}

function New-SafeDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        $null = New-Item -ItemType Directory -Path $Path -Force
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

function Save-Json {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $dir = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        $null = New-Item -ItemType Directory -Path $dir -Force
    }

    $InputObject | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Export-IfAny {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IEnumerable]$Data,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $items = @($Data)
    if ($items.Count -gt 0) {
        $dir = Split-Path -Parent $Path
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            $null = New-Item -ItemType Directory -Path $dir -Force
        }

        $items | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    }
}

function Resolve-ComputerNameSafe {
    [CmdletBinding()]
    param()

    try {
        if (-not [string]::IsNullOrWhiteSpace($env:COMPUTERNAME)) {
            return $env:COMPUTERNAME
        }

        return [System.Environment]::MachineName
    }
    catch {
        return 'DESCONHECIDO'
    }
}

function Test-IsAdmin {
    [CmdletBinding()]
    param()

    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = [Security.Principal.WindowsPrincipal]::new($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Get-PowerPlanName {
    [CmdletBinding()]
    param()

    try {
        $raw = powercfg /getactivescheme 2>$null
        if ($raw -match '\((?<name>.+)\)') {
            return $Matches.name.Trim()
        }

        return ($raw | Select-Object -First 1).Trim()
    }
    catch {
        return 'Unknown'
    }
}

function Convert-ToNullableDouble {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    try {
        return [double]$Value
    }
    catch {
        return $null
    }
}

function Convert-BytesToHuman {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [double]$Bytes
    )

    $units = @('B', 'KB', 'MB', 'GB', 'TB')
    $value = [double]$Bytes
    $index = 0

    while ($value -ge 1024 -and $index -lt ($units.Length - 1)) {
        $value = $value / 1024
        $index++
    }

    return ('{0:N2} {1}' -f $value, $units[$index])
}

function Convert-KBToHuman {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [double]$Kilobytes
    )

    return Convert-BytesToHuman -Bytes ($Kilobytes * 1024)
}

function Convert-ToGiB {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Bytes
    )

    $value = Convert-ToNullableDouble -Value $Bytes
    if ($null -eq $value) {
        return $null
    }

    return [Math]::Round(($value / 1GB), 2)
}

function Convert-ToMiB {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Bytes
    )

    $value = Convert-ToNullableDouble -Value $Bytes
    if ($null -eq $value) {
        return $null
    }

    return [Math]::Round(($value / 1MB), 2)
}

function Get-PerfCounterValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [double]$DefaultValue = ([double]::NaN)
    )

    try {
        $counter = Get-Counter -Counter $Path -ErrorAction Stop
        if ($null -eq $counter -or $null -eq $counter.CounterSamples -or $counter.CounterSamples.Count -eq 0) {
            return $DefaultValue
        }

        return [double]$counter.CounterSamples[0].CookedValue
    }
    catch {
        return $DefaultValue
    }
}

function Get-CimSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ClassName,

        [string]$Namespace = 'root/cimv2',

        [string]$Filter = ''
    )

    try {
        $params = @{
            ClassName   = $ClassName
            Namespace   = $Namespace
            ErrorAction = 'Stop'
        }

        if (-not [string]::IsNullOrWhiteSpace($Filter)) {
            $params.Filter = $Filter
        }

        return Get-CimInstance @params
    }
    catch {
        return $null
    }
}

function Convert-ToDateTimeSafe {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    try {
        if ($Value -is [datetime]) {
            return [datetime]$Value
        }

        if ($Value -is [datetimeoffset]) {
            return ([datetimeoffset]$Value).DateTime
        }

        $stringValue = [string]$Value
        if ([string]::IsNullOrWhiteSpace($stringValue)) {
            return $null
        }

        try {
            return [Management.ManagementDateTimeConverter]::ToDateTime($stringValue)
        }
        catch {}

        try {
            return [datetime]::Parse($stringValue, [System.Globalization.CultureInfo]::InvariantCulture)
        }
        catch {}

        try {
            return [datetime]::Parse($stringValue)
        }
        catch {}

        return [datetime]$Value
    }
    catch {
        return $null
    }
}

function Get-StorageMediaSnapshot {
    [CmdletBinding()]
    param()

    $items = @()

    try {
        if (Get-Command -Name Get-PhysicalDisk -ErrorAction SilentlyContinue) {
            $items = Get-PhysicalDisk | Select-Object `
                FriendlyName,
            MediaType,
            HealthStatus,
            OperationalStatus,
            @{Name = 'SizeGB'; Expression = { Convert-ToGiB $_.Size } }
        }
    }
    catch {
        $items = @()
    }

    if (@($items).Count -eq 0) {
        try {
            $items = Get-CimInstance -ClassName Win32_DiskDrive | Select-Object `
                Model,
            MediaType,
            Status,
            InterfaceType,
            @{Name = 'SizeGB'; Expression = { Convert-ToGiB $_.Size } }
        }
        catch {
            $items = @()
        }
    }

    return @($items)
}

function Get-ProcessSnapshot {
    [CmdletBinding()]
    param()

    $snapshot = @{}
    $processes = Get-Process -ErrorAction SilentlyContinue

    foreach ($proc in $processes) {
        try {
            $snapshot[$proc.Id] = [PSCustomObject]@{
                Id                  = $proc.Id
                ProcessName         = $proc.ProcessName
                CPU                 = $(if ($null -ne $proc.CPU) { [double]$proc.CPU } else { 0.0 })
                WorkingSet64        = [double]$proc.WorkingSet64
                PagedMemorySize64   = [double]$proc.PagedMemorySize64
                PrivateMemorySize64 = [double]$proc.PrivateMemorySize64
                Threads             = $proc.Threads.Count
                Handles             = $proc.Handles
                StartTime           = $null
                Path                = $null
                Company             = $null
            }

            try { $snapshot[$proc.Id].StartTime = $proc.StartTime } catch {}
            try { $snapshot[$proc.Id].Path = $proc.Path } catch {}
            try { $snapshot[$proc.Id].Company = $proc.Company } catch {}
        }
        catch {
            continue
        }
    }

    return $snapshot
}

function Get-ProcessCpuDelta {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Before,

        [Parameter(Mandatory)]
        [hashtable]$After,

        [Parameter(Mandatory)]
        [double]$IntervalSeconds,

        [Parameter(Mandatory)]
        [int]$LogicalCpuCount,

        [int]$TopCount = 15
    )

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($entry in $After.GetEnumerator()) {
        $procId = [int]$entry.Key
        $afterProc = $entry.Value

        if (-not $Before.ContainsKey($procId)) {
            continue
        }

        $beforeProc = $Before[$procId]
        $deltaCpuSeconds = [double]($afterProc.CPU - $beforeProc.CPU)
        if ($deltaCpuSeconds -lt 0) {
            continue
        }

        $cpuPercent = 0.0
        if ($IntervalSeconds -gt 0 -and $LogicalCpuCount -gt 0) {
            $cpuPercent = ($deltaCpuSeconds / ($IntervalSeconds * $LogicalCpuCount)) * 100.0
        }

        $results.Add([PSCustomObject]@{
                PID             = $procId
                ProcessName     = $afterProc.ProcessName
                CPUPercent      = [Math]::Round($cpuPercent, 2)
                CpuSecondsDelta = [Math]::Round($deltaCpuSeconds, 4)
                Threads         = $afterProc.Threads
                Handles         = $afterProc.Handles
                WorkingSetMB    = [Math]::Round(($afterProc.WorkingSet64 / 1MB), 2)
                PrivateMemoryMB = [Math]::Round(($afterProc.PrivateMemorySize64 / 1MB), 2)
                StartTime       = $afterProc.StartTime
                Path            = $afterProc.Path
                Company         = $afterProc.Company
            })
    }

    return $results |
    Sort-Object -Property @{ Expression = 'CPUPercent'; Descending = $true }, @{ Expression = 'WorkingSetMB'; Descending = $true } |
    Select-Object -First $TopCount
}

function Get-SystemSummary {
    [CmdletBinding()]
    param()

    $os = Get-CimSafe -ClassName 'Win32_OperatingSystem'
    $computer = Get-CimSafe -ClassName 'Win32_ComputerSystem'
    $cpu = Get-CimSafe -ClassName 'Win32_Processor'
    $logicalDisks = Get-CimSafe -ClassName 'Win32_LogicalDisk' -Filter "DriveType = 3"

    $memoryAvailableMB = Get-PerfCounterValue -Path '\Memory\Available MBytes' -DefaultValue 0
    $memoryCommittedPct = Get-PerfCounterValue -Path '\Memory\% Committed Bytes In Use' -DefaultValue 0
    $diskBusyPct = Get-PerfCounterValue -Path '\PhysicalDisk(_Total)\% Disk Time' -DefaultValue ([double]::NaN)
    $diskQueue = Get-PerfCounterValue -Path '\PhysicalDisk(_Total)\Avg. Disk Queue Length' -DefaultValue ([double]::NaN)
    $processorQueue = Get-PerfCounterValue -Path '\System\Processor Queue Length' -DefaultValue ([double]::NaN)
    $interruptPct = Get-PerfCounterValue -Path '\Processor(_Total)\% Interrupt Time' -DefaultValue ([double]::NaN)
    $dpcPct = Get-PerfCounterValue -Path '\Processor(_Total)\% DPC Time' -DefaultValue ([double]::NaN)
    $contextSwitchesPerSec = Get-PerfCounterValue -Path '\System\Context Switches/sec' -DefaultValue ([double]::NaN)
    $pagesPerSec = Get-PerfCounterValue -Path '\Memory\Pages/sec' -DefaultValue ([double]::NaN)

    $totalVisibleMemoryKB = if ($os) { [double]$os.TotalVisibleMemorySize } else { 0 }
    $freePhysicalMemoryKB = if ($os) { [double]$os.FreePhysicalMemory } else { 0 }
    $usedPhysicalMemoryKB = [Math]::Max(0, $totalVisibleMemoryKB - $freePhysicalMemoryKB)
    $memoryUsedPct = if ($totalVisibleMemoryKB -gt 0) { [Math]::Round(($usedPhysicalMemoryKB / $totalVisibleMemoryKB) * 100, 2) } else { 0 }

    $uptime = $null
    $lastBootDate = $null
    if ($os -and $os.LastBootUpTime) {
        $boot = Convert-ToDateTimeSafe -Value $os.LastBootUpTime
        if ($boot) {
            $lastBootDate = $boot
            $uptime = (Get-Date) - $boot
        }
    }

    $diskSummary = @()
    if ($logicalDisks) {
        foreach ($disk in @($logicalDisks)) {
            $freeGB = if ($disk.FreeSpace -ne $null) { [Math]::Round(([double]$disk.FreeSpace / 1GB), 2) } else { 0 }
            $sizeGB = if ($disk.Size -ne $null -and [double]$disk.Size -gt 0) { [Math]::Round(([double]$disk.Size / 1GB), 2) } else { 0 }
            $usedPct = if ($sizeGB -gt 0) { [Math]::Round((1 - ($freeGB / $sizeGB)) * 100, 2) } else { 0 }

            $diskSummary += [PSCustomObject]@{
                Drive      = $disk.DeviceID
                FileSystem = $disk.FileSystem
                SizeGB     = $sizeGB
                FreeGB     = $freeGB
                UsedPct    = $usedPct
                VolumeName = $disk.VolumeName
            }
        }
    }

    return [PSCustomObject]@{
        Timestamp             = (Get-Date).ToString('s')
        ComputerName          = Resolve-ComputerNameSafe
        UserName              = $env:USERNAME
        IsAdmin               = Test-IsAdmin
        PowerPlan             = Get-PowerPlanName
        OSCaption             = $(if ($os) { $os.Caption } else { 'N/D' })
        OSVersion             = $(if ($os) { $os.Version } else { 'N/D' })
        BuildNumber           = $(if ($os) { $os.BuildNumber } else { 'N/D' })
        LastBoot              = $(if ($lastBootDate) { $lastBootDate.ToString('s') } else { $null })
        Manufacturer          = $(if ($computer) { $computer.Manufacturer } else { 'N/D' })
        Model                 = $(if ($computer) { $computer.Model } else { 'N/D' })
        LogicalCpuCount       = [Environment]::ProcessorCount
        CPUName               = $(if ($cpu) { ($cpu | Select-Object -First 1 -ExpandProperty Name) } else { 'N/D' })
        CpuCores              = $(if ($cpu) { [int](($cpu | Select-Object -First 1).NumberOfCores) } else { $null })
        TotalRAM              = Convert-KBToHuman -Kilobytes $totalVisibleMemoryKB
        UsedRAM               = Convert-KBToHuman -Kilobytes $usedPhysicalMemoryKB
        FreeRAM               = Convert-KBToHuman -Kilobytes $freePhysicalMemoryKB
        TotalPhysicalMemoryGB = Convert-ToGiB (($computer | Select-Object -First 1).TotalPhysicalMemory)
        MemoryUsedPct         = $memoryUsedPct
        MemoryAvailableMB     = [Math]::Round($memoryAvailableMB, 2)
        MemoryCommittedPct    = [Math]::Round($memoryCommittedPct, 2)
        DiskBusyPct           = $(if ([double]::IsNaN($diskBusyPct)) { $null } else { [Math]::Round($diskBusyPct, 2) })
        DiskQueueLength       = $(if ([double]::IsNaN($diskQueue)) { $null } else { [Math]::Round($diskQueue, 2) })
        ProcessorQueueLength  = $(if ([double]::IsNaN($processorQueue)) { $null } else { [Math]::Round($processorQueue, 2) })
        InterruptPct          = $(if ([double]::IsNaN($interruptPct)) { $null } else { [Math]::Round($interruptPct, 2) })
        DpcPct                = $(if ([double]::IsNaN($dpcPct)) { $null } else { [Math]::Round($dpcPct, 2) })
        ContextSwitchesPerSec = $(if ([double]::IsNaN($contextSwitchesPerSec)) { $null } else { [Math]::Round($contextSwitchesPerSec, 2) })
        PagesPerSec           = $(if ([double]::IsNaN($pagesPerSec)) { $null } else { [Math]::Round($pagesPerSec, 2) })
        Uptime                = $(if ($uptime) { '{0:%d}d {0:hh}h {0:mm}m {0:ss}s' -f $uptime } else { 'N/D' })
        Disks                 = $diskSummary
        PageFiles             = @(
            Get-CimSafe -ClassName 'Win32_PageFileUsage' | ForEach-Object {
                [PSCustomObject]@{
                    Name            = $_.Name
                    AllocatedBaseMB = $_.AllocatedBaseSize
                    CurrentUsageMB  = $_.CurrentUsage
                    PeakUsageMB     = $_.PeakUsage
                }
            }
        )
        StorageMedia          = @(Get-StorageMediaSnapshot)
    }
}

function Get-SystemSnapshotForExport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$SystemSummary
    )

    return [PSCustomObject]@{
        Timestamp             = $SystemSummary.Timestamp
        ComputerName          = $SystemSummary.ComputerName
        UserName              = $SystemSummary.UserName
        IsAdmin               = $SystemSummary.IsAdmin
        PowerPlan             = $SystemSummary.PowerPlan
        OS                    = $SystemSummary.OSCaption
        OSVersion             = $SystemSummary.OSVersion
        BuildNumber           = $SystemSummary.BuildNumber
        LastBoot              = $SystemSummary.LastBoot
        Manufacturer          = $SystemSummary.Manufacturer
        Model                 = $SystemSummary.Model
        TotalPhysicalMemoryGB = $SystemSummary.TotalPhysicalMemoryGB
        CPUName               = $SystemSummary.CPUName
        CPULogicalProcessors  = $SystemSummary.LogicalCpuCount
        CPUCores              = $SystemSummary.CpuCores
        Drives                = @($SystemSummary.Disks | ForEach-Object {
                [PSCustomObject]@{
                    Drive      = $_.Drive
                    Label      = $_.VolumeName
                    FileSystem = $_.FileSystem
                    SizeGB     = $_.SizeGB
                    FreeGB     = $_.FreeGB
                    FreePct    = if ($_.SizeGB -gt 0) { [Math]::Round((($_.FreeGB / $_.SizeGB) * 100), 2) } else { $null }
                }
            })
        PageFiles             = @($SystemSummary.PageFiles)
        StorageMedia          = @($SystemSummary.StorageMedia)
    }
}

function Get-BrowserProcesses {
    [CmdletBinding()]
    param()

    $browserNames = @('msedge', 'chrome', 'firefox', 'brave', 'opera', 'opera_gx')
    $items = New-Object System.Collections.Generic.List[object]

    foreach ($name in $browserNames) {
        $procs = Get-Process -Name $name -ErrorAction SilentlyContinue
        foreach ($proc in $procs) {
            try {
                $items.Add([PSCustomObject]@{
                        Timestamp       = (Get-Date).ToString('s')
                        Browser         = $name
                        PID             = $proc.Id
                        ProcessName     = $proc.ProcessName
                        CPUSecondsTotal = $(if ($null -ne $proc.CPU) { [Math]::Round([double]$proc.CPU, 2) } else { 0 })
                        WorkingSetMB    = [Math]::Round(($proc.WorkingSet64 / 1MB), 2)
                        PrivateMemoryMB = [Math]::Round(($proc.PrivateMemorySize64 / 1MB), 2)
                        Threads         = $proc.Threads.Count
                        Handles         = $proc.Handles
                        MainWindowTitle = $proc.MainWindowTitle
                        StartTime       = $(try { $proc.StartTime } catch { $null })
                        Path            = $(try { $proc.Path } catch { $null })
                    })
            }
            catch {
                continue
            }
        }
    }

    return $items | Sort-Object -Property @{ Expression = 'WorkingSetMB'; Descending = $true }, @{ Expression = 'CPUSecondsTotal'; Descending = $true }
}

function Get-ServiceMapping {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int[]]$ProcessIds
    )

    $map = @{}
    if (@($ProcessIds).Count -eq 0) {
        return $map
    }

    try {
        $services = Get-CimInstance -ClassName Win32_Service -ErrorAction Stop | Where-Object { $_.ProcessId -in $ProcessIds }
        foreach ($svc in $services) {
            if (-not $map.ContainsKey([int]$svc.ProcessId)) {
                $map[[int]$svc.ProcessId] = New-Object System.Collections.Generic.List[string]
            }

            $map[[int]$svc.ProcessId].Add($svc.Name)
        }
    }
    catch {
        return $map
    }

    return $map
}

function Get-PerfSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$LogicalCpuCount,

        [Parameter(Mandatory)]
        [int]$TopN,

        [hashtable]$PreviousProcessMap,

        [double]$ElapsedSec = 0
    )

    $timestamp = Get-Date

    $cpu = $null
    $sys = $null
    $mem = $null
    $page = $null
    $disk = $null
    $liveProcesses = @()

    try { $cpu = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor -Filter "Name = '_Total'" | Select-Object -First 1 } catch {}
    try { $sys = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_System | Select-Object -First 1 } catch {}
    try { $mem = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Memory | Select-Object -First 1 } catch {}
    try { $page = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_PagingFile -Filter "Name = '_Total'" -ErrorAction Stop | Select-Object -First 1 } catch {}
    try { $disk = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name = '_Total'" -ErrorAction Stop | Select-Object -First 1 } catch {}
    try { $liveProcesses = @(Get-Process -ErrorAction Stop) } catch { $liveProcesses = @() }

    $systemRow = [PSCustomObject]@{
        Timestamp            = $timestamp.ToString('s')
        CpuTotalPct          = $( $v = Convert-ToNullableDouble $cpu.PercentProcessorTime; if ($null -ne $v) { [Math]::Round($v, 2) } else { $null } )
        ProcessorQueueLength = $( $v = Convert-ToNullableDouble $sys.ProcessorQueueLength; if ($null -ne $v) { [Math]::Round($v, 2) } else { $null } )
        AvailableMB          = $( $v = Convert-ToNullableDouble $mem.AvailableMBytes; if ($null -ne $v) { [Math]::Round($v, 2) } else { $null } )
        PagesInputPerSec     = $( $v = Convert-ToNullableDouble $mem.PagesInputPersec; if ($null -ne $v) { [Math]::Round($v, 2) } else { $null } )
        PagesPerSec          = $( $v = Convert-ToNullableDouble $mem.PagesPersec; if ($null -ne $v) { [Math]::Round($v, 2) } else { $null } )
        PageFileUsagePct     = $( $v = Convert-ToNullableDouble $page.PercentUsage; if ($null -ne $v) { [Math]::Round($v, 2) } else { $null } )
        DiskBusyPct          = $( $v = Convert-ToNullableDouble $disk.PercentDiskTime; if ($null -ne $v) { [Math]::Round($v, 2) } else { $null } )
        DiskQueueLength      = $( $v = Convert-ToNullableDouble $disk.AvgDiskQueueLength; if ($null -ne $v) { [Math]::Round($v, 2) } else { $null } )
        DiskSecRead          = $( $v = Convert-ToNullableDouble $disk.AvgDisksecPerRead; if ($null -ne $v) { [Math]::Round($v, 4) } else { $null } )
        DiskSecWrite         = $( $v = Convert-ToNullableDouble $disk.AvgDisksecPerWrite; if ($null -ne $v) { [Math]::Round($v, 4) } else { $null } )
        DiskTransfersPerSec  = $( $v = Convert-ToNullableDouble $disk.DiskTransfersPersec; if ($null -ne $v) { [Math]::Round($v, 2) } else { $null } )
    }

    $currentProcessMap = @{}
    foreach ($proc in @($liveProcesses)) {
        $name = $null
        $procId = $null
        $cpuSec = $null
        $privateBytes = $null

        try { $name = $proc.ProcessName } catch {}
        try { $procId = [int]$proc.Id } catch {}
        try { $cpuSec = Convert-ToNullableDouble $proc.CPU } catch {}
        try { $privateBytes = Convert-ToNullableDouble $proc.PrivateMemorySize64 } catch {}

        if ([string]::IsNullOrWhiteSpace($name) -or $null -eq $procId) {
            continue
        }

        $currentProcessMap[$procId] = [PSCustomObject]@{
            Timestamp       = $timestamp.ToString('s')
            ProcessId       = $procId
            ProcessName     = $name.ToLowerInvariant()
            CpuSeconds      = $cpuSec
            PrivateMemoryMB = Convert-ToMiB $privateBytes
        }
    }

    $processRows = @()
    if ($PreviousProcessMap -and $ElapsedSec -gt 0) {
        $rawRows = foreach ($entry in $currentProcessMap.GetEnumerator()) {
            $procId = $entry.Key
            $cur = $entry.Value
            $prev = $PreviousProcessMap[$procId]

            if ($null -eq $prev) { continue }
            if ($cur.ProcessName -ne $prev.ProcessName) { continue }

            $cpuPct = $null
            if ($null -ne $cur.CpuSeconds -and $null -ne $prev.CpuSeconds) {
                $deltaCpu = $cur.CpuSeconds - $prev.CpuSeconds
                if ($deltaCpu -lt 0) { $deltaCpu = 0 }
                $cpuPct = [Math]::Round((($deltaCpu / [Math]::Max($ElapsedSec, 0.1)) / [Math]::Max($LogicalCpuCount, 1)) * 100, 2)
            }

            [PSCustomObject]@{
                Timestamp           = $timestamp.ToString('s')
                ProcessName         = $cur.ProcessName
                CpuPct              = $cpuPct
                WorkingSetPrivateMB = $cur.PrivateMemoryMB
                InstanceCount       = 1
            }
        }

        $processRows = @(
            $rawRows |
            Group-Object -Property ProcessName |
            ForEach-Object {
                $cpuSum = ($_.Group | Measure-Object -Property CpuPct -Sum).Sum
                $wsSum = ($_.Group | Measure-Object -Property WorkingSetPrivateMB -Sum).Sum

                [PSCustomObject]@{
                    Timestamp           = $timestamp.ToString('s')
                    ProcessName         = $_.Name
                    CpuPct              = if ($null -ne $cpuSum) { [Math]::Round([double]$cpuSum, 2) } else { $null }
                    WorkingSetPrivateMB = if ($null -ne $wsSum) { [Math]::Round([double]$wsSum, 2) } else { $null }
                    InstanceCount       = $_.Count
                }
            } |
            Sort-Object @{ Expression = { if ($null -eq $_.CpuPct) { -1 } else { $_.CpuPct } } } -Descending |
            Select-Object -First ([Math]::Max($TopN, 1))
        )
    }

    return [PSCustomObject]@{
        SystemRow         = $systemRow
        ProcessRows       = $processRows
        CurrentProcessMap = $currentProcessMap
    }
}

function Invoke-TimelineSampling {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$DurationSec,

        [Parameter(Mandatory)]
        [int]$IntervalSec,

        [Parameter(Mandatory)]
        [int]$LogicalCpuCount,

        [Parameter(Mandatory)]
        [int]$TopN
    )

    $iterations = [int][Math]::Ceiling(($DurationSec * 1.0) / [Math]::Max($IntervalSec, 1))
    $systemSamples = @()
    $processSamples = @()
    $previousProcessMap = $null
    $lastTimestamp = $null

    for ($i = 1; $i -le $iterations; $i++) {
        if ($i -gt 1) {
            Start-Sleep -Seconds $IntervalSec
        }

        $now = Get-Date
        $elapsedSec = 0
        if ($lastTimestamp) {
            $elapsedSec = [Math]::Max((New-TimeSpan -Start $lastTimestamp -End $now).TotalSeconds, 0.1)
        }

        $perf = Get-PerfSnapshot -LogicalCpuCount $LogicalCpuCount -TopN $TopN -PreviousProcessMap $previousProcessMap -ElapsedSec $elapsedSec
        $systemSamples += @($perf.SystemRow)
        if (@($perf.ProcessRows).Count -gt 0) {
            $processSamples += @($perf.ProcessRows)
        }

        $previousProcessMap = $perf.CurrentProcessMap
        $lastTimestamp = $now
    }

    return [PSCustomObject]@{
        SystemSamples  = @($systemSamples)
        ProcessSamples = @($processSamples)
    }
}

function Get-HealthFindings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$SystemSummary,

        [Parameter(Mandatory)]
        [object[]]$TopProcesses,

        [Parameter(Mandatory)]
        [object[]]$BrowserProcesses
    )

    $findings = New-Object System.Collections.Generic.List[object]

    if ($SystemSummary.MemoryUsedPct -ge 90) {
        $findings.Add([PSCustomObject]@{ Severity = 'CRITICO'; Message = 'RAM praticamente estrangulada. Quando o teclado vira gelatina, essa porra costuma estar no topo da lista.' })
    }
    elseif ($SystemSummary.MemoryUsedPct -ge 80) {
        $findings.Add([PSCustomObject]@{ Severity = 'ALTO'; Message = 'Uso de RAM elevado. Navegador com múltiplos subprocessos pode transformar digitação em tortura medieval.' })
    }

    if ($null -ne $SystemSummary.DiskBusyPct -and $SystemSummary.DiskBusyPct -ge 90) {
        $findings.Add([PSCustomObject]@{ Severity = 'CRITICO'; Message = 'Disco extremamente ocupado. Em HDD velho isso assassina responsividade sem pedir licença.' })
    }
    elseif ($null -ne $SystemSummary.DiskBusyPct -and $SystemSummary.DiskBusyPct -ge 75) {
        $findings.Add([PSCustomObject]@{ Severity = 'ALTO'; Message = 'Disco sob carga pesada. Pode haver paginação, indexação ou antivírus triturando I/O.' })
    }

    if ($null -ne $SystemSummary.DpcPct -and $SystemSummary.DpcPct -ge 15) {
        $findings.Add([PSCustomObject]@{ Severity = 'ALTO'; Message = 'DPC Time alto. Cheiro clássico de driver, áudio, rede, USB ou GPU fazendo cosplay de demônio.' })
    }

    if ($null -ne $SystemSummary.InterruptPct -and $SystemSummary.InterruptPct -ge 10) {
        $findings.Add([PSCustomObject]@{ Severity = 'ALTO'; Message = 'Interrupt Time alto. Pode ser hardware, driver ou dispositivo brigando com a paz mundial.' })
    }

    $topCpu = @($TopProcesses | Select-Object -First 1)
    if ($topCpu.Count -gt 0 -and $topCpu[0].CPUPercent -ge 60) {
        $findings.Add([PSCustomObject]@{ Severity = 'CRITICO'; Message = "Processo dominante em CPU: $($topCpu[0].ProcessName) (PID $($topCpu[0].PID)) usando $($topCpu[0].CPUPercent)% da CPU média amostrada." })
    }
    elseif ($topCpu.Count -gt 0 -and $topCpu[0].CPUPercent -ge 30) {
        $findings.Add([PSCustomObject]@{ Severity = 'ALTO'; Message = "Processo líder em CPU: $($topCpu[0].ProcessName) (PID $($topCpu[0].PID)) com $($topCpu[0].CPUPercent)% da CPU média amostrada." })
    }

    $browserAggregate = @(
        $BrowserProcesses | Group-Object -Property Browser | ForEach-Object {
            [PSCustomObject]@{
                Browser           = $_.Name
                ProcessCount      = $_.Count
                TotalWorkingSetMB = [Math]::Round((($_.Group | Measure-Object -Property WorkingSetMB -Sum).Sum), 2)
                TotalCpuSeconds   = [Math]::Round((($_.Group | Measure-Object -Property CPUSecondsTotal -Sum).Sum), 2)
            }
        } | Sort-Object -Property TotalWorkingSetMB -Descending
    )

    foreach ($browser in $browserAggregate) {
        if ($browser.ProcessCount -ge 10 -and $browser.TotalWorkingSetMB -ge 1500) {
            $findings.Add([PSCustomObject]@{ Severity = 'ALTO'; Message = "Navegador $($browser.Browser) está com $($browser.ProcessCount) subprocessos e ~$($browser.TotalWorkingSetMB) MB em RAM. Tá montando um datacenter particular." })
        }
        elseif ($browser.ProcessCount -ge 6 -and $browser.TotalWorkingSetMB -ge 800) {
            $findings.Add([PSCustomObject]@{ Severity = 'MEDIO'; Message = "Navegador $($browser.Browser) já acumulou $($browser.ProcessCount) subprocessos e ~$($browser.TotalWorkingSetMB) MB em RAM." })
        }
    }

    if ($findings.Count -eq 0) {
        $findings.Add([PSCustomObject]@{ Severity = 'OK'; Message = 'Nada berrando de forma grotesca nessa amostra. Se continua travando, a treta pode ser intermitente ou ligada ao navegador/perfil/extensões.' })
    }

    return $findings
}

function Show-SystemSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Summary
    )

    Write-Section -Title 'VISÃO GERAL DO SISTEMA' -Color Cyan
    Write-KeyValue -Key 'Máquina' -Value $Summary.ComputerName
    Write-KeyValue -Key 'Sistema' -Value ("{0} (build {1}, versão {2})" -f $Summary.OSCaption, $Summary.BuildNumber, $Summary.OSVersion)
    Write-KeyValue -Key 'Fabricante/Modelo' -Value ("{0} / {1}" -f $Summary.Manufacturer, $Summary.Model)
    Write-KeyValue -Key 'CPU' -Value ("{0} | lógicas: {1}" -f $Summary.CPUName, $Summary.LogicalCpuCount)
    Write-KeyValue -Key 'RAM total' -Value $Summary.TotalRAM
    Write-KeyValue -Key 'RAM usada' -Value ("{0} ({1}%)" -f $Summary.UsedRAM, $Summary.MemoryUsedPct)
    Write-KeyValue -Key 'RAM livre' -Value ("{0} | Disponível: {1} MB | Commit: {2}%" -f $Summary.FreeRAM, $Summary.MemoryAvailableMB, $Summary.MemoryCommittedPct)

    $diskBusyText = if ($null -ne $Summary.DiskBusyPct) { "$($Summary.DiskBusyPct)%" } else { 'N/D' }
    $diskQueueText = if ($null -ne $Summary.DiskQueueLength) { "$($Summary.DiskQueueLength)" } else { 'N/D' }
    $processorQueueText = if ($null -ne $Summary.ProcessorQueueLength) { "$($Summary.ProcessorQueueLength)" } else { 'N/D' }
    $interruptText = if ($null -ne $Summary.InterruptPct) { "$($Summary.InterruptPct)%" } else { 'N/D' }
    $dpcText = if ($null -ne $Summary.DpcPct) { "$($Summary.DpcPct)%" } else { 'N/D' }
    $pagesText = if ($null -ne $Summary.PagesPerSec) { "$($Summary.PagesPerSec)/s" } else { 'N/D' }

    Write-KeyValue -Key 'Disco ativo' -Value ("{0} | Queue: {1}" -f $diskBusyText, $diskQueueText)
    Write-KeyValue -Key 'Fila de processador' -Value $processorQueueText
    Write-KeyValue -Key 'Interrupt/DPC' -Value ("Interrupts: {0} | DPC: {1}" -f $interruptText, $dpcText)
    Write-KeyValue -Key 'Pages/sec' -Value $pagesText
    Write-KeyValue -Key 'Context Switches/sec' -Value $(if ($null -ne $Summary.ContextSwitchesPerSec) { "$($Summary.ContextSwitchesPerSec)" } else { 'N/D' })
    Write-KeyValue -Key 'Uptime' -Value $Summary.Uptime

    if (@($Summary.Disks).Count -gt 0) {
        Write-Section -Title 'UNIDADES' -Color DarkCyan
        $Summary.Disks |
        Sort-Object -Property Drive |
        Format-Table Drive, FileSystem, SizeGB, FreeGB, UsedPct, VolumeName -AutoSize
    }
}

function Show-TopProcesses {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Processes,

        [hashtable]$ServiceMap = @{},

        [switch]$IncludeServices
    )

    Write-Section -Title 'TOP PROCESSOS POR CPU (DELTA REAL)' -Color Yellow

    $rows = foreach ($proc in @($Processes)) {
        $serviceNames = ''
        if ($IncludeServices -and $ServiceMap.ContainsKey([int]$proc.PID)) {
            $serviceNames = ($ServiceMap[[int]$proc.PID] | Sort-Object | Select-Object -Unique) -join ', '
        }

        [PSCustomObject]@{
            PID         = $proc.PID
            Name        = $proc.ProcessName
            CPU_Pct     = $proc.CPUPercent
            CPU_Delta_s = $proc.CpuSecondsDelta
            RAM_MB      = $proc.WorkingSetMB
            Private_MB  = $proc.PrivateMemoryMB
            Threads     = $proc.Threads
            Handles     = $proc.Handles
            Services    = $serviceNames
        }
    }

    if ($IncludeServices) {
        $rows | Format-Table PID, Name, CPU_Pct, CPU_Delta_s, RAM_MB, Private_MB, Threads, Handles, Services -AutoSize
    }
    else {
        $rows | Format-Table PID, Name, CPU_Pct, CPU_Delta_s, RAM_MB, Private_MB, Threads, Handles -AutoSize
    }
}

function Show-BrowserDiagnostics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$BrowserProcesses
    )

    Write-Section -Title 'DIAGNÓSTICO DE NAVEGADORES' -Color Magenta

    $browserRows = @($BrowserProcesses)
    if ($browserRows.Count -eq 0) {
        Write-Host 'Nenhum processo de navegador alvo encontrado. Milagre estatístico ou navegador fechado.' -ForegroundColor DarkGray
        return
    }

    $aggregate = $browserRows |
    Group-Object -Property Browser |
    ForEach-Object {
        $group = $_.Group
        [PSCustomObject]@{
            Browser         = $_.Name
            ProcessCount    = $_.Count
            TotalRAM_MB     = [Math]::Round((($group | Measure-Object -Property WorkingSetMB -Sum).Sum), 2)
            TotalPrivate_MB = [Math]::Round((($group | Measure-Object -Property PrivateMemoryMB -Sum).Sum), 2)
            TotalCPU_s      = [Math]::Round((($group | Measure-Object -Property CPUSecondsTotal -Sum).Sum), 2)
        }
    } |
    Sort-Object -Property TotalRAM_MB -Descending

    Write-Host '[Resumo agregado]' -ForegroundColor DarkMagenta
    $aggregate | Format-Table Browser, ProcessCount, TotalRAM_MB, TotalPrivate_MB, TotalCPU_s -AutoSize

    Write-Host "`n[Subprocessos mais pesados]" -ForegroundColor DarkMagenta
    $browserRows |
    Select-Object -First 12 |
    Format-Table Browser, PID, WorkingSetMB, PrivateMemoryMB, CPUSecondsTotal, Threads, Handles, MainWindowTitle -AutoSize
}

function Show-Findings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Findings
    )

    Write-Section -Title 'ACHADOS E SUSPEITAS' -Color Red

    foreach ($finding in @($Findings)) {
        $color = switch ($finding.Severity) {
            'CRITICO' { [ConsoleColor]::Red }
            'ALTO' { [ConsoleColor]::Yellow }
            'MEDIO' { [ConsoleColor]::DarkYellow }
            'OK' { [ConsoleColor]::Green }
            default { [ConsoleColor]::Gray }
        }

        Write-Host ("[{0}] {1}" -f $finding.Severity, $finding.Message) -ForegroundColor $color
    }
}

function Show-NextSteps {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$SystemSummary,

        [Parameter(Mandatory)]
        [object[]]$TopProcesses,

        [Parameter(Mandatory)]
        [object[]]$BrowserProcesses,

        [string]$RunOutputDir = ''
    )

    Write-Section -Title 'PRÓXIMOS PASSOS INTELIGENTES' -Color Green

    $top = @($TopProcesses | Select-Object -First 5)
    $browserNames = @($BrowserProcesses | Select-Object -ExpandProperty Browser -Unique)

    if ($browserNames.Count -gt 0) {
        Write-Host '- Testar InPrivate/Anônimo com todas extensões desativadas. Se melhorar, o culpado mora no perfil/extensões. Parabéns ao suspeito.' -ForegroundColor Green
    }

    if ($SystemSummary.MemoryUsedPct -ge 80) {
        Write-Host '- RAM alta: feche navegador pesado, Discord, launchers, Electron inútil e qualquer tranqueira residente.' -ForegroundColor Green
    }

    if ($null -ne $SystemSummary.DiskBusyPct -and $SystemSummary.DiskBusyPct -ge 75) {
        Write-Host '- Disco alto: confira Windows Search, Defender, OneDrive, update e paginação. HDD velho sofre mais que protagonista de novela.' -ForegroundColor Green
    }

    if ($null -ne $SystemSummary.DpcPct -and $SystemSummary.DpcPct -ge 15) {
        Write-Host '- DPC/Interrupt alto: atualizar/reverter drivers de vídeo, áudio, rede e chipset. Desconecte USB suspeito e veja se a histeria cai.' -ForegroundColor Green
    }

    if ($top.Count -gt 0) {
        $names = ($top | ForEach-Object { "$($_.ProcessName)[$($_.PID)]" }) -join ', '
        Write-Host ("- Investigue estes processos primeiro: {0}" -f $names) -ForegroundColor Green
    }

    if (-not [string]::IsNullOrWhiteSpace($RunOutputDir)) {
        Write-Host ("- Logs desta execução salvos em: {0}" -f $RunOutputDir) -ForegroundColor Green
    }

    Write-Host '- Rode este script duas ou três vezes: em idle, com o navegador aberto limpo, e no momento exato da engasgada. A comparação entrega o vagabundo.' -ForegroundColor Green
}

function Get-RunRetentionCutoff {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$Days
    )

    return (Get-Date).AddDays(-1 * [Math]::Abs($Days))
}

function Get-RunDirectories {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BaseDirectory
    )

    if (-not (Test-Path -LiteralPath $BaseDirectory)) {
        return @()
    }

    return @(
        Get-ChildItem -LiteralPath $BaseDirectory -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like 'run-*' } |
        Sort-Object -Property LastWriteTime
    )
}

function Compress-RunDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourceDirectory,

        [Parameter(Mandatory)]
        [string]$ArchiveDirectory
    )

    $resolvedArchiveDirectory = New-SafeDirectory -Path $ArchiveDirectory
    $zipPath = Join-Path -Path $resolvedArchiveDirectory -ChildPath ((Split-Path -Leaf $SourceDirectory) + '.zip')

    if (Test-Path -LiteralPath $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force -ErrorAction Stop
    }

    Compress-Archive -Path $SourceDirectory -DestinationPath $zipPath -CompressionLevel Optimal -Force -ErrorAction Stop
    return $zipPath
}

function Invoke-RunRetentionMaintenance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BaseDirectory,

        [Parameter(Mandatory)]
        [int]$RetentionDays,

        [switch]$ArchiveOldRuns,

        [switch]$PruneOldRuns
    )

    $summary = [PSCustomObject]@{
        BaseDirectory = $BaseDirectory
        RetentionDays = $RetentionDays
        CutoffDate    = Get-RunRetentionCutoff -Days $RetentionDays
        ArchivedCount = 0
        PrunedCount   = 0
        ArchivedPaths = @()
        PrunedPaths   = @()
        FailedPaths   = @()
        ArchiveRoot   = $null
    }

    if (-not $ArchiveOldRuns -and -not $PruneOldRuns) {
        return $summary
    }

    $runDirectories = @(Get-RunDirectories -BaseDirectory $BaseDirectory)
    if ($runDirectories.Count -eq 0) {
        return $summary
    }

    $eligibleDirectories = @($runDirectories | Where-Object { $_.LastWriteTime -lt $summary.CutoffDate })
    if ($eligibleDirectories.Count -eq 0) {
        return $summary
    }

    $archiveRoot = Join-Path -Path $BaseDirectory -ChildPath '_archives'
    if ($ArchiveOldRuns) {
        $summary.ArchiveRoot = New-SafeDirectory -Path $archiveRoot
    }

    foreach ($directory in $eligibleDirectories) {
        try {
            $currentPath = $directory.FullName

            if ($ArchiveOldRuns) {
                $archivedZip = Compress-RunDirectory -SourceDirectory $currentPath -ArchiveDirectory $summary.ArchiveRoot
                $summary.ArchivedPaths += $archivedZip
                $summary.ArchivedCount++
                Remove-Item -LiteralPath $currentPath -Recurse -Force -ErrorAction Stop
                continue
            }

            if ($PruneOldRuns) {
                Remove-Item -LiteralPath $currentPath -Recurse -Force -ErrorAction Stop
                $summary.PrunedPaths += $currentPath
                $summary.PrunedCount++
            }
        }
        catch {
            $summary.FailedPaths += $directory.FullName
        }
    }

    return $summary
}

function Show-RunRetentionSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Summary,

        [switch]$ArchiveOldRuns,

        [switch]$PruneOldRuns
    )

    if (-not $ArchiveOldRuns -and -not $PruneOldRuns) {
        return
    }

    if ($Summary.ArchivedCount -le 0 -and $Summary.PrunedCount -le 0 -and @($Summary.FailedPaths).Count -eq 0) {
        Write-Stage ("Retenção de runs: nada para tratar com corte em {0}." -f $Summary.CutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))
        return
    }

    if ($Summary.ArchivedCount -gt 0) {
        Write-Stage ("Retenção de runs: {0} pasta(s) arquivada(s) em {1}." -f $Summary.ArchivedCount, $Summary.ArchiveRoot)
    }

    if ($Summary.PrunedCount -gt 0) {
        Write-Stage ("Retenção de runs: {0} pasta(s) removida(s) definitivamente." -f $Summary.PrunedCount)
    }

    if (@($Summary.FailedPaths).Count -gt 0) {
        Write-Stage ("Retenção de runs: falha ao tratar {0} pasta(s)." -f @($Summary.FailedPaths).Count)
    }
}

function Save-JsonReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Payload,

        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $Path = Join-Path -Path (Get-Location) -ChildPath ("diagnostico_stress_{0}.json" -f $timestamp)
    }

    $dir = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        $null = New-Item -ItemType Directory -Path $dir -Force
    }

    $Payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
    return $Path
}


function Invoke-ViewerGenerationPrompt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RunOutputDir
    )

    $viewerRoot = 'C:\dev\DiagnosticoStressWindows'
    $viewerDir = Join-Path -Path $PSScriptRoot -ChildPath 'DiagnosticoStressWindows_Viewer'
    $viewerScript = Join-Path -Path $viewerDir -ChildPath 'build-run-viewer.ps1'
    $viewerOutput = Join-Path -Path $viewerDir -ChildPath 'viewer-report.html'

    if (-not (Test-Path -LiteralPath $viewerScript)) {
        Write-Stage 'Viewer: build-run-viewer.ps1 nao encontrado. Pulando geracao visual.'
        return
    }

    $answer = ''
    try {
        Write-Host '' -ForegroundColor DarkGray
        $answer = Read-Host 'Deseja gerar/atualizar o viewer agora? [S/N]'
    }
    catch {
        return
    }

    if ([string]::IsNullOrWhiteSpace($answer)) {
        return
    }

    $normalized = $answer.Trim().ToUpperInvariant()
    if ($normalized -notin @('S', 'SIM', 'Y', 'YES')) {
        return
    }

    Write-Stage ("Gerando viewer com raiz fixa em {0}..." -f $viewerRoot)

    try {
        & $viewerScript -RootPath $viewerRoot -OutputPath $viewerOutput
        if (Test-Path -LiteralPath $viewerOutput) {
            Write-Stage ("Viewer gerado em: {0}" -f $viewerOutput)
            try {
                Start-Process -FilePath $viewerOutput | Out-Null
            }
            catch {}
        }
    }
    catch {
        Write-Stage ("Viewer: falha ao gerar. {0}" -f $_.Exception.Message)
    }
}

try {
    if (-not $NoClear) {

        Clear-Host
    }

    $effectiveSampleSeconds = if ($Quick -and -not $PSBoundParameters.ContainsKey('SampleSeconds')) { 2 } else { $SampleSeconds }
    $effectiveTop = if ($Quick -and -not $PSBoundParameters.ContainsKey('Top')) { 10 } else { $Top }
    $effectiveTimelineIntervalSec = if ($Quick -and -not $PSBoundParameters.ContainsKey('TimelineIntervalSec')) { 1 } else { $TimelineIntervalSec }

    $runStartedAt = Get-Date
    $defaultRootOutputDir = 'C:\dev\DiagnosticoStressWindows'
    $baseOutputDir = if ([string]::IsNullOrWhiteSpace($OutDir)) { $defaultRootOutputDir } else { $OutDir }
    $resolvedBaseOutputDir = New-SafeDirectory -Path $baseOutputDir
    $runMaintenanceSummary = Invoke-RunRetentionMaintenance -BaseDirectory $resolvedBaseOutputDir -RetentionDays $RetentionDays -ArchiveOldRuns:$ArchiveOldRuns -PruneOldRuns:$PruneOldRuns
    $runFolderName = 'run-{0}' -f $runStartedAt.ToString('yyyyMMdd-HHmmss')
    $runOutputDir = New-SafeDirectory -Path (Join-Path -Path $resolvedBaseOutputDir -ChildPath $runFolderName)

    Write-Host '==============================================' -ForegroundColor Cyan
    Write-Host ' DIAGNÓSTICO ROBUSTO DE STRESS - WINDOWS ' -ForegroundColor Cyan
    Write-Host '==============================================' -ForegroundColor Cyan
    Write-Host ('Amostra: {0}s | Top: {1} processos | Intervalo timeline: {2}s{3}' -f $effectiveSampleSeconds, $effectiveTop, $effectiveTimelineIntervalSec, $(if ($Quick) { ' | Modo rápido' } else { '' })) -ForegroundColor DarkGray

    Write-Stage "Pasta de saída da execução: $runOutputDir"
    Show-RunRetentionSummary -Summary $runMaintenanceSummary -ArchiveOldRuns:$ArchiveOldRuns -PruneOldRuns:$PruneOldRuns

    $systemSummary = Get-SystemSummary
    Show-SystemSummary -Summary $systemSummary

    $systemSnapshotExport = Get-SystemSnapshotForExport -SystemSummary $systemSummary
    Save-Json -InputObject $systemSnapshotExport -Path (Join-Path -Path $runOutputDir -ChildPath 'system-snapshot.json')

    Write-Stage 'Capturando processos do navegador para CSV...'
    $browserProcesses = @(Get-BrowserProcesses)
    Export-IfAny -Data $browserProcesses -Path (Join-Path -Path $runOutputDir -ChildPath 'browser-processes.csv')

    Write-Section -Title 'COLETANDO AMOSTRA DE PROCESSOS' -Color DarkYellow
    Write-Host 'Pegando snapshot inicial...' -ForegroundColor DarkGray
    $before = Get-ProcessSnapshot

    Write-Stage ("Amostrando timeline por {0}s em intervalos de {1}s..." -f $effectiveSampleSeconds, $effectiveTimelineIntervalSec)
    $logicalCpuCount = if ($systemSummary.LogicalCpuCount -and [int]$systemSummary.LogicalCpuCount -gt 0) { [int]$systemSummary.LogicalCpuCount } else { [Environment]::ProcessorCount }
    $timeline = Invoke-TimelineSampling -DurationSec $effectiveSampleSeconds -IntervalSec $effectiveTimelineIntervalSec -LogicalCpuCount $logicalCpuCount -TopN $effectiveTop
    Export-IfAny -Data @($timeline.SystemSamples) -Path (Join-Path -Path $runOutputDir -ChildPath 'system-timeseries.csv')
    Export-IfAny -Data @($timeline.ProcessSamples) -Path (Join-Path -Path $runOutputDir -ChildPath 'top-process-timeseries.csv')

    Write-Host 'Pegando snapshot final...' -ForegroundColor DarkGray
    $after = Get-ProcessSnapshot

    $topProcesses = @(Get-ProcessCpuDelta -Before $before -After $after -IntervalSeconds $effectiveSampleSeconds -LogicalCpuCount $logicalCpuCount -TopCount $effectiveTop)
    $serviceMap = @{}

    if ($IncludeServices -and $topProcesses.Count -gt 0) {
        $serviceMap = Get-ServiceMapping -ProcessIds @($topProcesses | Select-Object -ExpandProperty PID)
    }

    Show-TopProcesses -Processes $topProcesses -ServiceMap $serviceMap -IncludeServices:$IncludeServices
    Show-BrowserDiagnostics -BrowserProcesses $browserProcesses

    $findings = @(Get-HealthFindings -SystemSummary $systemSummary -TopProcesses $topProcesses -BrowserProcesses $browserProcesses)
    Show-Findings -Findings $findings
    Show-NextSteps -SystemSummary $systemSummary -TopProcesses $topProcesses -BrowserProcesses $browserProcesses -RunOutputDir $runOutputDir

    if ($EmitJson) {
        $payload = @{
            GeneratedAt          = (Get-Date).ToString('o')
            QuickMode            = [bool]$Quick
            SampleSeconds        = $effectiveSampleSeconds
            TimelineIntervalSec  = $effectiveTimelineIntervalSec
            TopCount             = $effectiveTop
            RunOutputDir         = $runOutputDir
            RunMaintenance       = $runMaintenanceSummary
            SystemSummary        = $systemSummary
            SystemSnapshot       = $systemSnapshotExport
            TopProcesses         = $topProcesses
            BrowserProcesses     = $browserProcesses
            SystemTimeseries     = @($timeline.SystemSamples)
            TopProcessTimeseries = @($timeline.ProcessSamples)
            Findings             = $findings
        }

        $savedPath = Save-JsonReport -Payload $payload -Path $JsonPath
        Write-Section -Title 'RELATÓRIO JSON' -Color Blue
        Write-Host ("Salvo em: {0}" -f $savedPath) -ForegroundColor Blue
    }

    Invoke-ViewerGenerationPrompt -RunOutputDir $runOutputDir
}
catch {
    Write-Host "`n[FALHA] O script tropeçou feio:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.ScriptStackTrace) {
        Write-Host "`n[STACK]" -ForegroundColor DarkRed
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    }
    exit 1
}
