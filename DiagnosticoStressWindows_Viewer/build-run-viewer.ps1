
[CmdletBinding()]
param(
    [string]$RootPath = 'C:\dev\DiagnosticoStressWindows',
    [string]$OutputPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-FullPathSafe {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Path not found: $Path"
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

function Convert-ToNullableDouble {
    [CmdletBinding()]
    param($Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal] -or
        $Value -is [int] -or $Value -is [long] -or $Value -is [float]) {
        return [double]$Value
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $text = $text.Trim().Replace('%','')

    try {
        return [double]::Parse($text, [System.Globalization.CultureInfo]::InvariantCulture)
    }
    catch {}

    try {
        return [double]::Parse($text, [System.Globalization.CultureInfo]::GetCultureInfo('pt-BR'))
    }
    catch {}

    try {
        $normalized = $text.Replace('.', '').Replace(',', '.')
        return [double]::Parse($normalized, [System.Globalization.CultureInfo]::InvariantCulture)
    }
    catch {}

    return $null
}

function Import-CsvSafe {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return @() }

    try {
        return @(Import-Csv -LiteralPath $Path)
    }
    catch {
        return @()
    }
}

function Import-JsonSafe {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $null }

    try {
        $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        if ($raw[0] -eq [char]0xFEFF) {
            $raw = $raw.Substring(1)
        }
        return $raw | ConvertFrom-Json -Depth 20
    }
    catch {
        return $null
    }
}

function Test-RunDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    return (
        (Test-Path -LiteralPath (Join-Path $Path 'system-snapshot.json')) -and
        (Test-Path -LiteralPath (Join-Path $Path 'system-timeseries.csv')) -and
        (Test-Path -LiteralPath (Join-Path $Path 'top-process-timeseries.csv'))
    )
}

function Get-RunDirectories {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Root)

    $dirs = @(
        Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue |
            Where-Object { Test-RunDirectory -Path $_.FullName } |
            Sort-Object Name
    )

    return @($dirs | ForEach-Object { $_.FullName })
}

function Measure-MaxSafe {
    [CmdletBinding()]
    param($Values)

    $nums = @(@($Values) | Where-Object { $null -ne $_ })
    if ($nums.Count -eq 0) { return $null }
    return ($nums | Measure-Object -Maximum).Maximum
}

function Measure-MinSafe {
    [CmdletBinding()]
    param($Values)

    $nums = @(@($Values) | Where-Object { $null -ne $_ })
    if ($nums.Count -eq 0) { return $null }
    return ($nums | Measure-Object -Minimum).Minimum
}

function New-SystemRows {
    [CmdletBinding()]
    param($Rows)

    $out = @()
    foreach ($row in @($Rows)) {
        $out += [PSCustomObject]@{
            Timestamp            = [string]$row.Timestamp
            CpuTotalPct          = Convert-ToNullableDouble $row.CpuTotalPct
            ProcessorQueueLength = Convert-ToNullableDouble $row.ProcessorQueueLength
            AvailableMB          = Convert-ToNullableDouble $row.AvailableMB
            PagesInputPerSec     = Convert-ToNullableDouble $row.PagesInputPerSec
            PagesPerSec          = Convert-ToNullableDouble $row.PagesPerSec
            PageFileUsagePct     = Convert-ToNullableDouble $row.PageFileUsagePct
            DiskBusyPct          = Convert-ToNullableDouble $row.DiskBusyPct
            DiskQueueLength      = Convert-ToNullableDouble $row.DiskQueueLength
            DiskTransfersPerSec  = Convert-ToNullableDouble $row.DiskTransfersPerSec
        }
    }
    return @($out)
}

function New-TopProcessRows {
    [CmdletBinding()]
    param($Rows)

    $out = @()
    foreach ($row in @($Rows)) {
        $out += [PSCustomObject]@{
            Timestamp           = [string]$row.Timestamp
            ProcessName         = [string]$row.ProcessName
            CpuPct              = Convert-ToNullableDouble $row.CpuPct
            WorkingSetPrivateMB = Convert-ToNullableDouble $row.WorkingSetPrivateMB
            InstanceCount       = Convert-ToNullableDouble $row.InstanceCount
        }
    }
    return @($out)
}

function New-BrowserRows {
    [CmdletBinding()]
    param($Rows)

    $out = @()
    foreach ($row in @($Rows)) {
        $out += [PSCustomObject]@{
            Timestamp       = [string]$row.Timestamp
            Browser         = [string]$row.Browser
            PID             = Convert-ToNullableDouble $row.PID
            ProcessName     = [string]$row.ProcessName
            CPUSecondsTotal = Convert-ToNullableDouble $row.CPUSecondsTotal
            WorkingSetMB    = Convert-ToNullableDouble $row.WorkingSetMB
            PrivateMemoryMB = Convert-ToNullableDouble $row.PrivateMemoryMB
            Threads         = Convert-ToNullableDouble $row.Threads
            Handles         = Convert-ToNullableDouble $row.Handles
            MainWindowTitle = [string]$row.MainWindowTitle
            StartTime       = [string]$row.StartTime
            Path            = [string]$row.Path
        }
    }
    return @($out)
}

function New-BrowserSummary {
    [CmdletBinding()]
    param($Rows)

    $grouped = @(@($Rows) | Group-Object Browser)
    $out = @()

    foreach ($group in $grouped) {
        $out += [PSCustomObject]@{
            Browser           = $group.Name
            ProcessCount      = $group.Count
            TotalWorkingSetMB = [Math]::Round([double](($group.Group | Measure-Object -Property WorkingSetMB -Sum).Sum), 2)
            TotalPrivateMB    = [Math]::Round([double](($group.Group | Measure-Object -Property PrivateMemoryMB -Sum).Sum), 2)
            TotalCPUSeconds   = [Math]::Round([double](($group.Group | Measure-Object -Property CPUSecondsTotal -Sum).Sum), 2)
        }
    }

    return @($out | Sort-Object @{Expression='TotalPrivateMB';Descending=$true}, @{Expression='TotalWorkingSetMB';Descending=$true})
}

function New-TopProcessSummary {
    [CmdletBinding()]
    param($Rows)

    $grouped = @(@($Rows) | Group-Object ProcessName)
    $out = @()

    foreach ($group in $grouped) {
        $out += [PSCustomObject]@{
            ProcessName      = $group.Name
            MaxCpuPct        = Measure-MaxSafe -Values @($group.Group | ForEach-Object { $_.CpuPct })
            MaxPrivateMB     = Measure-MaxSafe -Values @($group.Group | ForEach-Object { $_.WorkingSetPrivateMB })
            MaxInstanceCount = Measure-MaxSafe -Values @($group.Group | ForEach-Object { $_.InstanceCount })
            Samples          = $group.Count
        }
    }

    return @($out | Sort-Object @{Expression='MaxCpuPct';Descending=$true}, @{Expression='MaxPrivateMB';Descending=$true})
}

function Get-PagefilePeakMB {
    [CmdletBinding()]
    param($Snapshot)

    try {
        $pages = @($Snapshot.PageFiles)
        if ($pages.Count -eq 0) { return $null }
        return Measure-MaxSafe -Values @($pages | ForEach-Object { Convert-ToNullableDouble $_.PeakUsageMB })
    }
    catch {
        return $null
    }
}

function Get-PagefileCurrentMB {
    [CmdletBinding()]
    param($Snapshot)

    try {
        $pages = @($Snapshot.PageFiles)
        if ($pages.Count -eq 0) { return $null }
        return Measure-MaxSafe -Values @($pages | ForEach-Object { Convert-ToNullableDouble $_.CurrentUsageMB })
    }
    catch {
        return $null
    }
}

function Get-StressScoreLabel {
    [CmdletBinding()]
    param([double]$Score)

    if ($Score -ge 76) { return 'CRÍTICO' }
    if ($Score -ge 56) { return 'ALTO' }
    if ($Score -ge 31) { return 'ATENÇÃO' }
    return 'OK'
}

function Get-StressBand {
    [CmdletBinding()]
    param([double]$Score)

    if ($Score -ge 76) { return 'critical' }
    if ($Score -ge 56) { return 'high' }
    if ($Score -ge 31) { return 'warn' }
    return 'ok'
}

function New-Metrics {
    [CmdletBinding()]
    param(
        [AllowNull()]$Snapshot,
        $SystemRows,
        $BrowserSummary,
        $TopProcessSummary
    )

    $availableMin = Measure-MinSafe -Values @($SystemRows | ForEach-Object { $_.AvailableMB })
    $availableMax = Measure-MaxSafe -Values @($SystemRows | ForEach-Object { $_.AvailableMB })
    $cpuMax = Measure-MaxSafe -Values @($SystemRows | ForEach-Object { $_.CpuTotalPct })
    $queueMax = Measure-MaxSafe -Values @($SystemRows | ForEach-Object { $_.ProcessorQueueLength })
    $pagePctMax = Measure-MaxSafe -Values @($SystemRows | ForEach-Object { $_.PageFileUsagePct })
    $pagesPerSecMax = Measure-MaxSafe -Values @($SystemRows | ForEach-Object { $_.PagesPerSec })
    $diskBusyMax = Measure-MaxSafe -Values @($SystemRows | ForEach-Object { $_.DiskBusyPct })
    $pagePeak = Get-PagefilePeakMB -Snapshot $Snapshot
    $pageCurrent = Get-PagefileCurrentMB -Snapshot $Snapshot
    $browserPrivateMax = Measure-MaxSafe -Values @($BrowserSummary | Where-Object { $null -ne $_ -and $_.PSObject.Properties['TotalPrivateMB'] } | ForEach-Object { $_.TotalPrivateMB })
    $browserWorkingSetMax = Measure-MaxSafe -Values @($BrowserSummary | Where-Object { $null -ne $_ -and $_.PSObject.Properties['TotalWorkingSetMB'] } | ForEach-Object { $_.TotalWorkingSetMB })

    $dominantBrowser = if (@($BrowserSummary).Count -gt 0) { @($BrowserSummary)[0].Browser } else { $null }
    $dominantProcess = if (@($TopProcessSummary).Count -gt 0) { @($TopProcessSummary)[0].ProcessName } else { $null }

    $score = 0.0

    if ($null -ne $availableMin) {
        if ($availableMin -le 150) { $score += 38 }
        elseif ($availableMin -le 250) { $score += 30 }
        elseif ($availableMin -le 400) { $score += 22 }
        elseif ($availableMin -le 700) { $score += 12 }
    }

    if ($null -ne $pagePeak) {
        if ($pagePeak -ge 1600) { $score += 24 }
        elseif ($pagePeak -ge 1000) { $score += 18 }
        elseif ($pagePeak -ge 500) { $score += 10 }
        elseif ($pagePeak -gt 0) { $score += 4 }
    }

    if ($null -ne $cpuMax) {
        if ($cpuMax -ge 85) { $score += 18 }
        elseif ($cpuMax -ge 60) { $score += 12 }
        elseif ($cpuMax -ge 35) { $score += 7 }
    }

    if ($null -ne $queueMax) {
        if ($queueMax -ge 12) { $score += 12 }
        elseif ($queueMax -ge 6) { $score += 8 }
        elseif ($queueMax -ge 2) { $score += 4 }
    }

    if ($null -ne $browserPrivateMax) {
        if ($browserPrivateMax -ge 1400) { $score += 10 }
        elseif ($browserPrivateMax -ge 900) { $score += 7 }
        elseif ($browserPrivateMax -ge 500) { $score += 4 }
    }

    if ($score -gt 100) { $score = 100 }

    return [PSCustomObject]@{
        AvailableMinMB       = if ($null -ne $availableMin) { [Math]::Round([double]$availableMin, 2) } else { $null }
        AvailableMaxMB       = if ($null -ne $availableMax) { [Math]::Round([double]$availableMax, 2) } else { $null }
        PagefilePeakMB       = if ($null -ne $pagePeak) { [Math]::Round([double]$pagePeak, 2) } else { $null }
        PagefileCurrentMB    = if ($null -ne $pageCurrent) { [Math]::Round([double]$pageCurrent, 2) } else { $null }
        PagefileUsagePctMax  = if ($null -ne $pagePctMax) { [Math]::Round([double]$pagePctMax, 2) } else { $null }
        CpuMaxPct            = if ($null -ne $cpuMax) { [Math]::Round([double]$cpuMax, 2) } else { $null }
        QueueMax             = if ($null -ne $queueMax) { [Math]::Round([double]$queueMax, 2) } else { $null }
        PagesPerSecMax       = if ($null -ne $pagesPerSecMax) { [Math]::Round([double]$pagesPerSecMax, 2) } else { $null }
        DiskBusyMax          = if ($null -ne $diskBusyMax) { [Math]::Round([double]$diskBusyMax, 2) } else { $null }
        BrowserPrivateMaxMB  = if ($null -ne $browserPrivateMax) { [Math]::Round([double]$browserPrivateMax, 2) } else { $null }
        BrowserWorkingSetMaxMB = if ($null -ne $browserWorkingSetMax) { [Math]::Round([double]$browserWorkingSetMax, 2) } else { $null }
        DominantBrowser      = $dominantBrowser
        DominantProcess      = $dominantProcess
        Score                = [Math]::Round($score, 0)
        ScoreLabel           = Get-StressScoreLabel -Score $score
        ScoreBand            = Get-StressBand -Score $score
    }
}

function New-RunModel {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Folder)

    $snapshot = Import-JsonSafe -Path (Join-Path $Folder 'system-snapshot.json')
    $systemRows = New-SystemRows -Rows @(Import-CsvSafe -Path (Join-Path $Folder 'system-timeseries.csv'))
    $topRows = New-TopProcessRows -Rows @(Import-CsvSafe -Path (Join-Path $Folder 'top-process-timeseries.csv'))
    $browserRows = New-BrowserRows -Rows @(Import-CsvSafe -Path (Join-Path $Folder 'browser-processes.csv'))
    $browserSummary = New-BrowserSummary -Rows $browserRows
    $processSummary = New-TopProcessSummary -Rows $topRows
    $metrics = New-Metrics -Snapshot $snapshot -SystemRows $systemRows -BrowserSummary $browserSummary -TopProcessSummary $processSummary

    $timestamp = $null
    if ($null -ne $snapshot -and -not [string]::IsNullOrWhiteSpace([string]$snapshot.Timestamp)) {
        $timestamp = [string]$snapshot.Timestamp
    }
    elseif (@($systemRows).Count -gt 0) {
        $timestamp = [string]@($systemRows)[0].Timestamp
    }
    else {
        $timestamp = Split-Path -Leaf $Folder
    }

    return [PSCustomObject]@{
        RunName              = (Split-Path -Leaf $Folder)
        FolderPath           = $Folder
        Timestamp            = $timestamp
        LastBoot             = if ($null -ne $snapshot) { [string]$snapshot.LastBoot } else { $null }
        Snapshot             = $snapshot
        SystemTimeseries     = @($systemRows)
        TopProcessTimeseries = @($topRows)
        BrowserProcesses     = @($browserRows)
        BrowserSummary       = @($browserSummary)
        TopProcessSummary    = @($processSummary)
        Metrics              = $metrics
    }
}

function Get-EmbeddedHtmlTemplate {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$JsonPayload)

    $safeJson = $JsonPayload.Replace('</script>', '<\/script>')

    $template = @'
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Diagnóstico Stress Windows — Viewer</title>
<style>
:root{
  --bg:#08111c;--bg2:#0b1624;--panel:#0f1b2d;--panel2:#12233a;--line:#24364f;--text:#edf3fb;--muted:#9eb2ca;
  --ok:#23c16b;--warn:#f5ad2e;--high:#ff7a3d;--critical:#ff4d57;--blue:#4da3ff;--cyan:#2dd4bf;
}
*{box-sizing:border-box}
html,body{margin:0;background:linear-gradient(180deg,var(--bg),#060d16 55%,#050a12);color:var(--text);font:14px/1.45 "Segoe UI",Arial,sans-serif}
body{min-height:100vh}
.app{display:grid;grid-template-columns:320px 1fr;min-height:100vh}
.sidebar{background:rgba(8,17,28,.92);border-right:1px solid var(--line);padding:20px;position:sticky;top:0;height:100vh;overflow:auto}
.brand{display:flex;align-items:center;gap:10px;margin-bottom:8px}
.brand-mark{width:32px;height:32px;border-radius:12px;display:grid;place-items:center;background:linear-gradient(135deg,#143861,#0b5db3);font-weight:900}
.sidebar h1{font-size:20px;margin:0}
.sidebar p{margin:0 0 16px;color:var(--muted)}
.search{width:100%;padding:11px 12px;border-radius:14px;border:1px solid var(--line);background:#0a1421;color:var(--text);outline:none}
.run-list{display:grid;gap:10px;margin-top:14px}
.run-item{border:1px solid var(--line);background:rgba(15,27,45,.72);border-radius:16px;padding:12px 13px;cursor:pointer}
.run-item:hover{background:#11213a}
.run-item.active{background:#163055;border-color:#36669f;box-shadow:0 0 0 1px rgba(77,163,255,.18) inset}
.run-top{display:flex;justify-content:space-between;align-items:flex-start;gap:10px}
.run-name{font-weight:800;font-size:14px;word-break:break-word}
.run-meta{margin-top:8px;color:var(--muted);font-size:12px}
.main{padding:24px 28px}
.hero{display:grid;grid-template-columns:1.4fr .95fr;gap:16px;align-items:stretch}
.panel{background:linear-gradient(180deg,rgba(16,29,48,.92),rgba(10,19,31,.96));border:1px solid var(--line);border-radius:24px;padding:18px}
.hero-title{display:flex;justify-content:space-between;gap:14px;align-items:flex-start;flex-wrap:wrap}
.hero-title h2{margin:0;font-size:28px;line-height:1.1}
.hero-sub{margin-top:6px;color:var(--muted);font-size:13px}
.badges{display:flex;gap:8px;flex-wrap:wrap}
.badge{display:inline-flex;align-items:center;gap:6px;padding:6px 10px;border-radius:999px;font-size:12px;font-weight:800;border:1px solid var(--line);background:#14243c}
.badge.ok{color:#9cf0bf;background:rgba(35,193,107,.13);border-color:rgba(35,193,107,.24)}
.badge.warn{color:#ffd88c;background:rgba(245,173,46,.13);border-color:rgba(245,173,46,.24)}
.badge.high{color:#ffb08b;background:rgba(255,122,61,.12);border-color:rgba(255,122,61,.22)}
.badge.critical{color:#ffacb1;background:rgba(255,77,87,.12);border-color:rgba(255,77,87,.22)}
.diagnosis{margin-top:16px;display:grid;grid-template-columns:1fr 1fr;gap:14px}
.callout{padding:16px;border-radius:18px;background:rgba(7,15,25,.55);border:1px solid var(--line)}
.callout h3{margin:0 0 8px;font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}
.callout p{margin:0;font-size:15px;line-height:1.55}
.bullets{margin:0;padding-left:18px}
.bullets li{margin:6px 0;color:#d9e6f7}
.score-wrap{display:grid;place-items:center;min-height:260px}
.score-circle{width:210px;height:210px;border-radius:50%;display:grid;place-items:center;border:10px solid #233852;background:radial-gradient(circle at 40% 35%,#17365e 0%,#0c1625 68%)}
.score-number{font-size:58px;font-weight:900;line-height:1}
.score-label{font-size:13px;color:var(--muted);font-weight:700;margin-top:6px;text-transform:uppercase;letter-spacing:.08em}
.kpi-grid{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px;margin-top:16px}
.kpi{padding:16px;border-radius:20px;background:rgba(7,15,25,.55);border:1px solid var(--line)}
.kpi-label{font-size:12px;color:var(--muted);margin-bottom:8px}
.kpi-value{font-size:28px;font-weight:900}
.kpi-sub{font-size:12px;color:var(--muted);margin-top:6px}
.section{margin-top:18px}
.section-head{display:flex;justify-content:space-between;align-items:flex-end;gap:12px;margin-bottom:10px;flex-wrap:wrap}
.section h3{margin:0;font-size:19px}
.section p{margin:0;color:var(--muted)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.chart-card{padding:16px;border-radius:22px;background:linear-gradient(180deg,rgba(15,27,45,.88),rgba(8,15,26,.95));border:1px solid var(--line)}
.chart-top{display:flex;justify-content:space-between;gap:12px;align-items:flex-start;margin-bottom:10px}
.chart-title{font-size:14px;font-weight:800}
.chart-desc{font-size:12px;color:var(--muted)}
.chart-pill{padding:6px 10px;border-radius:999px;background:#152844;border:1px solid var(--line);font-size:12px;font-weight:800}
.chart-svg{width:100%;height:260px;display:block;border-radius:16px;background:linear-gradient(180deg,#091321,#0a1625)}
.legend{display:flex;gap:14px;flex-wrap:wrap;margin-top:8px}
.legend-item{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--muted)}
.legend-dot{width:10px;height:10px;border-radius:999px}
.pressure-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
.pressure{padding:14px;border-radius:18px;background:rgba(7,15,25,.55);border:1px solid var(--line)}
.pressure .title{font-size:12px;color:var(--muted);margin-bottom:8px}
.bar{height:10px;border-radius:999px;background:#0b1522;overflow:hidden;border:1px solid #1b2b40}
.fill{height:100%;border-radius:999px}
.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:18px;background:rgba(11,21,34,.78)}
table{width:100%;border-collapse:collapse;min-width:720px}
th,td{padding:11px 13px;border-bottom:1px solid rgba(36,54,79,.82);text-align:left;vertical-align:top}
th{background:#13253b;color:var(--muted);font-size:12px;position:sticky;top:0}
.row-quiet{color:var(--muted)}
.metric-cell{display:flex;align-items:center;gap:10px}
.metric-bar{flex:1;min-width:90px;height:8px;background:#0b1522;border-radius:999px;overflow:hidden;border:1px solid #1b2b40}
.metric-bar > span{display:block;height:100%}
.comp-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px}
.comp-card{padding:14px;border-radius:18px;background:rgba(7,15,25,.55);border:1px solid var(--line)}
.comp-card h4{margin:0 0 6px;font-size:15px}
.comp-row{display:flex;justify-content:space-between;gap:12px;font-size:12px;color:var(--muted);margin-top:4px}
.footer{margin-top:16px;font-size:12px;color:var(--muted)}
.empty{padding:36px;border:1px dashed var(--line);border-radius:20px;color:var(--muted);background:rgba(7,15,25,.4)}
code{background:#0a1421;border:1px solid var(--line);padding:2px 6px;border-radius:8px}
@media (max-width:1200px){
  .app{grid-template-columns:1fr}
  .sidebar{position:static;height:auto;border-right:none;border-bottom:1px solid var(--line)}
  .hero,.diagnosis,.grid2{grid-template-columns:1fr}
  .kpi-grid,.pressure-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
}
@media (max-width:720px){
  .main{padding:16px}
  .kpi-grid,.pressure-grid{grid-template-columns:1fr}
}
</style>
</head>
<body>
<div class="app">
  <aside class="sidebar">
    <div class="brand"><div class="brand-mark">DS</div><div><h1>Viewer de Runs</h1><p>Agora com um mínimo de hierarquia. Milagre da engenharia.</p></div></div>
    <input id="runSearch" class="search" type="search" placeholder="Filtrar run, navegador ou processo..." />
    <div id="runList" class="run-list"></div>
  </aside>
  <main class="main"><div id="appContent"></div></main>
</div>

<script id="embedded-data" type="application/json">__PAYLOAD_JSON__</script>
<script>
(function(){
  const payload = JSON.parse(document.getElementById('embedded-data').textContent);
  const runs = Array.isArray(payload.Runs) ? payload.Runs : [];
  const state = { selectedRunName: runs.length ? runs[runs.length - 1].RunName : null, search: '' };
  const runListEl = document.getElementById('runList');
  const appContentEl = document.getElementById('appContent');
  const runSearchEl = document.getElementById('runSearch');

  const COLORS = {
    cpu:'#4da3ff',
    mem:'#23c16b',
    page:'#f5ad2e',
    queue:'#ff4d57',
    pages:'#8b5cf6',
    grid:'#203249',
    axis:'#38526f'
  };

  function n(value){
    if(value === null || value === undefined || value === '' || Number.isNaN(Number(value))) return null;
    return Number(value);
  }

  function fmt(value, digits = 0, suffix = ''){
    const v = n(value);
    if(v === null) return 'N/D';
    return v.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: digits }) + suffix;
  }

  function fmtTs(text){
    if(!text) return 'N/D';
    try{
      const d = new Date(text);
      if(Number.isNaN(d.getTime())) return text;
      return d.toLocaleString('pt-BR');
    }catch{ return text; }
  }

  function band(score){
    const s = n(score) || 0;
    if(s >= 76) return 'critical';
    if(s >= 56) return 'high';
    if(s >= 31) return 'warn';
    return 'ok';
  }

  function label(score){
    const s = n(score) || 0;
    if(s >= 76) return 'CRÍTICO';
    if(s >= 56) return 'ALTO';
    if(s >= 31) return 'ATENÇÃO';
    return 'OK';
  }

  function filteredRuns(){
    const q = (state.search || '').trim().toLowerCase();
    if(!q) return runs;
    return runs.filter(run => {
      const bag = [
        run.RunName,
        run.Metrics?.DominantBrowser,
        run.Metrics?.DominantProcess,
        run.Snapshot?.ComputerName,
        run.Snapshot?.CPUName
      ].filter(Boolean).join(' ').toLowerCase();
      return bag.includes(q);
    });
  }

  function getSelectedRun(){
    const list = filteredRuns();
    let selected = list.find(r => r.RunName === state.selectedRunName);
    if(!selected && list.length){
      selected = list[list.length - 1];
      state.selectedRunName = selected.RunName;
    }
    return selected || null;
  }

  function getDiagnosis(run){
    const m = run.Metrics || {};
    const parts = [];
    const avail = n(m.AvailableMinMB);
    const page = n(m.PagefilePeakMB);
    const cpu = n(m.CpuMaxPct);
    const queue = n(m.QueueMax);
    const browser = m.DominantBrowser || 'N/D';
    const proc = m.DominantProcess || 'N/D';

    if(avail !== null){
      if(avail <= 200) parts.push(`RAM livre mínima em ${fmt(avail,0,' MB')}: isso já é terreno de input lag e pagefile entrando no soco.`);
      else if(avail <= 400) parts.push(`RAM livre mínima em ${fmt(avail,0,' MB')}: sessão apertada, com pouco espaço para navegador e tranqueiras residentes.`);
      else parts.push(`RAM ainda respirando melhor (${fmt(avail,0,' MB')} livres no pior ponto), então o gargalo não é só memória pura.`);
    }

    if(page !== null){
      if(page >= 1000) parts.push(`Pagefile bateu ${fmt(page,0,' MB')}, então já teve pressão real de memória, não foi chilique imaginário.`);
      else if(page > 0) parts.push(`Pagefile apareceu (${fmt(page,0,' MB')} de pico), sinal de que a sessão já começou a empurrar coisa pro disco.`);
      else parts.push(`Sem uso relevante de pagefile. Pelo menos uma coisa se comportou.`);
    }

    if(cpu !== null){
      if(cpu >= 80) parts.push(`CPU total chegou a ${fmt(cpu,0,'%')}. Quando isso acontece num Core 2 Duo, o mouse começa a pedir aposentadoria.`);
      else if(cpu >= 40) parts.push(`CPU teve pico de ${fmt(cpu,0,'%')}, suficiente para deixar a interface menos responsiva.`);
      else parts.push(`CPU não foi o vilão principal: pico de ${fmt(cpu,0,'%')}.`);
    }

    if(queue !== null && queue >= 6){
      parts.push(`Processor Queue foi a ${fmt(queue,0)}, então teve fila de execução suficiente pra dar aquela sensação de engasgo em cascata.`);
    }

    parts.push(`Navegador dominante: ${browser}. Processo dominante: ${proc}. Sim, alguém precisava levar a culpa pelo nome.`);
    return parts.join(' ');
  }

  function getHighlights(run){
    const m = run.Metrics || {};
    const out = [];
    const avail = n(m.AvailableMinMB);
    const page = n(m.PagefilePeakMB);
    const cpu = n(m.CpuMaxPct);
    const queue = n(m.QueueMax);

    out.push(avail !== null && avail <= 250 ? `Memória estrangulada (${fmt(avail,0,' MB')} livres no mínimo).` : `Memória mínima em ${fmt(avail,0,' MB')}.`);
    out.push(page !== null ? `Pico de pagefile em ${fmt(page,0,' MB')}.` : 'Sem pico de pagefile disponível no snapshot.');
    out.push(`CPU máxima em ${fmt(cpu,0,'%')}. Queue máxima em ${fmt(queue,0)}.`);
    return out;
  }

  function buildPolyline(points, width, height, padding, minVal, maxVal){
    const span = (maxVal - minVal) || 1;
    return points.map((p, index) => {
      const x = padding + (index * ((width - padding * 2) / Math.max(points.length - 1, 1)));
      const y = height - padding - (((p - minVal) / span) * (height - padding * 2));
      return `${x},${y}`;
    }).join(' ');
  }

  function renderLineChart(title, description, rows, series, options = {}){
    const width = 760, height = 260, padding = 34;
    const labels = rows.map(r => r.Timestamp || '');
    const seriesData = series.map(s => ({
      ...s,
      values: rows.map(r => n(r[s.key]))
    })).filter(s => s.values.some(v => v !== null));

    if(!seriesData.length){
      return `<div class="chart-card"><div class="chart-top"><div><div class="chart-title">${title}</div><div class="chart-desc">${description}</div></div></div><div class="empty">Sem dados suficientes para este gráfico.</div></div>`;
    }

    const allValues = seriesData.flatMap(s => s.values.filter(v => v !== null));
    const minVal = options.min !== undefined ? options.min : Math.min(...allValues);
    const maxVal = options.max !== undefined ? options.max : Math.max(...allValues);
    const yMin = minVal === maxVal ? minVal - 1 : minVal;
    const yMax = minVal === maxVal ? maxVal + 1 : maxVal;

    let svg = `<svg class="chart-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none">`;
    for(let i=0;i<4;i++){
      const y = padding + (i * ((height - padding * 2) / 3));
      svg += `<line x1="${padding}" y1="${y}" x2="${width-padding}" y2="${y}" stroke="${COLORS.grid}" stroke-width="1"/>`;
    }
    svg += `<line x1="${padding}" y1="${height-padding}" x2="${width-padding}" y2="${height-padding}" stroke="${COLORS.axis}" stroke-width="1"/>`;
    svg += `<line x1="${padding}" y1="${padding}" x2="${padding}" y2="${height-padding}" stroke="${COLORS.axis}" stroke-width="1"/>`;

    const yTicks = [yMax, yMin + ((yMax-yMin)*2/3), yMin + ((yMax-yMin)/3), yMin];
    yTicks.forEach((tick, idx) => {
      const y = padding + (idx * ((height - padding * 2) / 3));
      svg += `<text x="8" y="${y+4}" fill="#91a8c5" font-size="11">${fmt(tick,0,options.suffix || '')}</text>`;
    });

    const labelStep = Math.max(Math.floor(labels.length / 4), 1);
    labels.forEach((label, idx) => {
      if(idx % labelStep !== 0 && idx !== labels.length - 1) return;
      const x = padding + (idx * ((width - padding * 2) / Math.max(labels.length - 1, 1)));
      let short = label;
      if(short && short.includes('T')) short = short.split('T')[1];
      if(short && short.length > 8) short = short.slice(0,8);
      svg += `<text x="${x}" y="${height-10}" text-anchor="middle" fill="#91a8c5" font-size="11">${short || ''}</text>`;
    });

    seriesData.forEach(s => {
      const values = s.values.map(v => v === null ? yMin : v);
      const points = buildPolyline(values, width, height, padding, yMin, yMax);
      svg += `<polyline fill="none" stroke="${s.color}" stroke-width="3" points="${points}" stroke-linecap="round" stroke-linejoin="round"/>`;
      values.forEach((v, idx) => {
        const x = padding + (idx * ((width - padding * 2) / Math.max(values.length - 1, 1)));
        const y = height - padding - (((v - yMin) / ((yMax-yMin)||1)) * (height - padding * 2));
        svg += `<circle cx="${x}" cy="${y}" r="3.2" fill="${s.color}" opacity="0.95"/>`;
      });
    });
    svg += `</svg>`;

    const topRight = seriesData.map(s => `${s.label}: pico ${fmt(Math.max(...s.values.filter(v => v !== null)),0,options.suffixMap?.[s.key] || options.suffix || '')}`).join(' · ');

    const legend = seriesData.map(s => `<div class="legend-item"><span class="legend-dot" style="background:${s.color}"></span>${s.label}</div>`).join('');
    return `<div class="chart-card">
      <div class="chart-top">
        <div><div class="chart-title">${title}</div><div class="chart-desc">${description}</div></div>
        <div class="chart-pill">${topRight}</div>
      </div>
      ${svg}
      <div class="legend">${legend}</div>
    </div>`;
  }

  function pressureBar(title, value, max, color, suffix=''){
    const v = n(value);
    const ratio = v === null ? 0 : Math.max(0, Math.min(100, (v / max) * 100));
    return `<div class="pressure"><div class="title">${title}</div><div class="metric-cell"><strong>${fmt(v,0,suffix)}</strong><div class="metric-bar"><span style="width:${ratio}%;background:${color}"></span></div></div></div>`;
  }

  function metricRowBar(value, max, color){
    const v = n(value);
    const ratio = v === null ? 0 : Math.max(0, Math.min(100, max > 0 ? (v / max) * 100 : 0));
    return `<div class="metric-bar"><span style="width:${ratio}%;background:${color}"></span></div>`;
  }

  function renderSidebar(){
    const list = filteredRuns();
    runListEl.innerHTML = list.map(run => {
      const m = run.Metrics || {};
      return `<div class="run-item ${run.RunName === state.selectedRunName ? 'active' : ''}" data-run="${run.RunName}">
        <div class="run-top">
          <div class="run-name">${run.RunName}</div>
          <span class="badge ${band(m.Score)}">${label(m.Score)}</span>
        </div>
        <div class="run-meta">score ${fmt(m.Score,0)} · RAM min ${fmt(m.AvailableMinMB,0,' MB')}<br/>pagefile ${fmt(m.PagefilePeakMB,0,' MB')} · CPU ${fmt(m.CpuMaxPct,0,'%')}</div>
      </div>`;
    }).join('') || `<div class="empty">Nenhuma run carregada.</div>`;

    Array.from(runListEl.querySelectorAll('.run-item')).forEach(el => {
      el.addEventListener('click', () => {
        state.selectedRunName = el.getAttribute('data-run');
        renderAll();
      });
    });
  }

  function renderBrowsers(summary){
    if(!summary.length){
      return `<div class="table-wrap"><table><thead><tr><th>Browser</th><th>Processos</th><th>Working Set</th><th>Private</th><th>CPU acumulada</th></tr></thead><tbody><tr><td colspan="5" class="row-quiet">Sem browser-processes.csv nessa run.</td></tr></tbody></table></div>`;
    }
    const maxPrivate = Math.max(...summary.map(r => n(r.TotalPrivateMB) || 0), 1);
    const rows = summary.map(r => `<tr>
      <td><strong>${r.Browser}</strong></td>
      <td>${fmt(r.ProcessCount,0)}</td>
      <td>${fmt(r.TotalWorkingSetMB,0,' MB')}</td>
      <td><div class="metric-cell"><span>${fmt(r.TotalPrivateMB,0,' MB')}</span>${metricRowBar(r.TotalPrivateMB, maxPrivate, COLORS.mem)}</div></td>
      <td>${fmt(r.TotalCPUSeconds,0,' s')}</td>
    </tr>`).join('');
    return `<div class="table-wrap"><table><thead><tr><th>Browser</th><th>Processos</th><th>Working Set</th><th>Memória privada</th><th>CPU acumulada</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  }

  function renderProcesses(summary){
    if(!summary.length){
      return `<div class="table-wrap"><table><thead><tr><th>Processo</th><th>CPU max</th><th>Memória privada max</th><th>Instâncias max</th><th>Amostras</th></tr></thead><tbody><tr><td colspan="5" class="row-quiet">Sem dados de processo.</td></tr></tbody></table></div>`;
    }
    const top = summary.slice(0, 12);
    const maxCpu = Math.max(...top.map(r => n(r.MaxCpuPct) || 0), 1);
    const maxMem = Math.max(...top.map(r => n(r.MaxPrivateMB) || 0), 1);
    const rows = top.map(r => `<tr>
      <td><strong>${r.ProcessName}</strong></td>
      <td><div class="metric-cell"><span>${fmt(r.MaxCpuPct,0,'%')}</span>${metricRowBar(r.MaxCpuPct, maxCpu, COLORS.cpu)}</div></td>
      <td><div class="metric-cell"><span>${fmt(r.MaxPrivateMB,0,' MB')}</span>${metricRowBar(r.MaxPrivateMB, maxMem, COLORS.page)}</div></td>
      <td>${fmt(r.MaxInstanceCount,0)}</td>
      <td>${fmt(r.Samples,0)}</td>
    </tr>`).join('');
    return `<div class="table-wrap"><table><thead><tr><th>Processo</th><th>CPU max</th><th>Memória privada max</th><th>Instâncias max</th><th>Amostras</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  }

  function renderComparison(allRuns){
    const cards = allRuns.slice().sort((a,b) => (n(b.Metrics?.Score)||0) - (n(a.Metrics?.Score)||0)).map(run => {
      const m = run.Metrics || {};
      return `<div class="comp-card">
        <h4>${run.RunName}</h4>
        <div><span class="badge ${band(m.Score)}">${label(m.Score)} · ${fmt(m.Score,0)}</span></div>
        <div class="comp-row"><span>RAM min</span><strong>${fmt(m.AvailableMinMB,0,' MB')}</strong></div>
        <div class="comp-row"><span>Pagefile pico</span><strong>${fmt(m.PagefilePeakMB,0,' MB')}</strong></div>
        <div class="comp-row"><span>CPU max</span><strong>${fmt(m.CpuMaxPct,0,'%')}</strong></div>
        <div class="comp-row"><span>Navegador</span><strong>${m.DominantBrowser || 'N/D'}</strong></div>
      </div>`;
    }).join('');
    return `<div class="comp-grid">${cards}</div>`;
  }

  function renderSelectedRun(run){
    if(!run){
      return `<div class="empty">Nenhuma run encontrada. Verifique se a pasta possui <code>system-snapshot.json</code>, <code>system-timeseries.csv</code> e <code>top-process-timeseries.csv</code>.</div>`;
    }

    const m = run.Metrics || {};
    const s = run.Snapshot || {};
    const systemRows = Array.isArray(run.SystemTimeseries) ? run.SystemTimeseries : [];
    const browserSummary = Array.isArray(run.BrowserSummary) ? run.BrowserSummary : [];
    const topSummary = Array.isArray(run.TopProcessSummary) ? run.TopProcessSummary : [];
    const drive0 = Array.isArray(s.Drives) && s.Drives.length ? s.Drives[0] : null;
    const highlights = getHighlights(run).map(x => `<li>${x}</li>`).join('');

    return `
      <div class="hero">
        <div class="panel">
          <div class="hero-title">
            <div>
              <h2>${run.RunName}</h2>
              <div class="hero-sub">Gerado em ${fmtTs(run.Timestamp)} · boot ${fmtTs(run.LastBoot)} · máquina ${s.ComputerName || 'N/D'} · ${s.CPUName || 'N/D'}</div>
            </div>
            <div class="badges">
              <span class="badge ${band(m.Score)}">Stress score ${fmt(m.Score,0)} · ${label(m.Score)}</span>
              <span class="badge">navegador dominante: ${m.DominantBrowser || 'N/D'}</span>
              <span class="badge">processo dominante: ${m.DominantProcess || 'N/D'}</span>
            </div>
          </div>
          <div class="diagnosis">
            <div class="callout">
              <h3>Leitura executiva</h3>
              <p>${getDiagnosis(run)}</p>
            </div>
            <div class="callout">
              <h3>O que olhar primeiro</h3>
              <ul class="bullets">${highlights}</ul>
            </div>
          </div>
          <div class="kpi-grid">
            <div class="kpi"><div class="kpi-label">RAM livre mínima</div><div class="kpi-value">${fmt(m.AvailableMinMB,0)}<span style="font-size:14px;color:var(--muted)"> MB</span></div><div class="kpi-sub">quanto menos, pior</div></div>
            <div class="kpi"><div class="kpi-label">Pico de pagefile</div><div class="kpi-value">${fmt(m.PagefilePeakMB,0)}<span style="font-size:14px;color:var(--muted)"> MB</span></div><div class="kpi-sub">swap real do Windows</div></div>
            <div class="kpi"><div class="kpi-label">CPU máxima</div><div class="kpi-value">${fmt(m.CpuMaxPct,0)}<span style="font-size:14px;color:var(--muted)"> %</span></div><div class="kpi-sub">pico do sistema</div></div>
            <div class="kpi"><div class="kpi-label">Queue máxima</div><div class="kpi-value">${fmt(m.QueueMax,0)}</div><div class="kpi-sub">fila de execução</div></div>
            <div class="kpi"><div class="kpi-label">Browser private max</div><div class="kpi-value">${fmt(m.BrowserPrivateMaxMB,0)}<span style="font-size:14px;color:var(--muted)"> MB</span></div><div class="kpi-sub">peso agregado do browser</div></div>
          </div>
        </div>
        <div class="panel score-wrap">
          <div class="score-circle" style="border-color:${m.ScoreBand === 'critical' ? 'var(--critical)' : m.ScoreBand === 'high' ? 'var(--high)' : m.ScoreBand === 'warn' ? 'var(--warn)' : 'var(--ok)'}">
            <div>
              <div class="score-number">${fmt(m.Score,0)}</div>
              <div class="score-label">${label(m.Score)}</div>
            </div>
          </div>
          <div style="margin-top:16px;width:100%">
            ${pressureBar('CPU', m.CpuMaxPct, 100, COLORS.cpu, '%')}
            ${pressureBar('RAM apertada', m.AvailableMinMB !== null ? Math.max(0, 1000 - m.AvailableMinMB) : null, 1000, COLORS.mem, ' pts')}
            ${pressureBar('Pagefile', m.PagefilePeakMB, 1600, COLORS.page, ' MB')}
            ${pressureBar('Fila', m.QueueMax, 16, COLORS.queue, '')}
          </div>
        </div>
      </div>

      <div class="section">
        <div class="section-head">
          <div><h3>Telemetria explicada</h3><p>Agora com rótulo decente. O eixo X são as amostras da run; o Y é a métrica original.</p></div>
        </div>
        <div class="grid2">
          ${renderLineChart('CPU total × Pagefile (%)', 'Mostra pico de CPU do sistema e uso percentual do pagefile no mesmo tempo.', systemRows, [
            { key:'CpuTotalPct', label:'CPU total', color:COLORS.cpu },
            { key:'PageFileUsagePct', label:'Pagefile %', color:COLORS.page }
          ], { suffix:'%', min:0 })}
          ${renderLineChart('Memória disponível (MB)', 'Quando essa linha despenca, o teclado começa a parecer sabotagem.', systemRows, [
            { key:'AvailableMB', label:'RAM disponível', color:COLORS.mem }
          ], { suffix:' MB' })}
          ${renderLineChart('Fila de processador × Pages/sec', 'Fila alta + pages/sec alto é receita de engasgo clássico.', systemRows, [
            { key:'ProcessorQueueLength', label:'Queue', color:COLORS.queue },
            { key:'PagesPerSec', label:'Pages/sec', color:COLORS.pages }
          ], { min:0 })}
          ${renderLineChart('Disco ocupado × transferências', 'Nem sempre o disco é o vilão, então aqui fica fácil ver se ele entrou na briga.', systemRows, [
            { key:'DiskBusyPct', label:'Disk busy %', color:'#ff7a3d' },
            { key:'DiskTransfersPerSec', label:'Transfers/sec', color:COLORS.cyan }
          ], { min:0 })}
        </div>
      </div>

      <div class="section">
        <div class="section-head">
          <div><h3>Contexto da máquina</h3><p>Base física e ambiente dessa run.</p></div>
        </div>
        <div class="kpi-grid">
          <div class="kpi"><div class="kpi-label">RAM física</div><div class="kpi-value">${fmt(s.TotalPhysicalMemoryGB,2)}<span style="font-size:14px;color:var(--muted)"> GB</span></div></div>
          <div class="kpi"><div class="kpi-label">Plano de energia</div><div class="kpi-value" style="font-size:20px">${s.PowerPlan || 'N/D'}</div></div>
          <div class="kpi"><div class="kpi-label">Sistema</div><div class="kpi-value" style="font-size:20px">${s.OS || 'N/D'}</div><div class="kpi-sub">build ${s.BuildNumber || 'N/D'}</div></div>
          <div class="kpi"><div class="kpi-label">C: livre</div><div class="kpi-value">${drive0 ? fmt(drive0.FreeGB,2) : 'N/D'}<span style="font-size:14px;color:var(--muted)"> GB</span></div><div class="kpi-sub">${drive0 ? fmt(drive0.FreePct,2,'%') : 'N/D'}</div></div>
          <div class="kpi"><div class="kpi-label">Pagefile atual</div><div class="kpi-value">${fmt(m.PagefileCurrentMB,0)}<span style="font-size:14px;color:var(--muted)"> MB</span></div></div>
        </div>
      </div>

      <div class="section">
        <div class="section-head">
          <div><h3>Vilões de navegadores</h3><p>Quem realmente está comendo RAM privada e CPU acumulada.</p></div>
        </div>
        ${renderBrowsers(browserSummary)}
      </div>

      <div class="section">
        <div class="section-head">
          <div><h3>Processos dominantes</h3><p>Ordenado por CPU máxima e memória privada máxima. Agora dá pra bater o olho sem rezar.</p></div>
        </div>
        ${renderProcesses(topSummary)}
      </div>

      <div class="section">
        <div class="section-head">
          <div><h3>Comparativo entre runs</h3><p>Pra matar o achismo sem abrir CSV na unha.</p></div>
        </div>
        ${renderComparison(filteredRuns())}
      </div>

      <div class="footer">Viewer offline, auto-contido e sem dependência de CDN. Finalmente uma leitura que não parece um castigo administrativo.</div>
    `;
  }

  function renderAll(){
    renderSidebar();
    appContentEl.innerHTML = renderSelectedRun(getSelectedRun());
  }

  runSearchEl.addEventListener('input', e => {
    state.search = e.target.value || '';
    renderAll();
  });

  renderAll();
})();
</script>
</body>
</html>
'@

    return $template.Replace('__PAYLOAD_JSON__', $safeJson)
}

$root = Resolve-FullPathSafe -Path $RootPath

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path -Path $PSScriptRoot -ChildPath 'viewer-report.html'
}

$runFolders = @(Get-RunDirectories -Root $root)
if ($runFolders.Count -eq 0) {
    throw "Nenhuma pasta de run encontrada em '$root'."
}

$runs = @()
foreach ($folder in $runFolders) {
    $runs += @(New-RunModel -Folder $folder)
}

$runs = @($runs | Sort-Object Timestamp)
$payload = [PSCustomObject]@{
    GeneratedAt = (Get-Date).ToString('s')
    RootPath    = $root
    RunCount    = @($runs).Count
    Runs        = @($runs)
}

$json = $payload | ConvertTo-Json -Depth 20
$html = Get-EmbeddedHtmlTemplate -JsonPayload $json
$html | Set-Content -LiteralPath $OutputPath -Encoding UTF8

Write-Host ''
Write-Host 'Viewer gerado com sucesso.' -ForegroundColor Green
Write-Host ('RootPath : {0}' -f $root)
Write-Host ('Runs     : {0}' -f @($runs).Count)
Write-Host ('Output   : {0}' -f $OutputPath)

