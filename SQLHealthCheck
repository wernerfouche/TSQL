<#
.SYNOPSIS
    SQL Server health check across multiple instances. Produces an HTML report
    and can optionally email it to a distribution list. Designed to run unattended
    via a Windows Scheduled Task.

.DESCRIPTION
    Checks performed per instance (all gathered over the SQL connection only,
    so the service account just needs SQL access - no WinRM/remote admin required):
        * Disk space for volumes hosting SQL data/log files  (red < CriticalDiskPct)
        * SQL-related services set to Automatic startup that are not Running (red)
        * Last full backup age per database (red if older than threshold / never)
        * SQL Agent jobs that failed in the last N hours (always red)

.NOTES
    Tested on Windows PowerShell 5.1 (the version shipped on most Windows Server
    hosts and the typical Scheduled Task environment) and PowerShell 7+.
    Send-MailMessage is used for email - it is deprecated by Microsoft but remains
    fully functional for internal SMTP relays, which is the usual scenario here.

.EXAMPLE
    .\SqlHealthCheck.ps1 -SqlServers 'SQL01','SQL02\INST1' -SendEmail

.EXAMPLE
    # Scheduled-task friendly call with everything passed in:
    powershell.exe -ExecutionPolicy Bypass -File C:\Scripts\SqlHealthCheck.ps1 `
        -SqlServers 'SQL01','SQL02' -SendEmail `
        -SmtpServer 'smtp.contoso.com' -MailFrom 'sqlalerts@contoso.com' `
        -MailTo 'dba-team@contoso.com','oncall@contoso.com'
#>

[CmdletBinding()]
param(
    # ---- Config file ---------------------------------------------------------
    # JSON file that supplies the server list (and optionally thresholds/email).
    # Defaults to a file sitting next to this script. Any value passed explicitly
    # on the command line overrides the matching value in the config file.
    [string]$ConfigFile = "$PSScriptRoot\SqlHealthCheck.config.json",

    # ---- What to check -------------------------------------------------------
    [string[]]$SqlServers = @('localhost'),     # used only if not supplied by the config file

    [int]$CriticalDiskPct      = 10,            # free % at/under this = red
    [int]$WarnDiskPct          = 20,            # free % at/under this = amber
    [int]$FullBackupMaxHours   = 24,            # full backup older than this = red
    [int]$LogBackupMaxHours    = 4,             # log backup older than this = red (FULL/BULK_LOGGED dbs)
    [int]$FailedJobLookbackHrs = 24,            # window for failed agent jobs

    # ---- Connection (optional SQL auth; default is Windows/integrated) -------
    [string]$SqlUser,                           # leave blank to use the task's Windows identity
    [string]$SqlPassword,
    [int]$QueryTimeoutSec = 30,

    # ---- Output --------------------------------------------------------------
    [string]$OutputFolder = "$PSScriptRoot\Reports",

    # ---- Email ---------------------------------------------------------------
    [switch]$SendEmail,
    [string]$SmtpServer  = 'smtp.yourdomain.com',
    [int]$SmtpPort       = 25,
    [string]$MailFrom    = 'sqlhealthcheck@yourdomain.com',
    [string[]]$MailTo    = @('dba-distribution@yourdomain.com'),
    [switch]$SmtpUseSsl,
    [pscredential]$SmtpCredential                # optional, for authenticated relays
)

$ErrorActionPreference = 'Stop'

# ----------------------------------------------------------------------------
#  Load external config file (JSON)
#  - SqlServers is the primary value driven from the file.
#  - Thresholds / Email / OutputFolder may optionally be supplied too.
#  - Anything passed explicitly on the command line always wins.
# ----------------------------------------------------------------------------
if ($PSBoundParameters.ContainsKey('ConfigFile') -and -not (Test-Path $ConfigFile)) {
    throw "Config file not found: $ConfigFile"
}

if (Test-Path $ConfigFile) {
    try {
        $cfg = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
    }
    catch {
        throw "Failed to parse config file '$ConfigFile': $($_.Exception.Message)"
    }

    # --- Server list (the value this file primarily drives) ---
    if (-not $PSBoundParameters.ContainsKey('SqlServers') -and $cfg.SqlServers) {
        $SqlServers = @($cfg.SqlServers)
    }

    # --- Optional threshold overrides ---
    if ($cfg.Thresholds) {
        if (-not $PSBoundParameters.ContainsKey('CriticalDiskPct')      -and $null -ne $cfg.Thresholds.CriticalDiskPct)      { $CriticalDiskPct      = [int]$cfg.Thresholds.CriticalDiskPct }
        if (-not $PSBoundParameters.ContainsKey('WarnDiskPct')          -and $null -ne $cfg.Thresholds.WarnDiskPct)          { $WarnDiskPct          = [int]$cfg.Thresholds.WarnDiskPct }
        if (-not $PSBoundParameters.ContainsKey('FullBackupMaxHours')   -and $null -ne $cfg.Thresholds.FullBackupMaxHours)   { $FullBackupMaxHours   = [int]$cfg.Thresholds.FullBackupMaxHours }
        if (-not $PSBoundParameters.ContainsKey('LogBackupMaxHours')    -and $null -ne $cfg.Thresholds.LogBackupMaxHours)    { $LogBackupMaxHours    = [int]$cfg.Thresholds.LogBackupMaxHours }
        if (-not $PSBoundParameters.ContainsKey('FailedJobLookbackHrs') -and $null -ne $cfg.Thresholds.FailedJobLookbackHrs) { $FailedJobLookbackHrs = [int]$cfg.Thresholds.FailedJobLookbackHrs }
    }

    # --- Optional email overrides ---
    if ($cfg.Email) {
        if (-not $PSBoundParameters.ContainsKey('SendEmail')  -and $null -ne $cfg.Email.SendEmail  -and [bool]$cfg.Email.SendEmail)  { $SendEmail  = $true }
        if (-not $PSBoundParameters.ContainsKey('SmtpUseSsl') -and $null -ne $cfg.Email.SmtpUseSsl -and [bool]$cfg.Email.SmtpUseSsl) { $SmtpUseSsl = $true }
        if (-not $PSBoundParameters.ContainsKey('SmtpServer') -and $cfg.Email.SmtpServer) { $SmtpServer = [string]$cfg.Email.SmtpServer }
        if (-not $PSBoundParameters.ContainsKey('SmtpPort')   -and $null -ne $cfg.Email.SmtpPort) { $SmtpPort = [int]$cfg.Email.SmtpPort }
        if (-not $PSBoundParameters.ContainsKey('MailFrom')   -and $cfg.Email.MailFrom) { $MailFrom = [string]$cfg.Email.MailFrom }
        if (-not $PSBoundParameters.ContainsKey('MailTo')     -and $cfg.Email.MailTo)   { $MailTo   = @($cfg.Email.MailTo) }
    }

    # --- Optional output folder override ---
    if (-not $PSBoundParameters.ContainsKey('OutputFolder') -and $cfg.OutputFolder) {
        $OutputFolder = [string]$cfg.OutputFolder
    }
}

if (-not $SqlServers -or $SqlServers.Count -eq 0) {
    throw "No SQL servers to check. Supply them via -SqlServers or the 'SqlServers' array in the config file."
}

# ----------------------------------------------------------------------------
#  Core: run a query and return rows as objects
# ----------------------------------------------------------------------------
function Invoke-SqlQuery {
    param(
        [Parameter(Mandatory)] [string]$ServerInstance,
        [string]$Database = 'master',
        [Parameter(Mandatory)] [string]$Query
    )

    if ($SqlUser) {
        $auth = "User ID=$SqlUser;Password=$SqlPassword"
    } else {
        $auth = 'Integrated Security=SSPI'
    }
    $connString = "Server=$ServerInstance;Database=$Database;$auth;Application Name=SqlHealthCheck;Connect Timeout=15"

    $conn = New-Object System.Data.SqlClient.SqlConnection $connString
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $QueryTimeoutSec
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $table   = New-Object System.Data.DataTable
        [void]$adapter.Fill($table)
        return $table.Rows
    }
    finally {
        $conn.Close()
        $conn.Dispose()
    }
}

# ----------------------------------------------------------------------------
#  SQL queries (gathered over the connection only)
# ----------------------------------------------------------------------------
$Q_Disk = @"
SELECT DISTINCT
    vs.volume_mount_point                                              AS Drive,
    CAST(vs.total_bytes / 1073741824.0 AS DECIMAL(18,2))              AS TotalGB,
    CAST(vs.available_bytes / 1073741824.0 AS DECIMAL(18,2))          AS FreeGB,
    CAST(vs.available_bytes * 100.0 / NULLIF(vs.total_bytes,0) AS DECIMAL(5,2)) AS FreePct
FROM sys.master_files AS f
CROSS APPLY sys.dm_os_volume_stats(f.database_id, f.file_id) AS vs
ORDER BY Drive;
"@

$Q_Services = @"
SELECT
    servicename            AS ServiceName,
    startup_type_desc      AS StartupType,
    status_desc            AS Status,
    service_account        AS ServiceAccount
FROM sys.dm_server_services;
"@

$Q_Backups = @"
SELECT
    d.name                                                       AS DatabaseName,
    d.recovery_model_desc                                        AS RecoveryModel,
    MAX(CASE WHEN b.type = 'D' THEN b.backup_finish_date END)    AS LastFullBackup,
    MAX(CASE WHEN b.type = 'L' THEN b.backup_finish_date END)    AS LastLogBackup
FROM sys.databases d
LEFT JOIN msdb.dbo.backupset b ON b.database_name = d.name
WHERE d.name <> 'tempdb'
  AND d.state_desc = 'ONLINE'
GROUP BY d.name, d.recovery_model_desc
ORDER BY d.name;
"@

$Q_FailedJobs = @"
SELECT
    j.name                                              AS JobName,
    msdb.dbo.agent_datetime(h.run_date, h.run_time)     AS RunDateTime,
    h.run_duration                                      AS DurationHHMMSS,
    h.message                                           AS Message
FROM msdb.dbo.sysjobhistory h
JOIN msdb.dbo.sysjobs j ON j.job_id = h.job_id
WHERE h.step_id   = 0          -- job outcome row
  AND h.run_status = 0         -- 0 = Failed
  AND msdb.dbo.agent_datetime(h.run_date, h.run_time) >= DATEADD(HOUR, -$FailedJobLookbackHrs, GETDATE())
ORDER BY RunDateTime DESC;
"@

# ----------------------------------------------------------------------------
#  Small HTML helpers
# ----------------------------------------------------------------------------
function HtmlEncode([object]$v) {
    if ($null -eq $v -or $v -is [DBNull]) { return '' }
    [System.Web.HttpUtility]::HtmlEncode([string]$v)
}
# Load System.Web for HtmlEncode (present on Windows; fall back to a simple replace)
try { Add-Type -AssemblyName System.Web -ErrorAction Stop } catch {
    function HtmlEncode([object]$v) {
        if ($null -eq $v -or $v -is [DBNull]) { return '' }
        ([string]$v) -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
    }
}

function Cell([string]$text, [string]$class = '') {
    if ($class) { "<td class='$class'>$text</td>" } else { "<td>$text</td>" }
}

# ----------------------------------------------------------------------------
#  Collect data per server and build report sections
# ----------------------------------------------------------------------------
$sb = [System.Text.StringBuilder]::new()
$now = Get-Date
$grandTotalIssues = 0

foreach ($server in $SqlServers) {

    [void]$sb.AppendLine("<h2>$([System.Net.WebUtility]::HtmlEncode($server))</h2>")
    $serverIssues = 0

    # -- Connectivity guard: don't let one dead server kill the whole report --
    try {
        $disk     = Invoke-SqlQuery -ServerInstance $server -Query $Q_Disk
        $services = Invoke-SqlQuery -ServerInstance $server -Query $Q_Services
        $backups  = Invoke-SqlQuery -ServerInstance $server -Query $Q_Backups
        $jobs     = Invoke-SqlQuery -ServerInstance $server -Query $Q_FailedJobs
    }
    catch {
        [void]$sb.AppendLine("<p class='unreachable'>UNREACHABLE - $([System.Net.WebUtility]::HtmlEncode($_.Exception.Message))</p>")
        $grandTotalIssues++
        continue
    }

    # ---------------- Disk space -----------------
    [void]$sb.AppendLine("<h3>Disk Space (volumes hosting SQL files)</h3>")
    [void]$sb.AppendLine("<table><tr><th>Drive</th><th>Total (GB)</th><th>Free (GB)</th><th>Free %</th><th>Status</th></tr>")
    foreach ($d in $disk) {
        $pct = [double]$d.FreePct
        if     ($pct -le $CriticalDiskPct) { $cls = 'crit'; $status = 'CRITICAL'; $serverIssues++ }
        elseif ($pct -le $WarnDiskPct)     { $cls = 'warn'; $status = 'Low' }
        else                               { $cls = 'ok';   $status = 'OK' }
        [void]$sb.AppendLine("<tr>" +
            (Cell (HtmlEncode $d.Drive)) +
            (Cell ('{0:N1}' -f [double]$d.TotalGB)) +
            (Cell ('{0:N1}' -f [double]$d.FreeGB)) +
            (Cell ('{0:N1}%' -f $pct) $cls) +
            (Cell $status $cls) + "</tr>")
    }
    [void]$sb.AppendLine("</table>")

    # ---------------- Services (Automatic but not Running) -----------------
    [void]$sb.AppendLine("<h3>SQL Services (Automatic startup)</h3>")
    [void]$sb.AppendLine("<table><tr><th>Service</th><th>Startup Type</th><th>Account</th><th>Status</th></tr>")
    $autoServices = $services | Where-Object { $_.StartupType -eq 'Automatic' }
    foreach ($s in $autoServices) {
        if ($s.Status -ne 'Running') { $cls = 'crit'; $serverIssues++ } else { $cls = 'ok' }
        [void]$sb.AppendLine("<tr>" +
            (Cell (HtmlEncode $s.ServiceName)) +
            (Cell (HtmlEncode $s.StartupType)) +
            (Cell (HtmlEncode $s.ServiceAccount)) +
            (Cell (HtmlEncode $s.Status) $cls) + "</tr>")
    }
    [void]$sb.AppendLine("</table>")

    # ---------------- Backups -----------------
    [void]$sb.AppendLine("<h3>Last Backups</h3>")
    [void]$sb.AppendLine("<table><tr><th>Database</th><th>Recovery Model</th><th>Last Full Backup</th><th>Full Age (hrs)</th><th>Last Log Backup</th><th>Status</th></tr>")
    foreach ($b in $backups) {
        $issuesForRow = @()

        # Full backup check
        if ($b.LastFullBackup -is [DBNull] -or $null -eq $b.LastFullBackup) {
            $fullAgeStr = 'NEVER'; $fullCls = 'crit'; $issuesForRow += 'no full backup'
        } else {
            $ageHrs = [math]::Round((New-TimeSpan -Start ([datetime]$b.LastFullBackup) -End $now).TotalHours, 1)
            $fullAgeStr = $ageHrs
            if ($ageHrs -gt $FullBackupMaxHours) { $fullCls = 'crit'; $issuesForRow += 'full backup stale' }
            else { $fullCls = 'ok' }
        }

        # Log backup check - only meaningful for FULL / BULK_LOGGED recovery
        $logCls = 'ok'
        if ($b.RecoveryModel -ne 'SIMPLE') {
            if ($b.LastLogBackup -is [DBNull] -or $null -eq $b.LastLogBackup) {
                $logCls = 'crit'; $issuesForRow += 'no log backup'
            } else {
                $logAge = (New-TimeSpan -Start ([datetime]$b.LastLogBackup) -End $now).TotalHours
                if ($logAge -gt $LogBackupMaxHours) { $logCls = 'crit'; $issuesForRow += 'log backup stale' }
            }
        }

        if ($issuesForRow.Count) { $serverIssues++; $rowStatus = ($issuesForRow -join ', '); $statusCls = 'crit' }
        else { $rowStatus = 'OK'; $statusCls = 'ok' }

        $fullStr = if ($b.LastFullBackup -is [DBNull]) { '' } else { ([datetime]$b.LastFullBackup).ToString('yyyy-MM-dd HH:mm') }
        $logStr  = if ($b.LastLogBackup  -is [DBNull]) { '' } else { ([datetime]$b.LastLogBackup ).ToString('yyyy-MM-dd HH:mm') }

        [void]$sb.AppendLine("<tr>" +
            (Cell (HtmlEncode $b.DatabaseName)) +
            (Cell (HtmlEncode $b.RecoveryModel)) +
            (Cell $fullStr $fullCls) +
            (Cell $fullAgeStr $fullCls) +
            (Cell $logStr $logCls) +
            (Cell (HtmlEncode $rowStatus) $statusCls) + "</tr>")
    }
    [void]$sb.AppendLine("</table>")

    # ---------------- Failed Agent Jobs -----------------
    [void]$sb.AppendLine("<h3>Failed Agent Jobs (last $FailedJobLookbackHrs hrs)</h3>")
    if ($jobs.Count -eq 0) {
        [void]$sb.AppendLine("<table><tr><td class='ok'>No failed jobs in the lookback window.</td></tr></table>")
    } else {
        $serverIssues += $jobs.Count
        [void]$sb.AppendLine("<table><tr><th>Job</th><th>Run Time</th><th>Message</th></tr>")
        foreach ($j in $jobs) {
            $rt = if ($j.RunDateTime -is [DBNull]) { '' } else { ([datetime]$j.RunDateTime).ToString('yyyy-MM-dd HH:mm') }
            [void]$sb.AppendLine("<tr>" +
                (Cell (HtmlEncode $j.JobName) 'crit') +
                (Cell $rt 'crit') +
                (Cell (HtmlEncode $j.Message) 'crit') + "</tr>")
        }
        [void]$sb.AppendLine("</table>")
    }

    $grandTotalIssues += $serverIssues
}

# ----------------------------------------------------------------------------
#  Assemble full HTML document
# ----------------------------------------------------------------------------
$headerBanner = if ($grandTotalIssues -gt 0) {
    "<div class='banner crit'>$grandTotalIssues issue(s) detected - review highlighted items.</div>"
} else {
    "<div class='banner ok'>All checks passed.</div>"
}

$css = @"
body   { font-family: Segoe UI, Arial, sans-serif; font-size: 13px; color: #222; margin: 20px; }
h1     { font-size: 20px; margin-bottom: 2px; }
h2     { font-size: 16px; margin-top: 28px; border-bottom: 2px solid #ccc; padding-bottom: 4px; }
h3     { font-size: 14px; margin-top: 16px; color: #444; }
table  { border-collapse: collapse; width: 100%; margin-bottom: 8px; }
th     { background: #2f5496; color: #fff; text-align: left; padding: 6px 8px; font-weight: 600; }
td     { border: 1px solid #ddd; padding: 5px 8px; }
tr:nth-child(even) td { background: #f7f7f7; }
.ok    { background: #d7f0d7 !important; color: #145214; }
.warn  { background: #fff2cc !important; color: #7f6000; }
.crit  { background: #f8c9c4 !important; color: #8b0000; font-weight: 600; }
.unreachable { color: #8b0000; font-weight: 600; }
.banner { padding: 10px 14px; border-radius: 4px; font-weight: 600; margin: 10px 0; }
.meta  { color: #666; font-size: 12px; }
"@

$html = @"
<!DOCTYPE html>
<html><head><meta charset='utf-8'><title>SQL Health Check</title><style>$css</style></head>
<body>
<h1>SQL Server Health Check</h1>
<p class='meta'>Generated $($now.ToString('yyyy-MM-dd HH:mm:ss')) on $env:COMPUTERNAME &middot; $($SqlServers.Count) instance(s)</p>
$headerBanner
$($sb.ToString())
</body></html>
"@

# ----------------------------------------------------------------------------
#  Write the report to disk
# ----------------------------------------------------------------------------
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$reportPath = Join-Path $OutputFolder ("SqlHealthCheck_{0:yyyyMMdd_HHmmss}.html" -f $now)
$html | Out-File -FilePath $reportPath -Encoding utf8
Write-Host "Report written to $reportPath"

# ----------------------------------------------------------------------------
#  Email (optional)
# ----------------------------------------------------------------------------
if ($SendEmail) {
    $subject = if ($grandTotalIssues -gt 0) {
        "[ACTION] SQL Health Check - $grandTotalIssues issue(s) - $($now.ToString('yyyy-MM-dd'))"
    } else {
        "SQL Health Check - All OK - $($now.ToString('yyyy-MM-dd'))"
    }

    $mailParams = @{
        From       = $MailFrom
        To         = $MailTo
        Subject    = $subject
        Body       = $html
        BodyAsHtml = $true
        SmtpServer = $SmtpServer
        Port       = $SmtpPort
    }
    if ($SmtpUseSsl)        { $mailParams.UseSsl = $true }
    if ($SmtpCredential)    { $mailParams.Credential = $SmtpCredential }

    try {
        Send-MailMessage @mailParams
        Write-Host "Report emailed to: $($MailTo -join ', ')"
    }
    catch {
        Write-Warning "Email send failed: $($_.Exception.Message)"
    }
}
