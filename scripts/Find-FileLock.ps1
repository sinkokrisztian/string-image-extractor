param(
    [Parameter(Mandatory = $true)]
    [string] $Path,

    [switch] $TryServerSmbCheck
)

$ErrorActionPreference = "Stop"

function Test-ExclusiveOpen {
    param([string] $LiteralPath)

    try {
        $stream = [System.IO.File]::Open(
            $LiteralPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $stream.Close()
        return $true
    }
    catch {
        Write-Host "Exclusive open failed:" -ForegroundColor Yellow
        Write-Host "  $($_.Exception.Message)"
        return $false
    }
}

function Get-LocalLockingProcess {
    param([string] $LiteralPath)

    if (-not ("RestartManager" -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class RestartManager {
    [StructLayout(LayoutKind.Sequential)]
    public struct RM_UNIQUE_PROCESS {
        public int dwProcessId;
        public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;
    }

    public const int RmRebootReasonNone = 0;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct RM_PROCESS_INFO {
        public RM_UNIQUE_PROCESS Process;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string strAppName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string strServiceShortName;
        public int ApplicationType;
        public uint AppStatus;
        public uint TSSessionId;
        [MarshalAs(UnmanagedType.Bool)]
        public bool bRestartable;
    }

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    public static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, StringBuilder strSessionKey);

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    public static extern int RmRegisterResources(uint pSessionHandle, uint nFiles, string[] rgsFilenames, uint nApplications, IntPtr rgApplications, uint nServices, string[] rgsServiceNames);

    [DllImport("rstrtmgr.dll")]
    public static extern int RmGetList(uint dwSessionHandle, out uint pnProcInfoNeeded, ref uint pnProcInfo, [In, Out] RM_PROCESS_INFO[] rgAffectedApps, ref uint lpdwRebootReasons);

    [DllImport("rstrtmgr.dll")]
    public static extern int RmEndSession(uint pSessionHandle);
}
"@
    }

    $sessionHandle = 0
    $sessionKey = New-Object System.Text.StringBuilder 64
    $result = [RestartManager]::RmStartSession([ref] $sessionHandle, 0, $sessionKey)
    if ($result -ne 0) {
        throw "RmStartSession failed with code $result"
    }

    try {
        $files = @($LiteralPath)
        $result = [RestartManager]::RmRegisterResources($sessionHandle, 1, $files, 0, [IntPtr]::Zero, 0, $null)
        if ($result -ne 0) {
            throw "RmRegisterResources failed with code $result"
        }

        $needed = 0
        $count = 0
        $reason = 0
        $null = [RestartManager]::RmGetList($sessionHandle, [ref] $needed, [ref] $count, $null, [ref] $reason)
        if ($needed -eq 0) {
            return @()
        }

        $count = $needed
        $processes = New-Object RestartManager+RM_PROCESS_INFO[] $count
        $result = [RestartManager]::RmGetList($sessionHandle, [ref] $needed, [ref] $count, $processes, [ref] $reason)
        if ($result -ne 0) {
            throw "RmGetList failed with code $result"
        }

        $out = @()
        for ($i = 0; $i -lt $count; $i++) {
            $pid = $processes[$i].Process.dwProcessId
            $proc = Get-Process -Id $pid -ErrorAction SilentlyContinue
            $out += [pscustomobject]@{
                Id = $pid
                Name = if ($proc) { $proc.ProcessName } else { $processes[$i].strAppName }
                AppName = $processes[$i].strAppName
                MainWindowTitle = if ($proc) { $proc.MainWindowTitle } else { "" }
            }
        }
        return $out
    }
    finally {
        [void][RestartManager]::RmEndSession($sessionHandle)
    }
}

$resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
Write-Host "Checking:" -ForegroundColor Cyan
Write-Host "  $resolved"

if (-not (Test-Path -LiteralPath $resolved)) {
    Write-Host "File does not exist, or this user cannot see it." -ForegroundColor Red
    exit 2
}

$item = Get-Item -LiteralPath $resolved -Force
Write-Host ""
Write-Host "File attributes:" -ForegroundColor Cyan
Write-Host "  $($item.Attributes)"

Write-Host ""
Write-Host "Local lock test:" -ForegroundColor Cyan
$canOpen = Test-ExclusiveOpen -LiteralPath $resolved
if ($canOpen) {
    Write-Host "  No exclusive-lock blocker detected from this machine." -ForegroundColor Green
    Write-Host "  If deletion still fails, check permissions, read-only attributes, sync/AV tooling, or a server-side SMB handle."
}

Write-Host ""
Write-Host "Local processes reported by Restart Manager:" -ForegroundColor Cyan
$localLocks = Get-LocalLockingProcess -LiteralPath $resolved
if ($localLocks.Count -eq 0) {
    Write-Host "  None found from this machine."
}
else {
    $localLocks | Format-Table -AutoSize
}

if ($TryServerSmbCheck) {
    $uri = [Uri]$resolved
    if (-not $uri.IsUnc) {
        Write-Host ""
        Write-Host "Server SMB check skipped: path is not UNC." -ForegroundColor Yellow
        exit 0
    }

    $server = $uri.Host
    $leaf = Split-Path -Leaf $resolved
    Write-Host ""
    Write-Host "Server-side SMB open-file check on ${server}:" -ForegroundColor Cyan
    Write-Host "  This requires admin rights / PowerShell remoting permissions on the file server."

    try {
        Invoke-Command -ComputerName $server -ScriptBlock {
            param($Leaf)
            Get-SmbOpenFile |
                Where-Object { $_.Path -like "*$Leaf*" } |
                Select-Object FileId, SessionId, ClientComputerName, ClientUserName, Path
        } -ArgumentList $leaf | Format-Table -AutoSize
    }
    catch {
        Write-Host "  Server-side SMB check failed:" -ForegroundColor Yellow
        Write-Host "  $($_.Exception.Message)"
        Write-Host ""
        Write-Host "Ask a file-server admin to run this on ${server}:" -ForegroundColor Cyan
        Write-Host "  `$leaf = '$leaf'"
        Write-Host "  Get-SmbOpenFile | Where-Object { `$_.Path -like `"*`$leaf*`" } | Select-Object FileId,SessionId,ClientComputerName,ClientUserName,Path"
        Write-Host ""
        Write-Host "After confirming the stale handle, they can close only that file handle with:" -ForegroundColor Cyan
        Write-Host "  Close-SmbOpenFile -FileId <FileId> -Force"
    }
}
