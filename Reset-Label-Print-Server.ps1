param(
    [int]$Port = $env:ZPL_SERVER_PORT,
    [string]$HostUrl = "http://127.0.0.1",
    [switch]$NoPause
)

$ErrorActionPreference = "Stop"

if (-not $Port) {
    $Port = 8787
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$MainPath = Join-Path $ScriptDir "main.py"
$LogDir = Join-Path $ScriptDir "logs"
$OutLog = Join-Path $LogDir "label-print-server.out.log"
$ErrLog = Join-Path $LogDir "label-print-server.err.log"
$HealthUrl = "${HostUrl}:$Port/health"

function Write-Step($Message) {
    Write-Host ""
    Write-Host "== $Message ==" -ForegroundColor Cyan
}

function Write-Good($Message) {
    Write-Host $Message -ForegroundColor Green
}

function Write-Bad($Message) {
    Write-Host $Message -ForegroundColor Red
}

function Write-Warn($Message) {
    Write-Host $Message -ForegroundColor Yellow
}

function Wait-Before-Exit {
    if (-not $NoPause) {
        Write-Host ""
        Read-Host "Press Enter to close this window"
    }
}

function Get-CommandLine($ProcessId) {
    $proc = Get-CimInstance Win32_Process -Filter "ProcessId = $ProcessId" -ErrorAction SilentlyContinue
    if ($proc) {
        return [string]$proc.CommandLine
    }
    return ""
}

function Find-ServerProcess {
    $connections = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
    foreach ($connection in $connections) {
        if (-not $connection.OwningProcess) {
            continue
        }

        $process = Get-Process -Id $connection.OwningProcess -ErrorAction SilentlyContinue
        if (-not $process) {
            continue
        }

        $commandLine = Get-CommandLine $process.Id
        $isPython = $process.ProcessName -like "python*" -or $process.ProcessName -eq "py"
        $usesAbsoluteMain = $commandLine.IndexOf($MainPath, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        $usesRelativeMain = $commandLine.IndexOf("main.py", [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        $isThisServer = $isPython -and ($usesAbsoluteMain -or $usesRelativeMain)
        [pscustomobject]@{
            Process = $process
            CommandLine = $commandLine
            IsThisServer = $isThisServer
        }
    }
}

try {
    Write-Host "Zebra Label Print Server Reset" -ForegroundColor White
    Write-Host "Folder: $ScriptDir"
    Write-Host "Port: $Port"

    if (-not (Test-Path -LiteralPath $MainPath)) {
        throw "Could not find main.py in $ScriptDir"
    }

    if (-not (Test-Path -LiteralPath $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir | Out-Null
    }

    Write-Step "Checking for the current server"
    $owners = @(Find-ServerProcess)
    $otherOwners = @($owners | Where-Object { -not $_.IsThisServer })
    $serverOwners = @($owners | Where-Object { $_.IsThisServer })

    if ($otherOwners.Count -gt 0) {
        Write-Bad "Port $Port is being used by another program. I will not stop it automatically."
        foreach ($owner in $otherOwners) {
            Write-Host "Process: $($owner.Process.ProcessName)  PID: $($owner.Process.Id)"
            if ($owner.CommandLine) {
                Write-Host "Command: $($owner.CommandLine)"
            }
        }
        Write-Warn "Ask someone technical to check what is using port $Port."
        Wait-Before-Exit
        exit 2
    }

    if ($serverOwners.Count -eq 0) {
        Write-Warn "No running label print server was found. Starting a fresh one."
    } else {
        foreach ($owner in $serverOwners) {
            Write-Host "Stopping old server PID $($owner.Process.Id)..."
            Stop-Process -Id $owner.Process.Id -Force
        }
        Start-Sleep -Seconds 2
    }

    Write-Step "Starting the label print server"
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if (-not $pythonCommand) {
        $pythonCommand = Get-Command py -ErrorAction SilentlyContinue
    }
    if (-not $pythonCommand) {
        throw "Could not find Python. Make sure Python is installed and available from the Start menu/terminal."
    }

    $startInfo = @{
        FilePath = $pythonCommand.Source
        ArgumentList = @("`"$MainPath`"")
        WorkingDirectory = $ScriptDir
        RedirectStandardOutput = $OutLog
        RedirectStandardError = $ErrLog
        WindowStyle = "Hidden"
        PassThru = $true
    }
    $newProcess = Start-Process @startInfo
    Write-Host "Started server process PID $($newProcess.Id)."

    Write-Step "Waiting for the server to answer"
    $healthy = $false
    for ($i = 1; $i -le 20; $i++) {
        Start-Sleep -Seconds 1
        try {
            $health = Invoke-RestMethod -Uri $HealthUrl -TimeoutSec 2
            if ($health.ok) {
                $healthy = $true
                break
            }
        } catch {
            Write-Host "Waiting... ($i/20)"
        }
    }

    if (-not $healthy) {
        Write-Bad "The server was started, but it did not answer the health check."
        Write-Host "Health URL: $HealthUrl"
        Write-Host "Output log: $OutLog"
        Write-Host "Error log: $ErrLog"
        Wait-Before-Exit
        exit 3
    }

    Write-Good "Server restarted successfully."
    Write-Host "Open or refresh: ${HostUrl}:$Port"
    Write-Host "If using Tailscale from home, open the office PC's Tailscale URL instead."
    Wait-Before-Exit
    exit 0
} catch {
    Write-Bad "Reset failed: $($_.Exception.Message)"
    Write-Host "Output log: $OutLog"
    Write-Host "Error log: $ErrLog"
    Wait-Before-Exit
    exit 1
}
