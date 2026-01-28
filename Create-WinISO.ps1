<#
Module: Create-WinISO

.SYNOPSIS
    Build a customized Windows Pro ISO with injected drivers.

.DESCRIPTION
    Copies a user-selected ISO into StockISO, extracts it, automatically detects
    the Windows Pro index, injects drivers, and rebuilds a BIOS+UEFI bootable ISO.

.AUTHOR
    Tyler Cox

.PARAMETERS
    -BuildRoot           : (string) Override build root directory. Default: C:\Build
    -Drivers             : (string) Override drivers folder path. Default: C:\Build\Drivers
    -IsoPath             : (string) Path to an existing Windows ISO (used in non-interactive mode)
    -OscdimgPath         : (string) Explicit path to oscdimg.exe (ADK)
    -NonInteractive      : (switch) Run without interactive prompts (fail fast on missing inputs)
    -DryRun              : (switch) Show actions without performing changes
    -AutoInstallAdk      : (switch) (future) Attempt unattended ADK install if missing

.EXAMPLES
    Interactive (menu):
        .\Create-WinISO.ps1

    Non-interactive (CI):
        .\Create-WinISO.ps1 -NonInteractive -IsoPath 'C:\isos\win11.iso' -OscdimgPath 'C:\Tools\oscdimg.exe'

    Override build root and drivers folder:
        pwsh -File .\Create-WinISO.ps1 -BuildRoot 'C:\Build' -Drivers 'C:\Build\Drivers'


.CHANGELOG
    v1.0.0 - Initial release 2026-01-15
        - Base ISO extraction and rebuild logic
        - Driver injection support
        - Bootable ISO creation (BIOS + UEFI)

    v1.1.0 2026-01-16
        - Replaced all working paths with C:\Build
        - Added user prompt for Windows ISO path
        - Automatic cleanup and folder preparation
    v1.2.0 2026-01-19
        - Automatic Windows image index detection
        - Explicit selection of Windows Pro edition

    v1.2.1 2026-01-19
        - Explicit oscdimg.exe path resolution (ADK)
        - Clearer error messaging during ISO creation

    v2.0.0 2026-01-20
        - Added menu system for Automatic, Download ISO, Select Drivers, and Build options
        - Automatic mode skips ISO download if one already exists in StockISO
        - Optimization: Skip ISO extraction if the ISO in StockISO has not changed since last build
        - Added interactive driver selection with categories (Wifi, Ethernet, Chipset)
    v2.1.0 2026-01-21
        - Copy driver folders by category
        - Ask user if they have more drivers to add
        - Check for mounted WIMs and unmount before mounting
    v2.1.1 2026-01-22
        - Added function to ensure oscdimg.exe is available
        - User prompt to install ADK if missing

    v2.2 2026-01-22
        - Use PowerShell to download Windows ADK with browser fallback
        - Improved non-interactive handling for ADK install prompt
        - Added silent ADK install option using `/quiet /norestart /ceip off` via `-AutoInstallAdk` or interactive selection
        - Make Silent the default ADK installer option and auto-select after 10 seconds of no input
        - Show a progress bar during silent ADK installation
        - Silent ADK install now limits components to Deployment Tools and Imaging and Configuration Designer

    v2.3.0 - 2026-01-23
        - Reuse existing `adksetup.exe` if present instead of deleting it
        - Retry silent ADK installs on exit code 1001 (rate limit) with exponential backoff and warn user
        - Remove STDERR capture/logging from silent installer helper
        - Download and run `DL-WinISO.ps1` locally (or use remote raw script) and handle rate-limit responses
        - If DL-WinISO fails, prompt user to provide an ISO (support Media Creation Tool flow)
        - Download Microsoft Media Creation Tool, wait 3s before download, and launch it elevated
        - Various robustness and UX improvements around downloads and prompts
#>

Param(
    [string]$BuildRoot,
    [string]$Drivers,
    [string]$IsoPath,
    [string]$OscdimgPath,
    [switch]$NonInteractive,
    [switch]$DryRun,
    [switch]$AutoInstallAdk,
    [string[]]$AdkFeatures = @('Deployment Tools','Imaging and Configuration Designer'),
    [switch]$Help
)

if ($Help) {
    Write-Host "Create-WinISO.ps1 - Parameters:" -ForegroundColor Cyan
    Write-Host "  -BuildRoot        : string    Build root directory (default C:\Build)"
    Write-Host "  -Drivers          : string    Drivers folder path"
    Write-Host "  -IsoPath          : string    Path to an existing Windows ISO (non-interactive)"
    Write-Host "  -OscdimgPath      : string    Explicit path to oscdimg.exe (ADK)"
    Write-Host "  -NonInteractive   : switch    Run without interactive prompts"
    Write-Host "  -DryRun           : switch    Show actions without performing changes"
    Write-Host "  -AutoInstallAdk   : switch    Attempt unattended ADK install if missing"
    Write-Host "Examples:" -ForegroundColor Cyan
    Write-Host "  .\Create-WinISO.ps1 -Help"
    return
}

# ---------------- Script Version Auto-Increment ----------------
$ScriptVersion = [version]2.2.0
$NewVersion    = [version]::new($ScriptVersion.Major, $ScriptVersion.Minor, $ScriptVersion.Build + 1)
Write-Host "Running script version $NewVersion"

# ---------------- Ensure Running as Administrator ----------------
function Ensure-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal   = New-Object System.Security.Principal.WindowsPrincipal($currentUser)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")) {
        Write-Warning "This script must be run as Administrator. Attempting to relaunch as Admin!... "
        Write-Host "Relaunching in 5 seconds..."
        Start-Sleep -Seconds 5

        $wtPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"
        if (-not (Test-Path $wtPath)) {
            Write-Warning "Windows Terminal not found. Falling back to PowerShell window."
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
            $psi.Verb = "runas"
            $psi.UseShellExecute = $true
            $psi.WindowStyle = 'Normal'
            [System.Diagnostics.Process]::Start($psi) | Out-Null
        } else {
            $arguments = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
            Start-Process -FilePath $wtPath -ArgumentList "new-tab -p `"Windows PowerShell`" -- $arguments" -Verb RunAs
        }
        exit 1
    }
}
Ensure-Admin

# Helper: run a silent installer and show a progress bar while it runs
function Run-SilentInstallerWithProgress {
    param(
        [Parameter(Mandatory)][string]$InstallerPath,
        [string[]]$FeatureList,
        [int]$EstimatedSeconds = 900
    )

    Write-Host "Starting silent install: $InstallerPath"
    $argList = '/quiet /norestart /ceip off'
    if ($FeatureList -and $FeatureList.Count -gt 0) {
        $argList += " /features `"$([string]::Join(',', $FeatureList))`""
    }

    # Prepare log files
    $logDir = Join-Path $env:TEMP "adk_install_logs"
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
    $stdoutLog = Join-Path $logDir "adk_stdout_$(Get-Date -Format yyyyMMdd-HHmmss).log"

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $InstallerPath
    $psi.Arguments = $argList
    $psi.RedirectStandardOutput = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.Start() | Out-Null

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $outputBuilder = New-Object System.Text.StringBuilder

    while (-not $proc.HasExited) {
        if ($proc.StandardOutput.Peek() -ne -1) {
            $s = $proc.StandardOutput.ReadToEnd()
            if ($s) { $outputBuilder.Append($s) | Out-Null }
        }
        
        $elapsed = [int]$sw.Elapsed.TotalSeconds
        $percent = [int](($elapsed / $EstimatedSeconds) * 95)
        if ($percent -gt 95) { $percent = 95 }
        if ($percent -lt 0) { $percent = 0 }
        Write-Progress -Activity 'Installing Windows ADK (silent)' -Status "Elapsed: ${elapsed}s" -PercentComplete $percent
        Start-Sleep -Seconds 1
    }

    # Read remaining output
    try { $outputBuilder.Append($proc.StandardOutput.ReadToEnd()) | Out-Null } catch {}

    $sw.Stop()
    Write-Progress -Activity 'Installing Windows ADK (silent)' -Completed -Status "Completed in $([int]$sw.Elapsed.TotalSeconds)s"

    # Write logs
    $outputBuilder.ToString() | Out-File -FilePath $stdoutLog -Encoding utf8
    $exit = 0
    try { $exit = $proc.ExitCode } catch { $exit = -1 }

    Write-Host "Installer exit code: $exit"
    Write-Host "Stdout log: $stdoutLog"
    return @{ ExitCode = $exit; StdOut = $stdoutLog }
}

# ---------------- Paths ----------------
# Allow overriding via parameters; fall back to defaults
$BuildRoot  = if ($BuildRoot) { $BuildRoot } else { "C:\Build" }
$StockISO   = Join-Path $BuildRoot "StockISO"
$TempISO    = Join-Path $BuildRoot "WindowsISO"
$WimMount   = Join-Path $BuildRoot "wim"
$CustomISO  = Join-Path $BuildRoot "CustomISO"
$Drivers    = if ($Drivers) { $Drivers } else { Join-Path $BuildRoot "Drivers" }

$ISOOut     = Join-Path $CustomISO "Windows11_VACO.iso"
$Oscdimg    = if ($OscdimgPath) { $OscdimgPath } else { "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe" }
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$DLWinISOScript = Join-Path $ScriptDir "DL-WinISO.ps1"

# Ensure oscdimg.exe is available early (before asking for drivers)
function Ensure-Oscdimg {
    param([string]$PreferredPath, [switch]$NonInteractive)

    if ($PreferredPath -and (Test-Path $PreferredPath)) { return $PreferredPath }
    if (Test-Path $Oscdimg) { return $Oscdimg }

    $cmd = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $searchRoots = @("$env:ProgramFiles(x86)\Windows Kits","$env:ProgramFiles\Windows Kits")
    foreach ($root in $searchRoots) {
        if ($root -and (Test-Path $root)) {
            $found = Get-ChildItem -Path $root -Filter oscdimg.exe -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($found) { return $found.FullName }
        }
    }

    Write-Host ""
    Write-Host "oscdimg.exe (Windows ADK) was not found. The Windows ADK (Deployment Tools) is required to build bootable ISOs." -ForegroundColor Yellow
    if ($NonInteractive) {
        throw 'oscdimg.exe not found (non-interactive).' 
    }
    Write-Host "Options:"
    Write-Host " 1) Download Windows ADK"
    Write-Host " 2) Enter path to oscdimg.exe"
    Write-Host " 3) Abort"

    do { $opt = Read-Host 'Choose 1,2 or 3' } while ($opt -notin '1','2','3')

    switch ($opt) {
        '1' {
            $adkUrl = 'https://go.microsoft.com/fwlink/?linkid=2337875'
            $dest = Join-Path $env:TEMP 'adksetup.exe'
            if (Test-Path $dest) {
                Write-Host "ADK installer already exists at $dest; will use existing file." -ForegroundColor Yellow
            } else {
                Write-Host "Attempting to download ADK installer to $dest..."
                try {
                    if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue) {
                        Start-BitsTransfer -Source $adkUrl -Destination $dest -ErrorAction Stop
                    } else {
                        Invoke-WebRequest -Uri $adkUrl -OutFile $dest -UseBasicParsing -ErrorAction Stop
                    }
                    Write-Host "Downloaded ADK installer to $dest"
                } catch {
                    Write-Warning "Download failed: $_. Opening browser as fallback."
                    Start-Process $adkUrl
                    Read-Host 'After installing ADK press Enter to continue (or Ctrl+C to abort)'
                    return Ensure-Oscdimg -PreferredPath $PreferredPath
                }
            }

                if ($NonInteractive) {
                    if ($AutoInstallAdk) {
                        Write-Host "Non-interactive: attempting silent ADK install..."
                        try {
                            $res = Run-SilentInstallerWithProgress -InstallerPath $dest -FeatureList $AdkFeatures
                            $exit = $res.ExitCode
                            if ($exit -eq 1001) { Write-Warning "Silent install exit code 1001 - possible rate limit from Microsoft. Consider waiting or downloading manually." }
                                if ($exit -ne 0) {
                                Write-Warning "Silent install returned exit code $exit. Please install ADK manually."
                                Write-Host "Stdout: $($res.StdOut)"
                                throw 'oscdimg.exe not found (ADK install failed).' 
                            }
                            Write-Host "Silent ADK install completed (exit code 0)."
                            return Ensure-Oscdimg -PreferredPath $PreferredPath
                        } catch {
                            Write-Warning "Silent install failed: $_. Please install ADK manually."
                            throw 'oscdimg.exe not found (ADK install failed).' 
                        }
                    }
                    Write-Host "Non-interactive: installer saved. Please install ADK and press Enter to continue."
                    Read-Host 'After installing ADK press Enter to continue (or Ctrl+C to abort)'
                    return Ensure-Oscdimg -PreferredPath $PreferredPath
                }

                Write-Host 'Run the ADK installer now? (I)nteractive / (S)ilent / (N)o [Default: S] (auto-selecting Silent in 10s)'
                $run = $null
                $timeoutSeconds = 10
                try {
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    while ($sw.Elapsed.TotalSeconds -lt $timeoutSeconds -and -not [Console]::KeyAvailable) { Start-Sleep -Milliseconds 100 }
                    if ([Console]::KeyAvailable) {
                        $key = [Console]::ReadKey($true)
                        $run = $key.KeyChar
                    } else {
                        $run = 'S'
                    }
                } catch {
                    $run = Read-Host 'Run the ADK installer now? (I)nteractive / (S)ilent / (N)o (default S)'
                    if (-not $run) { $run = 'S' }
                }

                switch ($run.ToUpper()) {
                    'I' {
                        Start-Process -FilePath $dest -Verb RunAs -Wait
                        Read-Host 'After installing ADK press Enter to continue (or Ctrl+C to abort)'
                        return Ensure-Oscdimg -PreferredPath $PreferredPath
                    }
                    'S' {
                        Write-Host "Running ADK installer silently..."
                            try {
                                $res = Run-SilentInstallerWithProgress -InstallerPath $dest -FeatureList $AdkFeatures
                                $exit = $res.ExitCode
                                if ($exit -eq 1001) { Write-Warning "Silent install exit code 1001 - possible rate limit from Microsoft. Consider waiting or downloading manually." }
                                if ($exit -ne 0) {
                                    Write-Warning "Silent install returned exit code $exit. Opening browser as fallback."
                                    Write-Host "Stdout: $($res.StdOut)"
                                    Start-Process $adkUrl
                                    Read-Host 'After installing ADK press Enter to continue (or Ctrl+C to abort)'
                                    return Ensure-Oscdimg -PreferredPath $PreferredPath
                                }
                                Write-Host "Silent install finished."
                                return Ensure-Oscdimg -PreferredPath $PreferredPath
                            } catch {
                                Write-Warning "Silent install failed: $_. Opening browser as fallback."
                                Start-Process $adkUrl
                                Read-Host 'After installing ADK press Enter to continue (or Ctrl+C to abort)'
                                return Ensure-Oscdimg -PreferredPath $PreferredPath
                            }
                    }
                    default {
                        Write-Host 'Installer saved. Opening browser to ADK page as a fallback.'
                        Start-Process $adkUrl
                        Read-Host 'After installing ADK press Enter to continue (or Ctrl+C to abort)'
                        return Ensure-Oscdimg -PreferredPath $PreferredPath
                    }
                }
        }
        '2' {
            do {
                $p = Read-Host 'Enter full path to oscdimg.exe'
                if (-not (Test-Path $p)) { Write-Warning 'Path not found.' }
            } while (-not (Test-Path $p))
            return $p
        }
        '3' { throw 'oscdimg.exe not found; aborting.' }
    }
}

# (oscdimg check will be performed when needed in Automatic mode)

# ---------------- Functions ----------------

# Hard cleanup of Build root, preserve Drivers folder
function Cleanup-BuildRootPreserveDrivers {
    param (
        [Parameter(Mandatory)]
        [string]$BuildRoot,
        [string]$PreserveFolderName = "Drivers"
    )

    Write-Host ""
    Write-Host "================ BUILD ROOT CLEANUP ================="
    Write-Host "Target   : $BuildRoot"
    Write-Host "Preserve : $PreserveFolderName"
    Write-Host "===================================================="

    if (-not (Test-Path $BuildRoot)) {
        Write-Host "Build root does not exist. Nothing to clean."
        return
    }

    # Unmount any lingering WIMs first
    Write-Host "Checking for mounted WIMs..."
    try {
        $mounted = dism.exe /Get-MountedWimInfo 2>$null
        foreach ($line in $mounted) {
            if ($line -match "Mount Dir : (.*)$") {
                $mountPath = $Matches[1].Trim()
                if ($mountPath -like "$BuildRoot*") {
                    Write-Warning "Unmounting WIM at $mountPath"
                    dism.exe /Unmount-Wim /MountDir:$mountPath /Discard
                }
            }
        }
    } catch {
        Write-Warning "Failed to enumerate mounted WIMs: $_"
    }

    # Enumerate folders to delete
    $targets = Get-ChildItem -Path $BuildRoot -Force | Where-Object { $_.Name -ne $PreserveFolderName }
    $deletedSummary = @()

    foreach ($item in $targets) {
        $targetPath = $item.FullName
        # Try fast delete first
        try {
            Remove-Item -Path $targetPath -Recurse -Force -ErrorAction Stop
            $deletedSummary += @{ Path = $targetPath; Result = 'Deleted' }
            continue
        } catch {
            # fall through to remediation
        }

        # Remediation steps for stubborn items (attempt once)
        try {
            cmd.exe /c "attrib -r -s -h `"$targetPath`" /S /D" | Out-Null
            takeown.exe /f "$targetPath" /r /d y | Out-Null
            icacls.exe "$targetPath" /reset /t /c | Out-Null
            icacls.exe "$targetPath" /grant Administrators:F /t /c | Out-Null
            Remove-Item -Path $targetPath -Recurse -Force -ErrorAction Stop
            $deletedSummary += @{ Path = $targetPath; Result = 'Deleted (remediated)' }
        } catch {
            $deletedSummary += @{ Path = $targetPath; Result = 'FAILED' }
        }
    }

    # Print consolidated summary
    Write-Host ""
    Write-Host "BUILD ROOT CLEANUP SUMMARY"
    foreach ($entry in $deletedSummary) {
        Write-Host " - $($entry.Path) : $($entry.Result)"
    }
    Write-Host ""
    Write-Host "Cleanup complete."

}

# Cleanup any leftover WIMs from previous sessions
function Cleanup-MountedWIMs {
    try {
        $mountedWims = dism.exe /Get-MountedWimInfo 2>$null | Select-String "Mount Dir"
        if ($mountedWims) {
            Write-Warning "Detected previously mounted WIM(s). Cleaning up previous session..."
            foreach ($line in $mountedWims) {
                if ($line -match "Mount Dir : (.*)$") {
                    $mountPath = $Matches[1].Trim()
                    Write-Host "Unmounting WIM at $mountPath (Discarding changes)"
                    dism.exe /Unmount-Wim /MountDir:$mountPath /Discard
                }
            }
        }
    } catch {
        Write-Warning "Error checking or unmounting WIMs: $_"
    }
}


function Get-WindowsProIndex {
    param ([string]$ImagePath)
    Write-Host "Detecting Windows Pro index..."
    $dismOutput = dism.exe /Get-WimInfo /WimFile:$ImagePath
    $currentIndex = $null
    foreach ($line in $dismOutput) {
        if ($line -match "^Index\s*:\s*(\d+)") { $currentIndex = $Matches[1] }
        if ($line -match "^Name\s*:\s*(.*Pro.*)") {
            Write-Host "Found Windows Pro at index $currentIndex"
            return $currentIndex
        }
    }
    throw "Windows Pro edition not found in image"
}

function Get-WindowsISO {
    # Non-interactive shortcut: use provided ISO path
    if ($NonInteractive -and $IsoPath) {
        if (Test-Path $IsoPath -PathType Leaf) { return @{ Path = $IsoPath; Source = 'Manual' } }
        else { throw "Non-interactive: provided ISO path not found: $IsoPath" }
    }
    if (Get-ChildItem -Path $StockISO -Filter *.iso -ErrorAction SilentlyContinue) {
        Write-Host "ISO found in StockISO. Skipping download or selection."
        $iso = Get-ChildItem -Path $StockISO -Filter *.iso | Select-Object -First 1
        return @{ Path = $iso.FullName; Source = "StockISO" }
    }

    Write-Host ""
    Write-Host "Select how you want to provide the Windows ISO:"
    Write-Host "1) Download ISO"
    Write-Host "2) Provide path to an existing ISO file"

    do { $choice = Read-Host "Enter 1 or 2" } while ($choice -notin "1","2")

    $TempISOFolder = $StockISO

    switch ($choice) {
        "1" {
            Write-Host "`Contacting Microsoft CDN."
            $dlUrl = 'https://raw.githubusercontent.com/Wh1t3Rose/Win11-DL/main/DL-WinISO.ps1'
            $destScript = $DLWinISOScript
            if (Test-Path $destScript) {
               # Write-Host "DL script already exists at $destScript; using existing file." -ForegroundColor Yellow
            } else {
                try {
                    if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue) {
                        Start-BitsTransfer -Source $dlUrl -Destination $destScript -ErrorAction Stop
                    } else {
                        Invoke-WebRequest -Uri $dlUrl -OutFile $destScript -UseBasicParsing -ErrorAction Stop
                    }
                    Write-Host "Downloaded DL-WinISO.ps1 to $destScript"
                } catch {
                    Write-Warning "Failed to download DL-WinISO.ps1: $_. Opening browser as fallback."
                    Start-Process 'https://github.com/Wh1t3Rose/Win11-DL/blob/main/DL-WinISO.ps1'
                    Read-Host 'After downloading/saving DL-WinISO.ps1 to the script folder press Enter to continue (or Ctrl+C to abort)'
                    if (-not (Test-Path $destScript)) { throw "DL-WinISO.ps1 not found at $destScript" }
                }
            }

            Write-Host "`Downloading ISO from Microsoft CDN."
            $outIsoPath = Join-Path $TempISOFolder 'Win11.iso'
                try {
                    # Run the downloaded script and capture its output (URLs or error text)
                    $dlOutput = & $destScript -Win 11 -Lang English -Arch x64 -Edition Pro -GetUrl 2>&1
                } catch {
                    Write-Warning "Failed to run local DL-WinISO.ps1: $_. Opening browser as fallback."
                    Start-Process 'https://github.com/Wh1t3Rose/Win11-DL/blob/main/DL-WinISO.ps1'
                    Read-Host 'After downloading/saving DL-WinISO.ps1 to the script folder press Enter to continue (or Ctrl+C to abort)'
                    $dlOutput = @()
                }

                # Detect explicit rate-limit message from the DL script
                $rateLimitMsg = 'Error: We are unable to complete your request at this time'
                if ($dlOutput -and ($dlOutput -join "`n") -match [regex]::Escape($rateLimitMsg)) {
                    Write-Warning "DL-WinISO reported: $rateLimitMsg"
                    Write-Host "You appear to be rate-limited. Please download the ISO manually from the provider and then provide its path." -ForegroundColor Yellow
                    do {
                        $manualIso = Read-Host 'Enter full path to the manually downloaded ISO (or Ctrl+C to abort)'
                        if (-not (Test-Path $manualIso -PathType Leaf)) { Write-Warning 'File not found. Please try again.' }
                    } while (-not (Test-Path $manualIso -PathType Leaf))
                    return @{ Path = $manualIso; Source = 'Manual' }
                }

                # Otherwise, treat output lines as URLs and download the first one to Win11.iso
                try {
                    $outIsoPath = Join-Path $TempISOFolder 'Win11.iso'
                    foreach ($line in $dlOutput) {
                        if ($line -match '^https?://') {
                            Invoke-WebRequest -Uri $line -OutFile $outIsoPath -UseBasicParsing -ErrorAction Stop
                            break
                        }
                    }
                } catch {
                    Write-Warning "Failed to download ISO from URL returned by DL-WinISO: $_. Opening browser as fallback."
                    Start-Process 'https://github.com/Wh1t3Rose/Win11-DL/blob/main/DL-WinISO.ps1'
                    Read-Host 'After downloading/saving DL-WinISO.ps1 to the script folder press Enter to continue (or Ctrl+C to abort)'
                }

            $DownloadedISO = Get-ChildItem -Path $TempISOFolder -Filter *.iso | Select-Object -First 1
            if (-not $DownloadedISO -and (Test-Path $outIsoPath)) { $DownloadedISO = Get-Item $outIsoPath }
            if (-not $DownloadedISO) {
                Write-Warning "DL-WinISO did not produce an ISO in $TempISOFolder"
                Write-Host "Please download the ISO using the Microsoft Media Creation Tool and provide the path to the resulting ISO file." -ForegroundColor Yellow
                $mctUrl = 'https://go.microsoft.com/fwlink/?linkid=2156295'
                $mctDest = Join-Path $env:TEMP 'MediaCreationTool.exe'
                Write-Host "Downloading Microsoft Media Creation Tool to $mctDest in 3 seconds... Press Ctrl+C to cancel." -ForegroundColor Yellow
                Start-Sleep -Seconds 3
                    try {
                        if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue) {
                            Start-BitsTransfer -Source $mctUrl -Destination $mctDest -ErrorAction Stop
                        } else {
                            Invoke-WebRequest -Uri $mctUrl -OutFile $mctDest -UseBasicParsing -ErrorAction Stop
                        }
                        Write-Host "Downloaded Media Creation Tool to: $mctDest"
                        Write-Host "Launching Media Creation Tool now (will request elevation)." -ForegroundColor Yellow
                        try {
                            Start-Process -FilePath $mctDest -Verb RunAs -Wait
                            Write-Host "Media Creation Tool has exited. If it created an ISO, provide its path below." -ForegroundColor Green
                        } catch {
                            Write-Warning "Failed to launch Media Creation Tool: $_. You can run $mctDest manually." 
                        }
                    } catch {
                        Write-Warning "Failed to download Media Creation Tool: $_. Opening browser as fallback."
                        Start-Process 'https://www.microsoft.com/software-download/windows11'
                    }
                    do {
                        $manualIso = Read-Host 'After running the Media Creation Tool, enter full path to the ISO (or Ctrl+C to abort)'
                        if (-not (Test-Path $manualIso -PathType Leaf)) { Write-Warning 'File not found. Please try again.' }
                    } while (-not (Test-Path $manualIso -PathType Leaf))
                Write-Host "Using provided ISO: $manualIso"
                return @{ Path = $manualIso; Source = 'Manual-MCT' }
            }
            Write-Host "ISO downloaded successfully: $($DownloadedISO.FullName)"
            return @{ Path = $DownloadedISO.FullName; Source = "DL" }
        }
        "2" {
            do {
                $ISOPath = Read-Host "Enter full path to your existing Windows ISO"
                if (-not (Test-Path $ISOPath -PathType Leaf)) { Write-Warning "File not found." }
            } while (-not (Test-Path $ISOPath -PathType Leaf))
            Write-Host "Using provided ISO: $ISOPath"
            return @{ Path = $ISOPath; Source = "Manual" }
        }
    }
}

function Show-Menu {
    Write-Host ""
    Write-Host '  ############################################################'
    Write-Host '  #                                                          #'
    Write-Host '  #               WINDOWS ISO BUILDER - Create-WinISO        #'
    Write-Host '  #         Extract • Inject Drivers • Rebuild (BIOS+UEFI)   #'
    Write-Host '  #                                                          #'
    Write-Host '  #        Tip: Run as Administrator for full capabilities   #'
    Write-Host '  #                                                          #'
    Write-Host '  ############################################################'
    Write-Host ""
    Write-Host "  1) Automatic (Full Build)     - Run the full build flow"
    Write-Host "  2) Download ISO               - Use DL-WinISO / MCT fallback"
    Write-Host "  3) Select Drivers             - Add or copy driver folders"
    Write-Host "  4) Build ISO                  - Extract, inject, and rebuild"
    Write-Host "  5) Exit                       - Quit the script"
    Write-Host ""

    # Decorative footer
    Write-Host '  ------------------------------------------------------------'

    do { $menuChoice = Read-Host 'Select an option (1-5)' } while ($menuChoice -notin '1','2','3','4','5')
    return $menuChoice
}

function Select-Drivers {
    $moreDrivers = $true
    $Manufacturers = @("Dell","Lenovo","Other")
    $Categories    = @("Wifi","Ethernet","Chipset","Storage","Video")

    while ($moreDrivers) {
        Write-Host "`nSelect manufacturer:"
        for ($i=0; $i -lt $Manufacturers.Count; $i++) {
            Write-Host "$($i+1)) $($Manufacturers[$i])"
        }
        do { $mChoice = Read-Host "Enter manufacturer number" } while (-not ($mChoice -ge 1 -and $mChoice -le $Manufacturers.Count))
        $manuName = $Manufacturers[$mChoice-1]

        Write-Host "`nSelect driver type:"
        for ($i=0; $i -lt $Categories.Count; $i++) {
            Write-Host "$($i+1)) $($Categories[$i])"
        }
        do { $catChoice = Read-Host "Enter driver type number" } while (-not ($catChoice -ge 1 -and $catChoice -le $Categories.Count))
        $catName = $Categories[$catChoice-1]

        do {
            $DriverPath = Read-Host "Enter path to extracted drivers for $manuName - $catName"
            if (-not (Test-Path $DriverPath)) {
                Write-Warning "Path not found."
                continue
            }
            $infFiles = Get-ChildItem -Path $DriverPath -Filter *.inf -File -Recurse
            if (-not $infFiles) {
                $confirm = Read-Host "No .inf files detected in this folder. Are you sure you want to continue? (Y/N)"
                if ($confirm -notin "Y","y") { $DriverPath = $null }
            }
        } while (-not $DriverPath)

        $destFolder = Join-Path $Drivers "$manuName\$catName"
        if (-not (Test-Path $destFolder)) { New-Item -ItemType Directory -Path $destFolder -Force | Out-Null }

        Copy-Item -Path "$DriverPath\*" -Destination $destFolder -Recurse -Force
        Write-Host "Drivers copied to $destFolder"

        $ans = Read-Host "`nDo you have more drivers to add? (Y/N)"
        if ($ans -notin "Y","y") { $moreDrivers = $false }
    }
}

# ---------------- Menu Execution ----------------
$MenuOption = Show-Menu
switch ($MenuOption) {
    "1" {
        Write-Host "`n=== AUTOMATIC MODE ==="
        # Hard cleanup of Build root
        Cleanup-BuildRootPreserveDrivers -BuildRoot $BuildRoot

        # Ensure necessary folders exist
        $RequiredFolders = @($TempISO, $WimMount, $CustomISO, $StockISO)
        $createdFolders = @()
        foreach ($folder in $RequiredFolders) {
            if (-not (Test-Path $folder)) {
                New-Item -ItemType Directory -Path $folder -Force | Out-Null
                $createdFolders += $folder
            }
        }
        if ($createdFolders.Count -gt 0) {
            Write-Host ""
            Write-Host "Required folders ensured."
        } else {
            Write-Host ""
            Write-Host "Required folders already present."
        }


        # Ensure oscdimg/ADK is available for later ISO build
        $Oscdimg = Ensure-Oscdimg -PreferredPath $Oscdimg -NonInteractive:$NonInteractive

        # Cleanup any leftover WIMs from previous sessions (now run after build root preparation)
        Cleanup-MountedWIMs

        # Get ISO
        $ISOInfo = Get-WindowsISO
        $WindowsISOPath = $ISOInfo.Path
        $ISOSource       = $ISOInfo.Source

        # Driver selection (skip in non-interactive mode)
        if ($NonInteractive) {
            Write-Host "Non-interactive: skipping driver selection." -ForegroundColor Yellow
        } else {
            Select-Drivers
        }
    }
    "2" {
        Write-Host "`n=== DOWNLOAD ISO ==="
        & $DLWinISOScript
    }
    "3" {
        Write-Host "`n=== SELECT DRIVERS ==="
        if ($NonInteractive) {
            Write-Host "Non-interactive: skipping driver selection." -ForegroundColor Yellow
        } else {
            Select-Drivers
        }
    }
    "4" {
        Write-Host "`n=== BUILD ISO ==="
        Cleanup-BuildRootPreserveDrivers -BuildRoot $BuildRoot
        $RequiredFolders = @($TempISO, $WimMount, $CustomISO, $StockISO)
        $createdFolders = @()
        foreach ($folder in $RequiredFolders) {
            if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null; $createdFolders += $folder }
        }
        if ($createdFolders.Count -gt 0) {
            Write-Host ""
            Write-Host "Required folders ensured."
        } else {
            Write-Host ""
            Write-Host "Required folders already present."
        }
        $ISOInfo = Get-WindowsISO
        $WindowsISOPath = $ISOInfo.Path
        $ISOSource       = $ISOInfo.Source
    }
    "5" {
        Write-Host "Exiting..."
        exit 0
    }
}

# ---------------- Copy ISO to StockISO ----------------
$ExistingISO = Get-ChildItem -Path $StockISO -Filter *.iso -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $ExistingISO -or $ExistingISO.FullName -ne $WindowsISOPath) {
    Write-Host "=================== COPYING ISO TO STOCKISO ==================="
    Copy-Item $WindowsISOPath -Destination $StockISO -Force
} else {
    Write-Host "ISO already exists in StockISO. Skipping copy."
}

# ---------------- Extract ISO ----------------
$IsoFile = Get-ChildItem -Path $StockISO -Filter *.iso | Select-Object -First 1
$TempExtractMarker = Join-Path $TempISO ".extracted"

if (Test-Path $TempExtractMarker) {
    Write-Host "ISO already extracted. Skipping extraction."
} else {
    Write-Host "=================== EXTRACTING ISO ==================="
    $DiskImage = Mount-DiskImage -ImagePath $IsoFile.FullName -PassThru
    $Volume    = $DiskImage | Get-Volume
    if (-not $Volume.DriveLetter) { Dismount-DiskImage -ImagePath $IsoFile.FullName; throw "Failed to determine ISO drive letter" }

    $IsoDrive = "$($Volume.DriveLetter):\"
    Copy-Item -Path "$IsoDrive*" -Destination $TempISO -Recurse -Force
    Dismount-DiskImage -ImagePath $IsoFile.FullName
    New-Item -Path $TempExtractMarker -ItemType File | Out-Null
}

Write-Host "`n=================== BUILDING WINDOWS ISO ==================="

# ---------------- Detect Install Image ----------------
$InstallESD = Join-Path $TempISO "sources\install.esd"
$InstallWIM = Join-Path $TempISO "sources\install.wim"
if (-not (Test-Path $InstallESD) -and -not (Test-Path $InstallWIM)) { throw "install.esd or install.wim not found after extraction" }

$SourceImage = if (Test-Path $InstallESD) { $InstallESD } else { $InstallWIM }
$WorkingWIM  = Join-Path $TempISO "sources\install.wim"

# ---------------- Detect Windows Pro Index ----------------
$ProIndex = Get-WindowsProIndex -ImagePath $SourceImage

# ---------------- Convert ESD to WIM if needed ----------------
if ($SourceImage -like "*.esd") {
    Write-Host "Converting ESD to WIM (Index $ProIndex)..."
    dism.exe /Export-Image `
        /SourceImageFile:$SourceImage `
        /SourceIndex:$ProIndex `
        /DestinationImageFile:$WorkingWIM `
        /Compress:XPRESS `
        /CheckIntegrity
} else {
    Copy-Item $SourceImage $WorkingWIM -Force
}

# ---------------- Mount WIM ----------------
Write-Host "Mounting WIM for driver injection..."
$Mounted = dism.exe /Get-MountedWimInfo 2>$null
if ($Mounted) {
    $Mounted | ForEach-Object {
        if ($_ -match "Mount Dir : (.*)$") {
            $MountPath = $Matches[1].Trim()
            dism.exe /Unmount-Wim /MountDir:$MountPath /Discard
        }
    }
}
dism.exe /Mount-Wim /WimFile:$WorkingWIM /Index:1 /MountDir:$WimMount

# ---------------- Add Drivers ----------------
Write-Host "Injecting drivers from $Drivers..."
dism.exe /Image:$WimMount /Add-Driver /Driver:$Drivers /Recurse

# ---------------- Commit WIM ----------------
Write-Host "Committing WIM changes..."
dism.exe /Unmount-Wim /MountDir:$WimMount /Commit

# ---------------- Convert WIM back to ESD ----------------
if (Test-Path $InstallESD) { attrib -r $InstallESD; Remove-Item $InstallESD -Force }
Write-Host "Converting WIM back to ESD..."
dism.exe /Export-Image `
    /SourceImageFile:$WorkingWIM `
    /SourceIndex:1 `
    /DestinationImageFile:$InstallESD `
    /Compress:Recovery

Remove-Item $WorkingWIM -Force

$# ---------------- Build Bootable ISO ----------------

# Ensure oscdimg.exe is available; prompt user to install ADK if missing
# ---------------- Build Bootable ISO ----------------
# Boot file paths
$BootBIOS = Join-Path $TempISO "boot\etfsboot.com"
$BootUEFI = Join-Path $TempISO "efi\microsoft\boot\efisys.bin"
if (-not (Test-Path $BootBIOS)) { throw "Missing BIOS boot file" }
if (-not (Test-Path $BootUEFI)) { throw "Missing UEFI boot file" }

Write-Host "Creating bootable ISO at $ISOOut..."
& $Oscdimg `
    -m `
    -o `
    -u2 `
    -udfver102 `
    -bootdata:2#p0,e,b"$BootBIOS"#pEF,e,b"$BootUEFI" `
    $TempISO `
    $ISOOut

Write-Host "=== SUCCESS ==="
Write-Host "Windows Pro ISO created at: $ISOOut"