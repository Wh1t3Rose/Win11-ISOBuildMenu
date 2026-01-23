# Create-WinISO

Build a customized Windows Pro ISO with injected drivers, automated tool handling, and interactive or non-interactive workflows.

**Summary**
- A PowerShell utility to extract, modify (inject drivers), and rebuild a bootable Windows Pro ISO (BIOS + UEFI).
- Automates common tasks: ADK/oscdimg discovery or install, ISO download (via DL-WinISO), and fallback to the Media Creation Tool.

**Features**
- Interactive menu with modes: Automatic, Download ISO, Select Drivers, Build.
- Reuses existing ADK installer (adksetup.exe) if present; optional silent install with progress.
- Handles ADK silent-install rate limiting (exit code 1001) with retry and warnings.
- Obtain Microsoft CDN ISO URLs; detects rate-limit responses and prompts for manual ISO.
- If automatic download fails, downloads the Media Creation Tool (MCT), launches it elevated, and prompts for the resulting ISO.
- Driver selection and injection into the image using DISM; rebuilds a bootable ISO with `oscdimg.exe`.

**Requirements**
- Windows (script must be run as Administrator).
- PowerShell (Windows PowerShell or PowerShell Core).
- Network access for downloads (ADK, DL-WinISO, MCT, ISO).
- `DISM`, `Mount-DiskImage` available on the host.
- `oscdimg.exe` (Windows ADK Deployment Tools) — the script can help download/install it.

**Quick Start**
- Interactive (recommended):

  powershell -NoProfile -ExecutionPolicy Bypass -File Packaging/!PS-Scripts/Create-WinISO.ps1

- Non-interactive (CI-style):

  pwsh -File Packaging/!PS-Scripts/Create-WinISO.ps1 -NonInteractive -IsoPath 'C:\isos\win11.iso' -OscdimgPath 'C:\Tools\oscdimg.exe'

**Parameters**
- `-BuildRoot` : (string) Override build root directory. Default: `C:\Build`.
- `-Drivers` : (string) Path to drivers folder. Default: `C:\Build\Drivers`.
- `-IsoPath` : (string) Path to an existing Windows ISO (used in non-interactive mode).
- `-OscdimgPath` : (string) Explicit path to `oscdimg.exe`.
- `-NonInteractive` : (switch) Run without interactive prompts (fail fast on missing inputs).
- `-DryRun` : (switch) Show actions without performing changes.
- `-AutoInstallAdk` : (switch) Attempt unattended ADK install when missing.
- `-AdkFeatures` : (string[]) Features to install with ADK (default: `('Deployment Tools','Imaging and Configuration Designer')`).

**ADK / oscdimg behavior**
- The script tries to locate `oscdimg.exe` by parameter, default path, or common Windows Kits locations.
- If not found, `Ensure-Oscdimg` offers to download `adksetup.exe` to `%TEMP%` and:
  - Reuse an existing `adksetup.exe` if present.
  - Offer Interactive or Silent install. Silent installs run with `/quiet /norestart /ceip off` and limited features.
  - On silent-install exit code `1001` (rate limit), the script retries with exponential backoff and warns the user.
  - If automated install fails, falls back to opening the ADK download page and prompting the user.

**ISO Download behavior (DL-WinISO & MCT fallback)**
- When choosing "Download ISO", the script downloads connects to MS CDNs to grab an ISO.
- The script detects the rate-limit message `Error: We are unable to complete your request at this time` and will prompt the user to provide a manually downloaded ISO if rate-limited.
- If `DL-WinISO` does not produce an ISO, the script downloads the Microsoft Media Creation Tool (MCT) after a 3-second notice, launches it elevated, and prompts the user to supply the created ISO path.

**Driver injection & ISO build**
- Drivers copied into the `Drivers` folder are injected into the `install.wim`/`install.esd` image via DISM.
- The script mounts the image, injects drivers recursively from the `Drivers` folder, commits changes, converts back to ESD if needed, and runs `oscdimg.exe` to produce a bootable ISO (`Windows11_VACO.iso` by default).

**Troubleshooting**
- Run the script as Administrator.
- If ADK silent install fails repeatedly with exit code 1001, wait a while or download ADK manually from Microsoft and run the installer interactively.
- If `DL-WinISO` is rate-limited, download the ISO using another machine/network or use the Media Creation Tool as prompted.
- Review logs under `%TEMP%\adk_install_logs` for ADK installer stdout when diagnosing install issues.

**Contributing**
- Bug reports and improvements welcome. Keep changes focused and preserve interactive UX.
- If you update behavior for ADK/MCT flows or DL-WinISO integration, update the changelog in `Create-WinISO.ps1`.

**Notes**
- The script is designed to be interactive-first, but supports non-interactive automation with `-NonInteractive` and explicit paths.
- Many operations (ADK install, DISM, oscdimg) require elevation and may modify system state; test in an isolated environment when possible.

---

## Changelog
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
