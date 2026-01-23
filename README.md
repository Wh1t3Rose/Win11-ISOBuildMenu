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
