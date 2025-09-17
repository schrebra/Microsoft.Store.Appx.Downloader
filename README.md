# Microsoft.Store.Appx.Downloader

<img style="width:50%; height:auto;" alt="image" src="https://github.com/user-attachments/assets/06c50864-15a0-4104-9887-a0f82a40f12c" />
<img style="width:50%; height:auto;" alt="image" src="https://github.com/user-attachments/assets/d3e349e6-d338-4c47-a5a8-65088bd7dd65" />


# UWP App Package Downloader and Installer (PowerShell GUI)

This PowerShell script provides a graphical user interface (GUI) built with XAML to download and install Universal Windows Platform (UWP) application packages directly from the Microsoft Store. It fetches `.appx`, `.appxbundle`, `.msix`, and `.msixbundle` files along with their dependencies.

**Why this script is useful:**

*   **Offline Installation:** Download UWP apps for installation on systems without internet access.
*   **Version Control:** Obtain specific versions of apps, which can be crucial for development, testing, or compatibility.
*   **Bulk Downloads:** Easily download multiple app packages and their dependencies.
*   **System Administration:** Deploy UWP apps across multiple machines in an enterprise environment without relying on the Microsoft Store app.
*   **Troubleshooting:** Reinstall problematic UWP apps from scratch with known good packages.

## Features

*   **Download UWP Packages:** Fetches `.appx`, `.msix`, `.appxbundle`, `.msixbundle` files and their dependencies.
*   **Select Architecture:** Choose between `Auto`, `Neutral`, `x64`, `x86`, and `ARM` architectures for downloads.
*   **Pre-defined Apps:** Quick selection for common Microsoft Store apps like Clock, Paint, Photos, etc.
*   **Custom URL Support:** Allows pasting any Microsoft Store app URL for download.
*   **Browse Download Path:** Easily select a local folder to save downloaded packages.
*   **Progress Tracking:** Provides real-time feedback on download and installation progress.
*   **Install Packages:** Functionality to install all downloaded UWP packages from a specified directory.
*   **GUI Interface:** User-friendly GUI for ease of use.
*   **Background Operations:** Downloads and installations run in separate PowerShell runspaces, keeping the UI responsive.

## How it Works

The script leverages the `store.rg-adguard.net` API to retrieve direct download links for UWP packages from the Microsoft Store. It then uses `Invoke-WebRequest` to download these files. For installation, it utilizes the built-in `Add-AppxPackage` PowerShell cmdlet.

## Who is this for?

*   **System Administrators:** For deploying and managing UWP applications in corporate environments.
*   **Developers:** For testing app installations, managing dependencies, or working with specific app versions.
*   **IT Enthusiasts:** For those who prefer direct package management or need to install apps offline.
*   **Users with Limited Internet Access:** To download apps once and install them on multiple machines.

## Installation and Usage

1.  **Save the Script:** Save the entire code block as a `.ps1` file (e.g., `DownloadUWPApps.ps1`).

2.  **Unblock the Script (if necessary):** If you downloaded the script, Windows might mark it as untrusted. Right-click the `.ps1` file, go to `Properties`, and check the `Unblock` box at the bottom, then click `OK`.

3.  **Run the Script:** Open PowerShell (as Administrator is recommended for installation) and navigate to the directory where you saved the script, then run it:

    ```powershell
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process # Only if you face execution policy issues
    .\DownloadUWPApps.ps1
    ```

    You can also simply right-click the `.ps1` file and choose "Run with PowerShell."

4.  **Using the GUI:**
    *   **Select an App:** Choose a common app from the dropdown or select "Custom URL" to paste your own.
    *   **Enter URL:** If "Custom URL" is selected, paste the Microsoft Store app URL into the text box. (Example: `https://apps.microsoft.com/detail/9wzdncrfj3pr?hl=en-US&gl=US`)
    *   **Set Download Path:** Specify where you want the packages to be saved. Use the "Browse" button to select a folder.
    *   **Choose Architecture:** Select the desired processor architecture for the packages.
    *   **Download:** Click the "Download" button to start fetching the files.
    *   **Install:** Once downloaded, click "Install Packages" to install all `.appx`/`.msix` files found in the download directory.

## Troubleshooting

*   **`store.rg-adguard.net` API:** This script relies on an external service to get download links. If the service is down or changes its API, the download functionality may break.
*   **Execution Policy:** If you encounter errors running the script, your PowerShell execution policy might be preventing it. Use `Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process` in your PowerShell session before running the script.
*   **Administrator Rights:** Installing APPX/MSIX packages often requires Administrator privileges. Run PowerShell as an Administrator when executing the script, especially if you plan to use the "Install Packages" feature.
*   **Network Issues:** Ensure you have a stable internet connection for downloading.
*   **"Couldn't determine file size" warning:** This might occur if the `Content-Length` header is not provided by the download server. The download will likely still proceed, but progress tracking for that specific file might be less accurate.

---

**Disclaimer:** This script uses a third-party service (`store.rg-adguard.net`) to obtain direct download links for Microsoft Store apps. Use at your own discretion. Always ensure you trust the source of any script you run on your system.
