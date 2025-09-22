# Microsoft.Store.Appx.Downloader


<img style="width:50%; height:auto;" alt="image" src="https://github.com/user-attachments/assets/24fd8894-d27e-42d4-9e1a-e8e9bef8949f" />

# Microsoft Store Appx Downloader and Installer

This PowerShell script provides a robust solution for downloading and installing Windows UWP (Universal Windows Platform) applications and their dependencies directly from the Microsoft Store. Unlike the native Store app, this tool offers granular control, allowing you to manage and deploy applications in various scenarios, including offline or restricted network environments. The project is designed with a user-friendly GUI built with WPF, making the entire process accessible and efficient.

---

## What It Does 🚀

The core function of this script is to act as an advanced package manager for Windows Store apps. It automates several critical tasks that are otherwise difficult or impossible to perform manually.

### Key Features:
* **Download App Packages:** The script can download all necessary installation files, including `.appx`, `.appxbundle`, `.msix`, and `.msixbundle` formats. It fetches these files by communicating with a third-party service, `store.rg-adguard.net`, which provides direct download links to the app packages.
* **Architecture Filtering:** The tool intelligently filters packages based on your system's architecture (`x64`, `x86`, `ARM`) or allows you to specify a different one. This ensures you download only the correct and compatible files, saving bandwidth and storage space.
* **Dependency Resolution:** A major advantage is its ability to download all required dependencies along with the main application package. This is crucial for offline installations, as it guarantees that all components needed for the app to function are available.
* **Customizable Downloads:** You can select from a predefined list of popular Windows apps or provide a custom Microsoft Store URL to download any publicly available app.
* **Bulk and Incremental Downloads:** The "Download All" feature allows you to fetch a collection of essential apps at once. The script also performs a file existence check, skipping downloads for any packages you already have, thus preventing redundant data transfer.
* **Offline Installation:** The "Install" functions enable you to install the downloaded packages without an internet connection. The script supports installing from a single directory or from multiple subfolders, ideal for bulk deployment.
* **Graphical User Interface (GUI):** A simple but effective WPF-based GUI makes the tool intuitive. You can easily select apps, specify download paths, choose architectures, and monitor the progress in real-time.

---

## Who It's For 🎯

This tool is a valuable asset for several user groups, streamlining app management and deployment tasks.

* **System Administrators & IT Professionals:** You can use this script to create a local, offline repository of core Windows applications. This is essential for deploying systems in secure or isolated networks where access to the Microsoft Store is restricted. It also serves as a reliable method for reinstalling default apps that might be corrupted or missing.
* **Power Users & Enthusiasts:** If you frequently set up new Windows machines, this script can save a significant amount of time by allowing you to batch-download all your favorite apps and their dependencies in advance. You can then quickly install them on a fresh OS install without a persistent internet connection.
* **Software Developers:** The tool can be used to acquire different versions of a UWP app for testing across various architectures or to debug installation issues with dependencies.

---

## Why You Need It ✨

The native Microsoft Store app, while convenient, lacks the flexibility and control required for advanced management scenarios. This PowerShell script fills that gap by providing a powerful, scriptable alternative.

### Common Use Cases:
* **Bypassing Store Restrictions:** In environments with strict firewalls or no internet access, this script is the only way to get official app packages.
* **Restoring Core Apps:** If a Windows update or system corruption removes built-in UWP apps like Notepad or Paint, this tool provides a straightforward way to reinstall them with all their dependencies.
* **Creating a Master Image:** For IT professionals, this script can be integrated into a post-installation script to automatically populate a new Windows image with all necessary UWP applications.
* **Portability:** Once you have downloaded the `.appx` and `.msix` files, they are completely portable. You can store them on a USB drive and install them on any compatible Windows PC, anywhere.
