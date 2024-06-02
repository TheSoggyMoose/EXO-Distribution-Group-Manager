# Distribution Group Manager

## Description
The Distribution Group Manager is a PowerShell script, now also available as an executable, designed to simplify the management of Exchange Online distribution groups. It provides a simple user-friendly graphical interface for adding or removing members from distribution groups without the need for in-depth knowledge of Exchange Online PowerShell cmdlets.

## Features
- **Executable Version:** Packaged as an `.exe` for ease of use.
- **Graphical User Interface (GUI):** Easy-to-use interface for managing distribution groups.
- **Multi-Group Selection:** Allows the selection of multiple groups within a tenant for batch processing.
- **User Feedback:** Provides confirmation messages upon successful addition or removal of users.
- **Module Check and Installation:** Automatically checks for the required `ExchangeOnlineManagement` module and installs it if not present. The `.exe` will prompt for administrator privileges if the module needs to be installed.

## Requirements
- PowerShell 5.1 or higher
- ExchangeOnlineManagement module
- Administrator privileges if the module is not installed

## Usage
To run the executable, right-click on `DistributionGroupManager.exe` and select 'Run as administrator' if the `ExchangeOnlineManagement` module is not already installed.

For the PowerShell script version, run the following command:
```powershell
.\DistributionGroupManager.ps1
