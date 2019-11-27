# ExchangeTrackingGUI
This is a Graphical User Interface for Get-MessageTrackingLog PowerShell command: Message Tracking Log Explorer Tool for Exchange 2016

## How to use it
1. Install script from PowerShellGallery:
  ```powershell
  Install-Script -Name ExchangeTrackingGUI
  ```
2. Start PowerShell console from an account with administrative rights to Exchange.
3. Run script: **ExchangeTrackingGUI.ps1**.
4. Click **[Config]** button and configure parameters. Use FQDN for 'Exchange PS connect' field.
5. Click **[Save]** button. Configuration will be saved at _"$($env:LOCALAPPDATA)\ExchangeTrackingGUI\Config.json"_ file.
6. Use **[Search]** button for view results or **[Export]** button for view and save CSV file.
