# ExchangeTrackingGUI
This is a Graphical User Interface for Get-MessageTrackingLog PowerShell command: Message Tracking Log Explorer Tool for Exchange 2016

## Releases

### GitHub master branch
[![Build status][appveyor-badge-master]][appveyor-build-master]
This is the branch containing the latest release, published at PowerShell Gallery.

### GitHub dev branch
[![Build status][appveyor-badge-dev]][appveyor-build-dev]
This is the development branch with latest changes.

### PowerShell Gallery
[![PowerShell Gallery Version][psgallery-version-badge]][psgallery]
[![PowerShell Gallery Downloads][psgallery-badge]][psgallery]
Official repository - latest module version and download count.

## Version
0.9.1

## How to use it
1. Install module from PowerShellGallery:
  ```powershell
  Install-Module -Name ExchangeTrackingGUI
  ```
2. Start PowerShell console from an account with administrative rights to Exchange.
3. Run Command **Get-MessageTrackingGUI**.
4. Click **[Config]** button and configure parameters. Use FQDN for 'Exchange PS connect' field. Use **Enter**, **ESC** and **Del** hot keys.
5. Click **[Save]** button. Configuration will be saved at <span style="color:green">_"$env:LOCALAPPDATA)\ExchangeTrackingGUI\Config.json"_</span> file.
6. Use **[Search]** button for view results or **[Export]** button for view and save CSV file.

## Screenshots
![](https://pvs043.github.io/ExchangeTrackingGUI/MainForm.png) <img align="top" src="https://pvs043.github.io/ExchangeTrackingGUI/ConfigForm.png">

## License

**MIT**

[appveyor-badge-master]: https://ci.appveyor.com/api/projects/status/jyrr9ji54lqxmmt7?branch=master&svg=true
[appveyor-build-master]: https://ci.appveyor.com/project/pvs043/exchangetrackinggui/branch/master?fullLog=true
[appveyor-badge-dev]: https://ci.appveyor.com/api/projects/status/jyrr9ji54lqxmmt7?branch=dev&svg=true
[appveyor-build-dev]: https://ci.appveyor.com/project/pvs043/exchangetrackinggui/branch/dev?fullLog=true
[psgallery-badge]: https://img.shields.io/powershellgallery/dt/exchangetrackinggui.svg
[psgallery]: https://www.powershellgallery.com/packages/exchangetrackinggui
[psgallery-version-badge]: https://img.shields.io/powershellgallery/v/exchangetrackinggui.svg
[psgallery-version]: https://www.powershellgallery.com/packages/exchangetrackinggui
