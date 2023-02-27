
# Download and Install Office on Windows 7/10/11




## Install Office on Windows 7/10/11

  
## Installation

Right click on the Windows Start icon 🪟 then select Windows PowerShell (Admin). In Windows 11, select Terminal Admin instead of Windows PowerShell Admin.

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/cons/powershell10.jpg)

Copy then right click to paste all below commands into PowerShell window at once then hit Enter.

```ps
Set-ExecutionPolicy Bypass -Scope Process -Force
$url='https://github.com/bonben365/office-installer/raw/main/install.ps1'
iex ((New-Object System.Net.WebClient).DownloadString($url))
```




➡️Please inspect [https://github.com/bonben365/office-installer/raw/main/install.ps1](https://github.com/bonben365/office-installer/raw/main/install.ps1) prior to running any of these scripts to ensure safety. We already know it's safe, but you should verify the security and contents of any script from the internet you are not familiar with.

## Documentation

[🌍https://bonben365.com/](https://bonben365.com/)

