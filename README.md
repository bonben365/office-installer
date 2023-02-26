
# Download and Install Office 365 on Windows 7




## Requirements

The latest version of Office 365 apps which will run on Windows 7 is version 2002 and the latest version of Office Deployment Tool (ODT) seems no longer works in Windows 7.
  
## Installation

- Search then open Windows PowerShell as administrator.
- Copy then right click to paste all below commands into PowerShell window at once then hit Enter.


```bash
$url="https://raw.githubusercontent.com/bonben365/office365-win7/main/install.ps1"
Set-ExecutionPolicy Bypass -Scope Process -Force
iex ((New-Object System.Net.WebClient).DownloadString($url))
```
➡️Please inspect [https://raw.githubusercontent.com/bonben365/office365-win7/main/install.ps1](https://raw.githubusercontent.com/bonben365/office365-win7/main/install.ps1) prior to running any of these scripts to ensure safety. We already know it's safe, but you should verify the security and contents of any script from the internet you are not familiar with.

## Screenshots

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/b9CLMjMxxfzks7hFUrjHpPBiyFPOGjGQVawPAZZGk9wAkFHyij3yq7Kjl2WX.jpg)

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/vPvb9WHBpInz6Z0LSSRPpChEdrITY0YfXvNOvGNC7finthFKlSZNlYwvM03n.jpg)


## Documentation

[🌍https://bonben365.com/](https://bonben365.com/)

