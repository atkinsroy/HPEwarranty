# HPEwarranty
Find warranty and support contract information for HPE hardware using PowerShell

## Requirements

* Internet facing computer
* Selenium PowerShell Module downloaded from the PowerShell Gallery

## Installation

To use the script on an Internet facing computer:
1. Install the PowerShell Module called Selenium (https://github.com/adamdriscoll/selenium-powershell) from the PowerShell Gallery:
    
```powershell
    PS> install-module Selenium -Scope AllUsers
    PS> Get-module Selenium -Listavailable
```
    
The location of the installed module will depend on the version of PowerShell you have and on the scope chosen above (Allusers/CurrentUser)
    
2. If you are using PowerShell 7.x and scope is AllUsers, enter the following (adjust as necessary):
    
```powershell
    PS> $env:PATH += "C:\Program Files\PowerShell\Modules\Selenium\3.0.1\assemblies\"
    PS> Add-Type -Path "C:\Program Files\PowerShell\Modules\Selenium\3.0.1\assemblies\WebDriver.dll"
```

3. Selenium ships with Browser drivers for Chrome, Firefox, Edge. These will probably be out of date for the version of the broswer you have installed. You will need to download a matching driver for the version of the brower you are using. For Chrome, install the appropriate driver from  https://sites.google.com/a/chromium.org/chromedriver/downloads, and replace the chromedriver.exe file in the same Selenium install folder, as above.

## Example usage
```powershell
    PS C:\> .\Get-HPEwarranty -Filename serial.csv -verbose
```

Get the warranty information for all serial numbers in the specified CSV file. Use -Verbose to provide status messages. This will create a new file called serial_new.csv with the HPE warranty information appended to the original objects.
    
```powershell
    PS C:\> .\Get-HPEwarranty -Filename serial.csv -verbose -ignore XXX1111XXXX,YYY2222YYYY
```

Use the -ignore parameter to exclude one or more problematic serial numbers. These will show up in the final report with a suitable message.

```powershell
    PS C:\> .\Get-HPEwarranty -Filename serial.csv -verbose -ignore C:\Ignore.txt
```

Use the -ignore parameter to exclude one or more problematic serial numbers from an input text file. These will show up in the final report with a suitable message.
