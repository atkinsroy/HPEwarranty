# HPEwarranty
Find warranty and support contract information for HPE hardware using PowerShell

## Requirements

Selenium PowerShell Module downloaded from the PowerShell Gallery.

## Installation

To use the script on an Internet facing computer:
    1. Install the PowerShell Module called Selenium (https://github.com/adamdriscoll/selenium-powershell) from the
    PowerShell Gallery:
    
    ```powershell
        PS> install-module Selenium -Scope AllUsers
    ```
    
    2. Use Get-module Selenium -Listavailable to see where it is installed. This depends on the installed
    PowerShell version and on the scope chosen in the above command
    
    3. If using PowerShell 7.x and scope is AllUsers, enter the following (adjust as necessary):
    
    ```powershell
        PS> $env:PATH += "C:\Program Files\PowerShell\Modules\Selenium\3.0.1\assemblies\"
        PS> Add-Type -Path "C:\Program Files\PowerShell\Modules\Selenium\3.0.1\assemblies\WebDriver.dll"
    ```
    
    4. Selenium ships with Browser drivers for Chrome, Firefox, Edge. These will probably be out of date. Find the 
    installed version of the broswer you intend to use and download the appropriate driver. For example, with 
    Chrome, download the matching driver from  https://sites.google.com/a/chromium.org/chromedriver/downloads, 
    and replace the chromedriver.exe file in the same Selenium folder, as above.

## Example usage
    ```powershell
    PS C:\> .\Get-HPEwarranty -Filename serial.csv -verbose
    ```
    Get the warranty information for all serial numbers in the specified CSV file. Use -Verbose to provide status
    messages.
    
    ```powershell
    PS C:\> .\Get-HPEwarranty -Filename serial.csv -verbose -ignore XXX1111XXXX,YYY2222YYYY
    ```
    Use the -ignore parameter to exclude one or more problematic serial numbers. These will show up in the final
    report with a suitable message.

    ```powershell
    PS C:\> .\Get-HPEwarranty -Filename serial.csv -verbose -ignore C:\Ignore.txt
    ```
    Use the -ignore parameter to exclude one or more problematic serial numbers from a text file. These will show up in the final
    report with a suitable message.
