<#
.SYNOPSIS
    Find warranty and support contract information for HPE hardware, given a list of valid serial numbers.
.DESCRIPTION
    Instructions to use on an Internet facing computer:
    1. Install the PowerShell Module called Selenium (https://github.com/adamdriscoll/selenium-powershell) from the
    PowerShell Gallery:

        PS> install-module Selenium -Scope AllUsers
    
    2. Use Get-module Selenium -Listavailable to see where it is installed. This depends on the installed
    PowerShell version and on the scope chosen in the above command
    
    3. If using PowerShell 7.x and scope is AllUsers, enter the following (adjust as necessary):

        PS> $env:PATH += "C:\Program Files\PowerShell\Modules\Selenium\3.0.1\assemblies\"
        PS> Add-Type -Path "C:\Program Files\PowerShell\Modules\Selenium\3.0.1\assemblies\WebDriver.dll"
    
    4. Selenium ships with Browser drivers for Chrome, Firefox, Edge. These will probably be out of date. Find the 
    installed version of the broswer you intend to use and download the appropriate driver. For example, with 
    Chrome, download the matching driver from  https://sites.google.com/a/chromium.org/chromedriver/downloads, 
    and replace the chromedriver.exe file in the same Selenium folder, as above.
    
    5. The script accepts a single CSV containing one or more valid serial numbers. The CSV column must be labelled
    "SerialNumber". Other columns can exist, warranty and support details are appended to the existing information
    in the file.

    6. The website sometimes throws 500 errors occasionally. It also asks for a Confirmation Code sometimes too. 
    In both cases, the script will retry with the same set of serial numbers. If there is an unrecognised serial 
    number, or a serial number has no warranty or support details, the whole set of serial numbers will fail. The 
    script pauses to allow you to capture problem serial numbers so you can ignore them with -ignore.
.EXAMPLE
    PS C:\> .\Get-HPEwarranty -Filename serial.csv -verbose

    Get the warranty information for all serial numbers in the specified CSV file. Use -Verbose to provide status
    messages.
.EXAMPLE
    PS C:\> .\Get-HPEwarranty -Filename serial.csv -verbose -ignore XXX1111XXXX,YYY2222YYYY

    Use the -ignore parameter to exclude one or more problematic serial numbers. These will show up in the final
    report with a suitable message.
.EXAMPLE
    PS C:\> .\Get-HPEwarranty -Filename serial.csv -verbose -ignore C:\Ignore.txt

    Use the -ignore parameter to exclude one or more problematic serial numbers from a text file. These will show up in the final
    report with a suitable message.
.INPUTS
    A list of serial numbers in a CSV file. Column header must be 'SerialNumber'
.OUTPUTS
    A list of serial numbers in a new CSV file with warranty/support information appended to the original
    information
#>

<#
(C) Copyright 2020 Hewlett Packard Enterprise Development LP

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the "Software"),
to deal in the Software without restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, sublicense,
and/or sell copies of the Software, and to permit persons to whom the
Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR
OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$filename,

    [parameter(Mandatory = $false)]
    [string[]]$Ignore
)

begin {
    # define a helper function to lookup up to 10 serial numbers on the website
    Function Resolve-HPEwarranty {
        [CmdletBinding()]
        param (
            [System.Array]$SerialNumber
        )
        Write-Verbose "Processing $($SerialNumber -join ',')..."

        # start chrome
        $Driver = Start-SeChrome -Quiet
        Open-SeUrl -Url https://support.hpe.com/hpsc/wc/public/home -Driver $Driver
    
        $Count = 0
        foreach ($sn in $SerialNumber) {
            # enter this serial number in the Web form
            $WebElement = Get-SeElement -Driver $Driver -Id "serialNumber$Count"
            Send-SeKeys -Element $WebElement -Keys $sn
            $Count += 1
        }

        # click Submit
        $Send = Get-SeElement -Driver $Driver -Name "submitButton"
        Send-SeClick -Element $Send -SleepSeconds 5

        # process serial number list from webpage
        #$ResponseSerial = (Find-SeElement -Driver $Driver -By ClassName  "hpui-standalone-internal-link").text
        #$WebSerial = $ResponseSerial | Select-String -Pattern "SN: " | ForEach-Object { ($_ -Split "SN: ")[1] }

        # process warranty output from all tables on the web page
        $Response = (Get-SeElement -Driver $Driver -By ClassName  "hpui-standard-table").text

        # check response for anticipated conditions before closing the browser. Errors are thrown to calling function.
        if ($null -eq $Response) {
            # probable causes are 500 error or confirmation code request
            Stop-SeDriver -Target $Driver
            throw "Bad response from HPE warranty website"
        }
        if ($Response -match "Item Product serial number Country/Region of purchase Remove item" ) {
            Write-Warning 'At least one serial number not found, waiting 30 seconds before closing browser'
            Start-sleep -Seconds 30
            Stop-SeDriver -Target $Driver
            throw "One or more unrecognised serial numbers, use the -Ignore parameter to exclude them"
        }
        if ($Response.Count -ne $SerialNumber.Count) {
            Write-Warning 'At least one serial number has no warranty/support status,  waiting 30 seconds before closing browser'
            Start-Sleep -Seconds 30
            Stop-SeDriver -Target $Driver
            throw "Response count ($($Response.Count)) doesn't match the number of serial numbers ($($SerialNumber.Count))"
        }

        # close Browser
        Stop-SeDriver -Target $Driver

        for ($i = 0; $i -lt $SerialNumber.Count; $i++) {
            $ThisSerialNumber = $SerialNumber[$i]
            $RecWarranty = ($Response[$i] -split "`n" | Select-String -Pattern "Base Warranty") -split ' '
            $RecSupport = $Response[$i] -split "`n" | Select-String -Pattern "HPE Hardware Maintenance Onsite Support"

            if ($RecWarranty) {
                Write-Verbose "$RecWarranty"
                if ($RecWarranty[2] -eq $ThisSerialNumber) {
                    # Write-Verbose "Found base warranty and serial number matches $ThisSerialNumber"
                }
                else {
                    throw "Found base warranty, but the serial number $($RecWarranty[2]) doesn't match expected ($ThisSerialNumber)"
                }
                $WarrantyStatus = ($RecWarranty[-1]).Trim()
                $WarrantyEndDate = "$($RecWarranty[-3] -replace ',','')-$($RecWarranty[-4])-$($RecWarranty[-2])"
            }
            else {
                $WarrantyStatus = "Expired"
                $WarrantyEndDate = "Base warranty not found"
            }
            if ($RecSupport) {
                # older assets may have multiple (renewed)) support contracts.
                # We're assuming here that the latest support contract appears last from the web output which
                # appears to be the case. A safer way would be to pull out the dates and compare.
                If ($RecSupport.Count -gt 1) {
                    $RecSupport = ($RecSupport | Select-Object -Last 1) -split ' '
                }
                else {
                    $RecSupport = $RecSupport -split ' '
                }
                Write-Verbose "$RecSupport"
                $SupportStatus = ($RecSupport[-1]).Trim()
                $SupportEndDate = "$($RecSupport[-3] -replace ',','')-$($RecSupport[-4])-$($RecSupport[-2])"
            }
            else {
                $SupportStatus = "Expired"
                $SupportEndDate = "Support maintenance not found"
            }
            # write custom object to pipeline 
            [PSCustomObject]@{
                SerialNumber    = $ThisSerialNumber
                WarrantyStatus  = $WarrantyStatus
                WarrantyEndDate = $WarrantyEndDate
                SupportStatus   = $SupportStatus
                SupportEndDate  = $SupportEndDate
            }
        } # end for
    } # end Resolve-HPEwarranty

    # define a helper function to retry resolving serial numbers if a bad response is received from the
    # HPE warranty website
    Function Get-HPEwarranty {
        [CmdletBinding()]
        param (
            [System.Array]$SerialNumber
        )
        $RetryCount = 0
        do {
            try {
                # process this set of ten serial numbers and write results to pipeline
                Resolve-HPEwarranty -SerialNumber $SerialNumber -ErrorAction Stop
                
                # success, so break out of do/until
                break
            }
            catch {
                # web site 500 error or confirmation code
                if ($_.Exception -match 'Bad response') {
                    Write-Warning "$($_.Exception.Message), trying again ($(1+$RetryCount) of 4)..."
                    $RetryCount += 1
                }
                # If its not a web site issue, ignore this set of serial numbers and break out
                else {
                    Write-Warning $_.Exception.Message
                    # serial numbers wont be in the pipeline, this gets picked up by calling script
                    break
                }
            } # end try
        }
        until ($RetryCount -gt 4)
    } # end Get-HPEwarranty
}

process {
    $file = Get-ChildItem -Path $filename
    if (-not (test-path $file)) {
        throw "Input file not found"
    }

    if ($Ignore) {
        if (Test-Path $Ignore) {
            # if the ignore parameter is a file, grab the contents
            $Ignore = Get-Content $ignore
        }
        Write-Verbose "ignoring: $($Ignore -join ',')"
    }

    $AllServer = Import-Csv -Path $filename
    Write-Verbose "Checking $(($AllServer | Measure-Object).Count) serial numbers"

    $ServerList = $AllServer | Where-Object serialNumber -notin $Ignore
    $Count = 0
    $TotalCount = 0
    $SerialNumber = @()
    $Warranty = @()
    $Report = @()

    foreach ($Server in $ServerList) {
        # create a list of up to 10 serial numbers
        if ($Count -lt 10 ) {
            $SerialNumber += $Server.serialNumber
            $Count += 1
        }
        else {
            $TotalCount += $Count
            Write-Verbose "$($TotalCount - 9) to $TotalCount"

            # process this set of 10 serial numbers
            $Warranty += Get-HPEwarranty -SerialNumber $SerialNumber

            # setup for next set of srial numbers
            $Count = 1
            $SerialNumber = @()
            $SerialNumber += $Server.serialNumber
        }
    }

    if ($Count -gt 0) {
        $TotalCount += $Count
        Write-Verbose "$($TotalCount - $($Count-1)) to $TotalCount"

        # process any remaining serial numbers from a partial list
        $Warranty += Get-HPEwarranty -SerialNumber $SerialNumber
    }

    # now we have the warranty information, add this to the original collection
    foreach ($Server in $AllServer) {
        $ThisWarranty = $Warranty | Where-Object SerialNumber -eq $Server.serialNumber
    
        if ($Server.serialNumber -in $ignore) {
            # make sure the ignored serial numbers are added to the report
            $Server | Add-Member -NotePropertyName WarrantyStatus -NotePropertyValue "Serial number ignored by user"
            $Server | Add-Member -NotePropertyName WarrantyEndDate -NotePropertyValue 'Not found'
            $Server | Add-Member -NotePropertyName SupportStatus -NotePropertyValue 'Not found'
            $Server | Add-Member -NotePropertyName SupportEndDate -NotePropertyValue 'Not found'
        }
        else {
            if ($ThisWarranty) {
                # serial number was found, so add the details
                $Server | Add-Member -NotePropertyName WarrantyStatus -NotePropertyValue $ThisWarranty.WarrantyStatus
                $Server | Add-Member -NotePropertyName WarrantyEndDate -NotePropertyValue $ThisWarranty.WarrantyEndDate
                $Server | Add-Member -NotePropertyName SupportStatus -NotePropertyValue $ThisWarranty.SupportStatus
                $Server | Add-Member -NotePropertyName SupportEndDate -NotePropertyValue $ThisWarranty.SupportEndDate
            }
            else {
                # serial number wasn't found. This could be because the web site failed repeatedly, there is at
                # least one unrecognised serial number in a set, or at least one serial number in a set doesn't
                # have any warranty information (shows as no table on the website). In these cases, manually specify
                # problem serial numbers in the -ignore parameter (comma delimited list) 
                $Server | Add-Member -NotePropertyName WarrantyStatus -NotePropertyValue 'Unexpected error'
                $Server | Add-Member -NotePropertyName WarrantyEndDate -NotePropertyValue 'Unexpected error'
                $Server | Add-Member -NotePropertyName SupportStatus -NotePropertyValue 'Unexpected error'
                $Server | Add-Member -NotePropertyName SupportEndDate -NotePropertyValue 'Unexpected error'
            }
        }
        $Report += $Server
    }
}

end {
    # Write collection to a new CSV
    $NewFile = $file.Basename + "_new" + $file.Extension
    $Report | Export-Csv $NewFile
    $Warranty | Export-Csv .\Warranty.csv #to confirm
}