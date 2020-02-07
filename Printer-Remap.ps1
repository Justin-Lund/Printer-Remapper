######################################################### 
#             Printer & Driver Reset Tool v2            # 
#                Created By: Justin Lund                # 
#             https://github.com/Justin-Lund/           # 
######################################################### 


#--------------------------------------#
#            Initialization            #
#--------------------------------------#

# The location for the prndrvr.vbs script, which is used to manipulate the drivers, is in a different directory depending on the OS Version
$PrnDrvrPathXP = "C:\Windows\system32\prndrvr.vbs"
$PrnDrvrPath7 = "C:\Windows\system32\Printing_Admin_Scripts\en-US\prndrvr.vbs"
$PrnDrvrPath10 = "C:\Windows\system32\Printing_Admin_Scripts\en-US\prndrvr.vbs"
$ComputerOSVersion = Get-WMIObject -Class Win32_OperatingSystem

# Initialize Arrays
$PrinterList = @()
$DriverList = @()

# Output path for text file
$Path = "C:\Temp"

    # Create the directory if the path does not exist
    If(!(Test-Path $Path))
    {
          New-Item -ItemType Directory -Force -Path $Path | Out-Null

    }

# Output Display Divider
Function Divider {
    Write-Host ""
    Write-Host "------------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
}


#--------------------------------------#
#        Information Gathering         #
#--------------------------------------#

# Determine Default Printer
$DefaultPrinter = Get-WmiObject Win32_Printer | Where {$_.Default -eq $true}

# Get list of Network Printers
$NetworkPrinters = Get-WmiObject -Class Win32_Printer | Where {$_.Network -eq $true}

# Loop through the list of Network Printers
ForEach ($Printer in $NetworkPrinters){
    
    # Save Printer names into PrinterList array
    $PrinterList += $Printer.Name

    # Save Driver names into DriverList array
    $DriverList += $Printer.DriverName
}

# Remove duplicates from the arrays
$PrinterList = ($PrinterList | Select-Object -Unique)
$DriverList = ($DriverList | Select-Object -Unique)


#--------------------------------------#
#                Output                #
#--------------------------------------#

# Create a list of all printers separated onto new lines
$PrinterListText = $PrinterList -Join "`n"

# Save list of printers in a text file (in case of any issues)
Date | Out-File -Force -FilePath $Path\PrinterList.txt

Add-Content -Path $Path\PrinterList.txt -Value "
Default Printer:
$DefaultPrinter

Printers:
$PrinterListText
"

# List printers on screen
Write-Host ""
Write-Host "Default Printer:" -ForegroundColor Green
Write-Host $DefaultPrinter.Name
Write-Host ""
Write-Host "Printers:" -ForegroundColor Green
Write-Host $PrinterListText

Divider


#--------------------------------------#
#              Functional              #
#--------------------------------------#

# Delete all Network Printers
Write-Host "Deleting Network Printers" -ForegroundColor Red
Write-Host ""
Get-WmiObject Win32_Printer | Where {$_.Network -eq $true} | ForEach {$_.Delete()}


# Delete all Network Printer Drivers
ForEach ($Driver in $DriverList){

    # Stop Print Spooler
    (Get-WmiObject -Class Win32_Service -Filter "Name='Spooler'").StopService() | Out-Null
    Sleep 2

    # Clear Local Printer Cache
    Remove-Item -Path "C:\Windows\System32\Spool\PRINTERS\*"

    # Restart Print Spooler
    (Get-WmiObject -Class Win32_Service -Filter "Name='Spooler'").StartService() | Out-Null
    Sleep 2

    # Delete Printer Drivers

    # For WinXP Computers
    If ($ComputerOSVersion.Version.StartsWith(5)) {
        Write-Host "Deleting Driver: " -ForegroundColor Red -NoNewline
        Write-Host $Driver
        Invoke-Expression "cscript $PrnDrvrPathXP -d -m `"$Driver`" -v 3 -e `"Windows NT x86`""
    }

    # For Win7 Computers
    If ($ComputerOSVersion.Version.StartsWith(6)) {
        Write-Host "Deleting Driver: " -ForegroundColor Red -NoNewline
        Write-Host $Driver
        Invoke-Expression "cscript $PrnDrvrPath7 -d -m `"$Driver`" -v 3 -e `"Windows NT x86`""
    }

    # For Win10 Computers
    If ($ComputerOSVersion.Version.StartsWith(10)) {
        Write-Host "Deleting Driver: " -ForegroundColor Red -NoNewline
        Write-Host $Driver
        Invoke-Expression "cscript $PrnDrvrPath10 -d -m `"$Driver`" -v 3 -e `"Windows x64`""
    }
}

Divider

# Re-add Network Printers
ForEach ($Printer in $PrinterList){
    Write-Host "Readding " -ForegroundColor Green -NoNewline
    Write-Host $Printer
    (New-Object -ComObject WScript.Network).AddWindowsPrinterConnection("$Printer")
}

Write-Host ""

# Restore Default Printer
Write-Host "Restoring Default Printer" -ForegroundColor Green
$DefaultPrinter.SetDefaultPrinter() | Out-Null

Write-Host ""
Write-Host ""
Pause
