######################################################### 
#              Printer & Driver Reset Tool              # 
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
#              Functional              #
#--------------------------------------#

# Delete all Network Printers
Get-WmiObject Win32_Printer | Where {$_.Network -eq $true} | ForEach {$_.Delete()}

# Delete all Network Printer Drivers
ForEach ($Driver in $DriverList){

    # Restart Print Spooler
    (Get-WmiObject -Class Win32_Service -Filter "Name='Spooler'").StopService()
    Sleep 2
    (Get-WmiObject -Class Win32_Service -Filter "Name='Spooler'").StartService()
    Sleep 2

    # Delete Printer Drivers

    # For WinXP Computers
    If ($ComputerOSVersion.Version.StartsWith(5)) {
        Invoke-Expression "cscript $PrnDrvrPathXP -d -m `"$Driver`" -v 3 -e `"Windows NT x86`""
    }

    # For Win7 Computers
    If ($ComputerOSVersion.Version.StartsWith(6)) {
        Invoke-Expression "cscript $PrnDrvrPath7 -d -m `"$Driver`" -v 3 -e `"Windows NT x86`""
    }

    # For Win10 Computers
    If ($ComputerOSVersion.Version.StartsWith(10)) {
    Invoke-Expression "cscript $PrnDrvrPath10 -d -m `"$Driver`" -v 3 -e `"Windows x64`""
    }
}

# Re-add Network Printers
ForEach ($Printer in $PrinterList){
(New-Object -ComObject WScript.Network).AddWindowsPrinterConnection("$Printer")
}

# Restore Default Printer
$DefaultPrinter.SetDefaultPrinter()
