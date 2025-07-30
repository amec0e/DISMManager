#Requires -RunAsAdministrator

# DISM Manager Script
# Author: Amec0e
# Description: Menu-driven PowerShell script for DISM operations

# Global Variables
$Global:ExtractedISOPath = ""
$Global:InstallWimPath = ""
$Global:ISOOutputPath = ""
$Global:MountPath = ""
$Global:DriverPath = ""
$Global:PackagePath = ""
$Global:ForceUnsignedDrivers = $false
$Global:CurrentOperation = "None"
$Global:CleanupRequired = $false
$Global:EnabledFeatures = @()
$Global:DisabledFeatures = @()
$Global:AddedPackages = @()
$Global:RemovedPackages = @()
$Global:RemovedDrivers = @()
$Global:RemovedAppxPackages = @()
$Global:AddedAppxPackages = @()
$Global:AppxPackagePath = ""
$Global:BackupPath = ""
$Global:WimlibPath = ""
$Global:DesktopFilesPath = ""
$Global:RemovedDesktopFiles = @()
$Global:SelectedDesktopFiles = @()
$Global:ShouldAddDesktopFiles = $false
# Generic Volume License Keys (RTMs) - RTM Generic Digital Keys
$Global:RTMKeys = @{
    "Windows 11 Home" = "YTMG3-N6DKC-DKB77-7M9GH-8HVX7"
    "Windows 11 Home N" = "4CPRK-NM3K3-X6XXQ-RXX86-WXCHW"
    "Windows 11 Home Single Language" = "BT79Q-G7N6G-PGBYW-4YWX6-6F4BT"
    "Windows 11 Home Country Specific" = "N2434-X9D7W-8PF6X-8DV9T-8TYMD"
    "Windows 11 Pro" = "VK7JG-NPHTM-C97JM-9MPGT-3V66T"
    "Windows 11 Pro N" = "2B87N-8KFHP-DKV6R-Y2C8J-PKCKT"
    "Windows 11 Pro for Workstations" = "DXG7C-N36C4-C4HTG-X4T3X-2YV77"
    "Windows 11 Pro N for Workstations" = "WYPNQ-8C467-V2W6J-TX4WX-WT2RQ"
    "Windows 11 Pro Education" = "8PTT6-RNW4C-6V7J2-C2D3X-MHBPB"
    "Windows 11 Pro Education N" = "GJTYN-HDMQY-FRR76-HVGC7-QPF8P"
    "Windows 11 Education" = "YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY"
    "Windows 11 Education N" = "84NGF-MHBT6-FXBX8-QWJK7-DRR8H"
    "Windows 11 Enterprise" = "XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"
    "Windows 11 Enterprise N" = "WGGHN-J84D6-QYCPR-T7PJ7-X766F"
    "Windows 11 Enterprise G N" = "FW7NV-4T673-HF4VX-9X4MM-B4H4T"
    "Windows 11 Enterprise Evaluation" = "XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"
    "Windows 11 Enterprise N Evaluation" = "WGGHN-J84D6-QYCPR-T7PJ7-X766F"
    "Windows 10 Home" = "YTMG3-N6DKC-DKB77-7M9GH-8HVX7"
    "Windows 10 Home N" = "4CPRK-NM3K3-X6XXQ-RXX86-WXCHW"
    "Windows 10 Home Single Language" = "BT79Q-G7N6G-PGBYW-4YWX6-6F4BT"
    "Windows 10 Pro" = "VK7JG-NPHTM-C97JM-9MPGT-3V66T"
    "Windows 10 Pro N" = "2B87N-8KFHP-DKV6R-Y2C8J-PKCKT"
    "Windows 10 Pro for Workstations" = "DXG7C-N36C4-C4HTG-X4T3X-2YV77"
    "Windows 10 Pro N for Workstations" = "WYPNQ-8C467-V2W6J-TX4WX-WT2RQ"
    "Windows 10 Education" = "YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY"
    "Windows 10 Education N" = "84NGF-MHBT6-FXBX8-QWJK7-DRR8H"
    "Windows 10 Pro Education" = "8PTT6-RNW4C-6V7J2-C2D3X-MHBPB"
    "Windows 10 Pro Education N" = "GJTYN-HDMQY-FRR76-HVGC7-QPF8P"
    "Windows 10 Enterprise" = "XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"
    "Windows 10 Enterprise G N" = "FW7NV-4T673-HF4VX-9X4MM-B4H4T"
    "Windows 10 Enterprise N" = "WGGHN-J84D6-QYCPR-T7PJ7-X766F"
    "Windows 10 Enterprise S" = "NK96Y-D9CD8-W44CQ-R8YTK-DYJWX"
    "Windows 10 Enterprise N LTSB 2016" = "RW7WN-FMT44-KRGBK-G44WK-QV7YK"
    "Windows 10 Enterprise Evaluation" = "XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"
    "Windows 10 Enterprise N Evaluation" = "WGGHN-J84D6-QYCPR-T7PJ7-X766F"
}
# KMS Client Product Keys (Official Microsoft List) - Separated for Easy Matching
$Global:KMSClientKeys = @{
    # Windows 11 Client Editions
    "Windows 11 Pro" = "W269N-WFGWX-YVC9B-4J6C9-T83GX"
    "Windows 11 Pro N" = "MH37W-N47XK-V7XM9-C7227-GCQG9"
    "Windows 11 Pro for Workstations" = "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J"
    "Windows 11 Pro N for Workstations" = "9FNHH-K3HBT-3W4TD-6383H-6XYWF"
    "Windows 11 Pro Education" = "6TP4R-GNPTD-KYYHQ-7B7DP-J447Y"
    "Windows 11 Pro Education N" = "YVWGF-BXNMC-HTQYQ-CPQ99-66QFC"
    "Windows 11 Education" = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"
    "Windows 11 Education N" = "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"
    "Windows 11 Enterprise" = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
    "Windows 11 Enterprise N" = "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"
    "Windows 11 Enterprise G" = "YYVX9-NTFWV-6MDM3-9PT4T-4M68B"
    "Windows 11 Enterprise G N" = "44RPN-FTY23-9VTTB-MP9BX-T84FV"
    "Windows 11 Enterprise LTSC 2019" = "M7XTQ-FN8P6-TTKYV-9D4CC-J462D"
    "Windows 11 Enterprise LTSC 2021" = "M7XTQ-FN8P6-TTKYV-9D4CC-J462D"
    "Windows 11 Enterprise LTSC 2024" = "M7XTQ-FN8P6-TTKYV-9D4CC-J462D"
    "Windows 11 Enterprise N LTSC 2019" = "92NFX-8DJQP-P6BBQ-THF9C-7CG2H"
    "Windows 11 Enterprise N LTSC 2021" = "92NFX-8DJQP-P6BBQ-THF9C-7CG2H"
    "Windows 11 Enterprise N LTSC 2024" = "92NFX-8DJQP-P6BBQ-THF9C-7CG2H"
    "Windows 11 Enterprise Evaluation" = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
    "Windows 11 Enterprise N Evaluation" = "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"
    
    # Windows 10 Client Editions
    "Windows 10 Pro" = "W269N-WFGWX-YVC9B-4J6C9-T83GX"
    "Windows 10 Pro N" = "MH37W-N47XK-V7XM9-C7227-GCQG9"
    "Windows 10 Pro for Workstations" = "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J"
    "Windows 10 Pro N for Workstations" = "9FNHH-K3HBT-3W4TD-6383H-6XYWF"
    "Windows 10 Pro Education" = "6TP4R-GNPTD-KYYHQ-7B7DP-J447Y"
    "Windows 10 Pro Education N" = "YVWGF-BXNMC-HTQYQ-CPQ99-66QFC"
    "Windows 10 Education" = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"
    "Windows 10 Education N" = "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"
    "Windows 10 Enterprise" = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
    "Windows 10 Enterprise N" = "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"
    "Windows 10 Enterprise G" = "YYVX9-NTFWV-6MDM3-9PT4T-4M68B"
    "Windows 10 Enterprise G N" = "44RPN-FTY23-9VTTB-MP9BX-T84FV"
    "Windows 10 Enterprise LTSC 2019" = "M7XTQ-FN8P6-TTKYV-9D4CC-J462D"
    "Windows 10 Enterprise LTSC 2021" = "M7XTQ-FN8P6-TTKYV-9D4CC-J462D"
    "Windows 10 Enterprise LTSC 2024" = "M7XTQ-FN8P6-TTKYV-9D4CC-J462D"
    "Windows 10 Enterprise N LTSC 2019" = "92NFX-8DJQP-P6BBQ-THF9C-7CG2H"
    "Windows 10 Enterprise N LTSC 2021" = "92NFX-8DJQP-P6BBQ-THF9C-7CG2H"
    "Windows 10 Enterprise N LTSC 2024" = "92NFX-8DJQP-P6BBQ-THF9C-7CG2H"
    
    # Windows IoT Enterprise
    "Windows IoT Enterprise LTSC 2021" = "KBN8V-HFGQ4-MGXVD-347P6-PDQGT"
    "Windows IoT Enterprise LTSC 2024" = "KBN8V-HFGQ4-MGXVD-347P6-PDQGT"
    
    # Windows Server editions
    "Windows Server 2025 Standard" = "TVRH6-WHNXV-R9WG3-9XRFY-MY832"
    "Windows Server 2025 Datacenter" = "D764K-2NDRG-47T6Q-P8T8W-YP6DF"
    "Windows Server 2025 Datacenter: Azure Edition" = "XGN3F-F394H-FD2MY-PP6FD-8MCRC"
    "Windows Server 2022 Standard" = "VDYBN-27WPP-V4HQT-9VMD4-VMK7H"
    "Windows Server 2022 Datacenter" = "WX4NM-KYWYW-QJJR4-XV3QB-6VM33"
    "Windows Server 2022 Datacenter: Azure Edition" = "NTBV8-9K7Q8-V27C6-M2BTV-KHMXV"
    "Windows Server 2019 Standard" = "N69G4-B89J2-4G8F4-WWYCC-J464C"
    "Windows Server 2019 Datacenter" = "WMDGN-G9PQG-XVVXX-R3X43-63DFG"
    "Windows Server 2019 Essentials" = "WVDHN-86M7X-466P6-VHXV7-YY726"
    "Windows Server 2016 Standard" = "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY"
    "Windows Server 2016 Datacenter" = "CB7KF-BWN84-R7R2Y-793K2-8XDDG"
    "Windows Server 2016 Essentials" = "JCKRF-N37P4-C2D82-9YXRT-4M63B"
}

$null = Register-EngineEvent PowerShell.Exiting -Action {
    Write-Host "`nGraceful shutdown initiated..." -ForegroundColor Yellow
    Invoke-Cleanup
}

function Test-SystemRequirements {
    $requirements = @{
        PowerShell = $false
        OSCDIMG = $false
        ADK = $false
    }
    
    if ($PSVersionTable.PSVersion.Major -ge 5) {
        $requirements.PowerShell = $true
    }
    
    $oscdimgPath = Get-OSCDImgPath
    if ($oscdimgPath) {
        $requirements.OSCDIMG = $true
        if ($oscdimgPath -like "*Windows Kits*") {
            $requirements.ADK = $true
        }
    }
    
    return $requirements
}

function Test-RequiredPaths {
    return ($Global:ExtractedISOPath -and $Global:InstallWimPath -and $Global:MountPath)
}

function Test-MountPathEmpty {
    if (-not $Global:MountPath -or -not (Test-Path $Global:MountPath)) {
        return $true
    }
    
    $items = Get-ChildItem $Global:MountPath -Force -ErrorAction SilentlyContinue
    return ($items.Count -eq 0)
}

function Invoke-Cleanup {
    if ($Global:CleanupRequired -and $Global:MountPath -and (Test-Path $Global:MountPath)) {
        Write-Host "Performing cleanup - unmounting any mounted images..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Cleanup completed." -ForegroundColor Green
        }
    }
}

function Show-Menu {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "         DISM Manager" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    $sysReq = Test-SystemRequirements
    Write-Host "SYSTEM STATUS:" -ForegroundColor Magenta
    Write-Host "PowerShell $($PSVersionTable.PSVersion): " -NoNewline
    Write-Host $(if ($sysReq.PowerShell) { "OK" } else { "NEEDS UPDATE" }) -ForegroundColor $(if ($sysReq.PowerShell) { "Green" } else { "Red" })
    Write-Host "OSCDIMG: " -NoNewline
    Write-Host $(if ($sysReq.OSCDIMG) { "OK" } else { "NOT FOUND" }) -ForegroundColor $(if ($sysReq.OSCDIMG) { "Green" } else { "Red" })
    Write-Host "Windows ADK: " -NoNewline
    Write-Host $(if ($sysReq.ADK) { "OK" } else { "NOT DETECTED" }) -ForegroundColor $(if ($sysReq.ADK) { "Green" } else { "Red" })
    Write-Host "Wimlib: " -NoNewline
    $wimlibAvailable = $Global:WimlibPath -and (Test-Path (Join-Path $Global:WimlibPath "wimlib-imagex.exe"))
    if ($wimlibAvailable) {
        Write-Host "OK (Advanced features available)" -ForegroundColor Green
    } else {
        Write-Host "NOT CONFIGURED (Limited functionality)" -ForegroundColor Yellow
    }
    Write-Host ""
    
    Write-Host "CONFIGURATION:" -ForegroundColor Magenta
    Write-Host "Current ISO Path: " -NoNewline
    if ($Global:ExtractedISOPath) {
        Write-Host $Global:ExtractedISOPath -ForegroundColor Green
    } else {
        Write-Host "Not Set" -ForegroundColor Red
    }
    Write-Host "Install Image Path: " -NoNewline
    if ($Global:InstallWimPath) {
        Write-Host $Global:InstallWimPath -ForegroundColor Green
    } else {
        Write-Host "Not Found" -ForegroundColor Red
    }
    Write-Host "Mount Path: " -NoNewline
    if ($Global:MountPath) {
        $mountEmpty = Test-MountPathEmpty
        if ($mountEmpty) {
            Write-Host "${Global:MountPath} (Empty)" -ForegroundColor Green
        } else {
            Write-Host "${Global:MountPath} (IN USE - May need cleanup)" -ForegroundColor Red
        }
    } else {
        Write-Host "Not Set" -ForegroundColor Red
    }
    Write-Host "Driver Path: " -NoNewline
    if ($Global:DriverPath) {
        Write-Host $Global:DriverPath -ForegroundColor Green
    } else {
        Write-Host "Not Set" -ForegroundColor Yellow
    }
    Write-Host "ISO Output Path: " -NoNewline
    if ($Global:ISOOutputPath) {
        Write-Host $Global:ISOOutputPath -ForegroundColor Green
    } else {
        Write-Host "Not Set" -ForegroundColor Yellow
    }
    Write-Host "Package Path: " -NoNewline
    if ($Global:PackagePath) {
        Write-Host $Global:PackagePath -ForegroundColor Green
    } else {
        Write-Host "Not Set" -ForegroundColor Yellow
    }
    Write-Host "Desktop Files Path: " -NoNewline
    if ($Global:DesktopFilesPath) {
        Write-Host $Global:DesktopFilesPath -ForegroundColor Green
    } else {
        Write-Host "Not Set" -ForegroundColor Yellow
    }
    Write-Host "Backup Path: " -NoNewline
    if ($Global:BackupPath) {
        Write-Host $Global:BackupPath -ForegroundColor Green
    } else {
        Write-Host "Not Set" -ForegroundColor Yellow
    }
    Write-Host "Wimlib Path: " -NoNewline
    if ($wimlibAvailable) {
        Write-Host $Global:WimlibPath -ForegroundColor Green
    } else {
        Write-Host "Not Set (Setup in Configuration)" -ForegroundColor Yellow
    }
    Write-Host "Force Unsigned Drivers: " -NoNewline
    if ($Global:ForceUnsignedDrivers) {
        Write-Host "Enabled" -ForegroundColor Green
    } else {
        Write-Host "Disabled" -ForegroundColor Red
    }
    Write-Host ""
    
    $requiredPathsSet = Test-RequiredPaths
    $mountPathEmpty = Test-MountPathEmpty
    
    if (-not $requiredPathsSet) {
        Write-Host "WARNING: Required paths must be set before using most options!" -ForegroundColor Red
        Write-Host ""
    }
    
    if (-not $mountPathEmpty) {
        Write-Host "WARNING: Mount path is not empty! Use 'Cleanup/Unmount' option if needed." -ForegroundColor Red
        Write-Host ""
    }
    
    if (-not $wimlibAvailable) {
        Write-Host "INFO: Setup wimlib in Configuration for advanced metadata features." -ForegroundColor Cyan
        Write-Host ""
    }
    
    Write-Host "MENU OPTIONS:" -ForegroundColor Cyan
    Write-Host "1. Configuration Settings" -ForegroundColor Yellow
    Write-Host "2. Manage Drivers (Add/Remove/Export)" -ForegroundColor Yellow
    Write-Host "3. Manage Packages and Features" -ForegroundColor $(if ($requiredPathsSet -and $mountPathEmpty) { "Yellow" } else { "DarkGray" })
    Write-Host "4. Manage Indexes (Add/Remove)" -ForegroundColor $(if ($requiredPathsSet -and $mountPathEmpty) { "Yellow" } else { "DarkGray" })
    Write-Host "5. Manage Editions" -ForegroundColor $(if ($requiredPathsSet -and $mountPathEmpty) { "Yellow" } else { "DarkGray" })
    Write-Host "6. File Conversion (WIM/ESD)" -ForegroundColor $(if ($requiredPathsSet -and $mountPathEmpty) { "Yellow" } else { "DarkGray" })
    Write-Host "7. Create ISO (UEFI/BIOS/Hybrid)" -ForegroundColor $(if ($requiredPathsSet) { "Yellow" } else { "DarkGray" })
    Write-Host "8. View WIM/ESD Information" -ForegroundColor $(if ($requiredPathsSet) { "Yellow" } else { "DarkGray" })
    Write-Host "9. Manage Desktop Files" -ForegroundColor $(if ($requiredPathsSet -and $mountPathEmpty) { "Yellow" } else { "DarkGray" })
    Write-Host "10. Manage Registry" -ForegroundColor $(if ($requiredPathsSet -and $mountPathEmpty) { "Yellow" } else { "DarkGray" })
    Write-Host "11. Mount Index" -ForegroundColor $(if ($requiredPathsSet -and $mountPathEmpty) { "Yellow" } else { "DarkGray" })
    Write-Host "12. Cleanup/Unmount" -ForegroundColor $(if ($Global:MountPath -and -not $mountPathEmpty) { "Yellow" } else { "DarkGray" })
    Write-Host "13. Exit" -ForegroundColor Red
    Write-Host ""
    Write-Host "Select an option (1-13): " -NoNewline -ForegroundColor White
}

function Manage-Configuration {
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "    Configuration Settings" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "CURRENT CONFIGURATION:" -ForegroundColor Magenta
        Write-Host "Extracted ISO Path: " -NoNewline
        if ($Global:ExtractedISOPath) {
            Write-Host $Global:ExtractedISOPath -ForegroundColor Green
        } else {
            Write-Host "Not Set" -ForegroundColor Red
        }
        
        Write-Host "Install Image Path: " -NoNewline
        if ($Global:InstallWimPath) {
            Write-Host $Global:InstallWimPath -ForegroundColor Green
        } else {
            Write-Host "Not Found" -ForegroundColor Red
        }
        
        Write-Host "Mount Path: " -NoNewline
        if ($Global:MountPath) {
            $mountEmpty = Test-MountPathEmpty
            if ($mountEmpty) {
                Write-Host "${Global:MountPath} (Empty)" -ForegroundColor Green
            } else {
                Write-Host "${Global:MountPath} (IN USE - May need cleanup)" -ForegroundColor Red
            }
        } else {
            Write-Host "Not Set" -ForegroundColor Red
        }
        
        Write-Host "Driver Path: " -NoNewline
        if ($Global:DriverPath) {
            Write-Host $Global:DriverPath -ForegroundColor Green
        } else {
            Write-Host "Not Set" -ForegroundColor Yellow
        }
        
        Write-Host "ISO Output Path: " -NoNewline
        if ($Global:ISOOutputPath) {
            Write-Host $Global:ISOOutputPath -ForegroundColor Green
        } else {
            Write-Host "Not Set" -ForegroundColor Yellow
        }
        
        Write-Host "Package Path: " -NoNewline
        if ($Global:PackagePath) {
            Write-Host $Global:PackagePath -ForegroundColor Green
        } else {
            Write-Host "Not Set" -ForegroundColor Yellow
        }
        
        Write-Host "Desktop Files Path: " -NoNewline
        if ($Global:DesktopFilesPath) {
            Write-Host $Global:DesktopFilesPath -ForegroundColor Green
        } else {
            Write-Host "Not Set" -ForegroundColor Yellow
        }
        
        Write-Host "Backup Path: " -NoNewline
        if ($Global:BackupPath) {
            Write-Host $Global:BackupPath -ForegroundColor Green
        } else {
            Write-Host "Not Set" -ForegroundColor Yellow
        }
        
        Write-Host "Wimlib Path: " -NoNewline
        $wimlibAvailable = $Global:WimlibPath -and (Test-Path (Join-Path $Global:WimlibPath "wimlib-imagex.exe"))
        if ($wimlibAvailable) {
            Write-Host $Global:WimlibPath -ForegroundColor Green
        } else {
            Write-Host "Not Set (Advanced features unavailable)" -ForegroundColor Red
        }
        
        Write-Host "Force Unsigned Drivers: " -NoNewline
        if ($Global:ForceUnsignedDrivers) {
            Write-Host "Enabled" -ForegroundColor Green
        } else {
            Write-Host "Disabled" -ForegroundColor Red
        }
        Write-Host ""
        
        $requiredPathsSet = Test-RequiredPaths
        if (-not $requiredPathsSet) {
            Write-Host "WARNING: Required paths must be set before using other options!" -ForegroundColor Red
            Write-Host ""
        }
        
        Write-Host "CONFIGURATION OPTIONS:" -ForegroundColor Cyan
        Write-Host "1. Set Extracted ISO Path" -ForegroundColor Yellow
        Write-Host "2. Set Mount Path" -ForegroundColor Yellow
        Write-Host "3. Set ISO Output Path" -ForegroundColor $(if ($requiredPathsSet) { "Yellow" } else { "DarkGray" })
        Write-Host "4. Set Driver Path" -ForegroundColor $(if ($requiredPathsSet) { "Yellow" } else { "DarkGray" })
        Write-Host "5. Set Package Path" -ForegroundColor $(if ($requiredPathsSet) { "Yellow" } else { "DarkGray" })
        Write-Host "6. Set Backup Path" -ForegroundColor $(if ($requiredPathsSet) { "Yellow" } else { "DarkGray" })
        Write-Host "7. Set Desktop Files Path" -ForegroundColor $(if ($requiredPathsSet) { "Yellow" } else { "DarkGray" })
        Write-Host "8. Setup Wimlib (Advanced Features)" -ForegroundColor Cyan
        Write-Host "9. Toggle Force Unsigned Drivers" -ForegroundColor $(if ($requiredPathsSet) { "Yellow" } else { "DarkGray" })
        Write-Host "10. Return to Main Menu" -ForegroundColor Red
        Write-Host ""
        Write-Host "Select an option (1-10): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { Set-ISOPath }
            "2" { Set-MountPath }
            "3" { 
                if (-not $requiredPathsSet) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Set-ISOOutputPath
                }
            }
            "4" { 
                if (-not $requiredPathsSet) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Set-DriverPath
                }
            }
            "5" { 
                if (-not $requiredPathsSet) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Set-PackagePath
                }
            }
            "6" { 
                if (-not $requiredPathsSet) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Set-ConversionBackupPath
                }
            }
            "7" { 
                if (-not $requiredPathsSet) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Set-DesktopFilesPath
                }
            }
            "8" { Manage-WimlibSetup }
            "9" { 
                if (-not $requiredPathsSet) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Toggle-ForceUnsignedDrivers
                }
            }
            "10" { return }
            default { 
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

function Test-DirectoryPath {
    param(
        [string]$Path,
        [string]$PathType
    )
    
    $Path = $Path.Trim('"', "'")
    $expandedPath = [System.Environment]::ExpandEnvironmentVariables($Path)
    
    if (Test-Path $expandedPath) {
        try {
            $actualPath = (Get-Item $expandedPath).FullName
            if ($actualPath -cne $expandedPath) {
                Write-Host "Path found (corrected casing): ${actualPath}" -ForegroundColor Cyan
            }
            
            return $actualPath
        } catch {
            Write-Host "Warning: Could not resolve exact path casing, using input as-is" -ForegroundColor Yellow
            return $expandedPath
        }
    } else {
        Write-Host "$PathType path does not exist: ${expandedPath}" -ForegroundColor Red
        return $null
    }
}

function Set-ISOPath {
    $path = Read-Host "Enter the path to your extracted ISO folder"
    
    $validatedPath = Test-DirectoryPath -Path $path -PathType "ISO"
    if ($validatedPath) {
        $wimPath = Join-Path $validatedPath "sources\install.wim"
        $esdPath = Join-Path $validatedPath "sources\install.esd"
        
        $wimExists = Test-Path $wimPath
        $esdExists = Test-Path $esdPath
        
        if ($wimExists -and $esdExists) {
            Write-Host "`nBoth install.wim and install.esd found!" -ForegroundColor Yellow
            Write-Host "WIM file: $wimPath" -ForegroundColor Gray
            Write-Host "ESD file: $esdPath" -ForegroundColor Gray
            Write-Host ""
            Write-Host "1. Use install.wim (Default)" -ForegroundColor Green
            Write-Host "2. Use install.esd" -ForegroundColor Cyan
            
            do {
                $choice = Read-Host "Select which file to use (1-2, default: 1)"
                if (-not $choice) { $choice = "1" }
                
                switch ($choice) {
                    "1" {
                        $Global:ExtractedISOPath = $validatedPath
                        $Global:InstallWimPath = $wimPath
                        Write-Host "Successfully set ISO path and selected install.wim" -ForegroundColor Green
                        break
                    }
                    "2" {
                        $Global:ExtractedISOPath = $validatedPath
                        $Global:InstallWimPath = $esdPath
                        Write-Host "Successfully set ISO path and selected install.esd" -ForegroundColor Green
                        break
                    }
                    default {
                        Write-Host "Invalid selection. Please choose 1 or 2." -ForegroundColor Red
                        continue
                    }
                }
                break
            } while ($true)
            
        } elseif ($wimExists) {
            $Global:ExtractedISOPath = $validatedPath
            $Global:InstallWimPath = $wimPath
            Write-Host "Successfully set ISO path and found install.wim" -ForegroundColor Green
        } elseif ($esdExists) {
            $Global:ExtractedISOPath = $validatedPath
            $Global:InstallWimPath = $esdPath
            Write-Host "Successfully set ISO path and found install.esd" -ForegroundColor Green
        } else {
            Write-Host "Neither install.wim nor install.esd found in sources folder. Please check the path." -ForegroundColor Red
            Write-Host "Expected locations:" -ForegroundColor Gray
            Write-Host "  - $wimPath" -ForegroundColor Gray
            Write-Host "  - $esdPath" -ForegroundColor Gray
        }
    } else {
        Write-Host "Please provide a valid existing directory path when ready." -ForegroundColor Red
    }
    
    Read-Host "Press Enter to continue"
}

function Set-MountPath {
    $path = Read-Host "Enter the mount path for WIM/ESD operations"
    
    $validatedPath = Test-DirectoryPath -Path $path -PathType "Mount"
    if ($validatedPath) {
        $Global:MountPath = $validatedPath
        Write-Host "Mount path set to: ${validatedPath}" -ForegroundColor Green
    } else {
        Write-Host "Please provide a valid existing directory path when ready." -ForegroundColor Red
    }
    
    Read-Host "Press Enter to continue"
}

function Set-DriverPath {
    $currentPath = if ($Global:DriverPath) { " (Current: $Global:DriverPath)" } else { "" }
    $path = Read-Host "Enter the path to your drivers folder${currentPath}"
    
    if (-not $path -and $Global:DriverPath) {
        Write-Host "Keeping current driver path: $Global:DriverPath" -ForegroundColor Green
    } else {
        $validatedPath = Test-DirectoryPath -Path $path -PathType "Driver"
        if ($validatedPath) {
            $Global:DriverPath = $validatedPath
            Write-Host "Driver path set to: ${validatedPath}" -ForegroundColor Green
        } else {
            Write-Host "Please provide a valid existing directory path when ready." -ForegroundColor Red
        }
    }
    
    Read-Host "Press Enter to continue"
}

function Set-PackagePath {
    $currentPath = if ($Global:PackagePath) { " (Current: $Global:PackagePath)" } else { "" }
    $path = Read-Host "Enter the path to your packages folder${currentPath}"
    
    if (-not $path -and $Global:PackagePath) {
        Write-Host "Keeping current package path: $Global:PackagePath" -ForegroundColor Green
    } else {
        $validatedPath = Test-DirectoryPath -Path $path -PathType "Package"
        if ($validatedPath) {
            $Global:PackagePath = $validatedPath
            Write-Host "Package path set to: ${validatedPath}" -ForegroundColor Green
        } else {
            Write-Host "Please provide a valid existing directory path when ready." -ForegroundColor Red
        }
    }
    
    Read-Host "Press Enter to continue"
}

function Set-DesktopFilesPath {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nDesktop Files Path Configuration" -ForegroundColor Cyan
    Write-Host "This path will contain files that you want to copy to desktop folders in Windows images." -ForegroundColor Gray
    Write-Host ""
    
    $currentPath = if ($Global:DesktopFilesPath) { " (Current: $Global:DesktopFilesPath)" } else { "" }
    $path = Read-Host "Enter the path to your desktop files folder${currentPath}"
    
    if (-not $path -and $Global:DesktopFilesPath) {
        Write-Host "Keeping current desktop files path: $Global:DesktopFilesPath" -ForegroundColor Green
    } else {
        $validatedPath = Test-DirectoryPath -Path $path -PathType "Desktop Files"
        if ($validatedPath) {
            $Global:DesktopFilesPath = $validatedPath
            Write-Host "Desktop files path set to: ${validatedPath}" -ForegroundColor Green
        } else {
            Write-Host "Please provide a valid existing directory path when ready." -ForegroundColor Red
        }
    }
    
    Read-Host "Press Enter to continue"
}

function Set-ISOOutputPath {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $currentPath = if ($Global:ISOOutputPath) { " (Current: $Global:ISOOutputPath)" } else { "" }
    $path = Read-Host "Enter the output path for the ISO file (including filename.iso)${currentPath}"
    
    if (-not $path -and $Global:ISOOutputPath) {
        Write-Host "Keeping current ISO output path: $Global:ISOOutputPath" -ForegroundColor Green
    } else {
        if ($path) {
            $path = $path.Trim('"', "'")
            $expandedPath = [System.Environment]::ExpandEnvironmentVariables($path)
            
            $directory = Split-Path $expandedPath -Parent
            if ($directory -and (Test-Path $directory)) {
                $Global:ISOOutputPath = $expandedPath
                Write-Host "ISO output path set to: ${expandedPath}" -ForegroundColor Green
            } else {
                Write-Host "Output directory does not exist: ${directory}" -ForegroundColor Red
                Write-Host "Please provide a valid path when ready." -ForegroundColor Red
            }
        } else {
            Write-Host "Please provide a valid path when ready." -ForegroundColor Red
        }
    }
    
    Read-Host "Press Enter to continue"
}

function Set-ConversionBackupPath {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nBackup Path Configuration" -ForegroundColor Cyan
    Write-Host "This path will be used as the default location for backup files during WIM/ESD conversions." -ForegroundColor Gray
    Write-Host ""
    
    $currentPath = if ($Global:BackupPath) { " (Current: $Global:BackupPath)" } else { "" }
    $path = Read-Host "Enter the path for backup files${currentPath}"
    
    if (-not $path -and $Global:BackupPath) {
        Write-Host "Keeping current Backup Path: $Global:BackupPath" -ForegroundColor Green
    } else {
        $validatedPath = Test-DirectoryPath -Path $path -PathType "Conversion Backup"
        if ($validatedPath) {
            $Global:BackupPath = $validatedPath
            Write-Host "Backup Path set to: ${validatedPath}" -ForegroundColor Green
            Write-Host "This will be used as the default backup location for WIM/ESD conversions." -ForegroundColor Cyan
        } else {
            Write-Host "Please provide a valid existing directory path when ready." -ForegroundColor Red
        }
    }
    
    Read-Host "Press Enter to continue"
}

function Toggle-ForceUnsignedDrivers {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $Global:ForceUnsignedDrivers = -not $Global:ForceUnsignedDrivers
    $status = if ($Global:ForceUnsignedDrivers) { "Enabled" } else { "Disabled" }
    $color = if ($Global:ForceUnsignedDrivers) { "Green" } else { "Red" }
    Write-Host "Force Unsigned Drivers: $status" -ForegroundColor $color
    
    Read-Host "Press Enter to continue"
}

function Invoke-ManualCleanup {
    if (-not $Global:MountPath) {
        Write-Host "No mount path set." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-Path $Global:MountPath)) {
        Write-Host "Mount path does not exist: $Global:MountPath" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $mountEmpty = Test-MountPathEmpty
    if ($mountEmpty) {
        Write-Host "Mount path is already empty." -ForegroundColor Green
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "Attempting to unmount any mounted images..." -ForegroundColor Yellow
    Write-Host "Mount path: $Global:MountPath" -ForegroundColor Gray
    
    try {
        dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Successfully unmounted and discarded changes." -ForegroundColor Green
        } else {
            Write-Host "DISM unmount failed. Trying alternative cleanup..." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error during cleanup: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    $mountEmpty = Test-MountPathEmpty
    if ($mountEmpty) {
        Write-Host "Mount path is now clean." -ForegroundColor Green
    } else {
        Write-Host "Mount path may still contain files. Manual intervention might be needed." -ForegroundColor Yellow
    }
    
    Read-Host "Press Enter to continue"
}

function Get-WimInfo {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "WIM/ESD Information:" -ForegroundColor Cyan
    Write-Host "=================" -ForegroundColor Cyan
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
    } catch {
        Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Read-Host "Press Enter to continue"
}

function Add-DriversToWim {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $driverPath = $Global:DriverPath
    if (-not $driverPath) {
        do {
            $path = Read-Host "Enter the path to your drivers folder"
            $validatedPath = Test-DirectoryPath -Path $path -PathType "Driver"
            if ($validatedPath) {
                $driverPath = $validatedPath
                $save = Read-Host "Save this driver path for future use? (Y/n)"
                if ($save.ToLower() -ne 'n') {
                    $Global:DriverPath = $driverPath
                }
                break
            } else {
                Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
            }
        } while ($true)
    } else {
        Write-Host "Using saved driver path: $driverPath" -ForegroundColor Green
        $useExisting = Read-Host "Use this path? (Y/n)"
        if ($useExisting.ToLower() -eq 'n') {
            do {
                $path = Read-Host "Enter the path to your drivers folder"
                $validatedPath = Test-DirectoryPath -Path $path -PathType "Driver"
                if ($validatedPath) {
                    $driverPath = $validatedPath
                    break
                } else {
                    Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
                }
            } while ($true)
        }
    }
    
    Write-Host "`nDriver Installation Options:" -ForegroundColor Cyan
    Write-Host "1. Add to specific index"
    Write-Host "2. Add to all indexes"
    $choice = Read-Host "Select option (1-2)"
    
    $Global:CurrentOperation = "Adding Drivers"
    $Global:CleanupRequired = $true
    
    try {
        if ($choice -eq "1") {
            Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
            try {
                dism /Get-WimInfo /WimFile:$Global:InstallWimPath
                if ($LASTEXITCODE -ne 0) {
                    Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                    return
                }
            } catch {
                Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
                Read-Host "Press Enter to continue"
                return
            }
            
            $indexNumber = Read-Host "`nEnter the index number"
            Add-DriversToIndex -IndexNumber $indexNumber -DriverPath $driverPath
        } elseif ($choice -eq "2") {
            $indexes = Get-WimIndexes
            foreach ($index in $indexes) {
                Add-DriversToIndex -IndexNumber $index -DriverPath $driverPath
            }
        } else {
            Write-Host "Invalid selection." -ForegroundColor Red
        }
    } catch {
        Write-Host "Error during driver installation: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
    
    Read-Host "Press Enter to continue"
}

function Get-WimIndexes {
    $indexes = @()
    try {
        $wimInfo = dism /Get-WimInfo /WimFile:$Global:InstallWimPath
        foreach ($line in $wimInfo) {
            if ($line -match "Index : (\d+)") {
                $indexes += $matches[1]
            }
        }
    } catch {
        Write-Host "Error getting WIM/ESD indexes." -ForegroundColor Red
    }
    return $indexes
}

function Add-FilesToDesktop {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nDesktop Files Addition Options:" -ForegroundColor Cyan
    Write-Host "1. Add to specific index"
    Write-Host "2. Add to all indexes"
    $choice = Read-Host "Select option (1-2)"
    
    $Global:CurrentOperation = "Adding Desktop Files"
    $Global:CleanupRequired = $true
    
    try {
        if ($choice -eq "1") {
            Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
            try {
                dism /Get-WimInfo /WimFile:$Global:InstallWimPath
                if ($LASTEXITCODE -ne 0) {
                    Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                    return
                }
            } catch {
                Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
                Read-Host "Press Enter to continue"
                return
            }
            
            $indexNumber = Read-Host "`nEnter the index number"
            Add-FilesToDesktopForIndex -IndexNumber $indexNumber
        } elseif ($choice -eq "2") {
            $indexes = Get-WimIndexes
            foreach ($index in $indexes) {
                Add-FilesToDesktopForIndex -IndexNumber $index
            }
        } else {
            Write-Host "Invalid selection." -ForegroundColor Red
        }
    } catch {
        Write-Host "Error during desktop files addition: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
    
    Read-Host "Press Enter to continue"
}

function Add-FilesToDesktopForIndex {
    param([string]$IndexNumber)
    
    # Validate desktop files path BEFORE mounting
    $sourceFiles = Get-DesktopFilesToCopy
    if (-not $sourceFiles -or $sourceFiles.Count -eq 0) {
        Write-Host "No files selected for copying." -ForegroundColor Yellow
        return
    }
    
    try {
        Write-Host "`nMounting WIM/ESD index ${IndexNumber}..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$IndexNumber /MountDir:$Global:MountPath /CheckIntegrity
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${IndexNumber}"
        }
        
        # Find desktop folders and copy files
        $desktopFolders = Find-DesktopFolders -MountPath $Global:MountPath
        
        if ($desktopFolders.Count -eq 0) {
            Write-Host "No desktop folders found in mounted image." -ForegroundColor Yellow
        } else {
            Copy-FilesToDesktopFolders -SourceFiles $sourceFiles -DesktopFolders $desktopFolders -IndexNumber $IndexNumber
        }
        
        Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Unmounting"
        dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "SUCCESS: Desktop files added to index ${IndexNumber}" -ForegroundColor Green
        } else {
            throw "Failed to commit changes for index ${IndexNumber}"
        }
        
    } catch {
        Write-Host "ERROR processing index ${IndexNumber} : $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Attempting to discard changes and unmount..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    }
}

function Get-DesktopFilesToCopy {
    $desktopFilesPath = $Global:DesktopFilesPath
    if (-not $desktopFilesPath) {
        do {
            $path = Read-Host "Enter the path to your desktop files folder"
            $validatedPath = Test-DirectoryPath -Path $path -PathType "Desktop Files"
            if ($validatedPath) {
                $desktopFilesPath = $validatedPath
                $save = Read-Host "Save this desktop files path for future use? (Y/n)"
                if ($save.ToLower() -ne 'n') {
                    $Global:DesktopFilesPath = $desktopFilesPath
                }
                break
            } else {
                Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
            }
        } while ($true)
    } else {
        Write-Host "Using saved desktop files path: ${desktopFilesPath}" -ForegroundColor Green
        $useExisting = Read-Host "Use this path? (Y/n)"
        if ($useExisting.ToLower() -eq 'n') {
            do {
                $path = Read-Host "Enter the path to your desktop files folder"
                $validatedPath = Test-DirectoryPath -Path $path -PathType "Desktop Files"
                if ($validatedPath) {
                    $desktopFilesPath = $validatedPath
                    break
                } else {
                    Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
                }
            } while ($true)
        }
    }
    
    Write-Host "`nScanning for files in: ${desktopFilesPath}" -ForegroundColor Yellow
    $allFiles = Get-ChildItem -Path $desktopFilesPath -File -Recurse -ErrorAction SilentlyContinue
    
    if ($allFiles.Count -eq 0) {
        Write-Host "No files found in: ${desktopFilesPath}" -ForegroundColor Red
        return @()
    }
    
    Write-Host "Found $($allFiles.Count) file(s):" -ForegroundColor Green
    foreach ($file in $allFiles) {
        $relativePath = $file.FullName.Substring($desktopFilesPath.Length + 1)
        Write-Host "  - ${relativePath}" -ForegroundColor Gray
    }
    
    $proceed = Read-Host "`nProceed with copying these files to desktop folders? (Y/n)"
    if ($proceed.ToLower() -eq 'n') {
        return @()
    }
    
    return $allFiles
}

function Find-DesktopFolders {
    param([string]$MountPath)
    
    $desktopFolders = @()
    
    # Only target Default user desktop folder
    $defaultDesktopPath = "$MountPath\Users\Default\Desktop"
    
    if (Test-Path $defaultDesktopPath) {
        $desktopFolders += $defaultDesktopPath
        Write-Host "Found desktop folder: Users\Default\Desktop" -ForegroundColor Green
    } else {
        # Create the desktop folder if it doesn't exist
        try {
            New-Item -ItemType Directory -Path $defaultDesktopPath -Force | Out-Null
            $desktopFolders += $defaultDesktopPath
            Write-Host "Created desktop folder: Users\Default\Desktop" -ForegroundColor Cyan
        } catch {
            Write-Host "Warning: Could not create desktop folder: Users\Default\Desktop" -ForegroundColor Yellow
        }
    }
    
    return $desktopFolders
}

function Copy-FilesToDesktopFolders {
    param(
        [array]$SourceFiles,
        [array]$DesktopFolders,
        [string]$IndexNumber
    )
    
    $sourceBasePath = if ($Global:DesktopFilesPath) { $Global:DesktopFilesPath } else { (Get-Location).Path }
    
    Write-Host "`nCopying files to desktop folders for index ${IndexNumber}..." -ForegroundColor Yellow
    
    foreach ($desktopFolder in $DesktopFolders) {
        $folderName = ($desktopFolder -split '\\')[-2..-1] -join '\'
        Write-Host "`nTarget: ${folderName}" -ForegroundColor Cyan
        
        $successCount = 0
        $failCount = 0
        
        foreach ($file in $SourceFiles) {
            try {
                # Preserve directory structure if files are in subdirectories
                $relativePath = $file.FullName.Substring($sourceBasePath.Length + 1)
                $targetFile = Join-Path $desktopFolder $relativePath
                $targetDir = Split-Path $targetFile -Parent
                
                # Create target directory if it doesn't exist
                if (-not (Test-Path $targetDir)) {
                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                }
                
                Copy-Item -Path $file.FullName -Destination $targetFile -Force
                Write-Host "  ✓ ${relativePath}" -ForegroundColor Green
                $successCount++
                
            } catch {
                Write-Host "  ✗ $($file.Name) - Error: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        Write-Host "  Summary: ${successCount} copied, $failCount failed" -ForegroundColor $(if ($failCount -eq 0) { "Green" } else { "Yellow" })
    }
}

function Add-DesktopFilesToPackageManagement {
    param([string]$ImagePath, [bool]$IsMultiIndex, [bool]$IsFirstIndex = $true)
    
    if ($IsMultiIndex -and -not $IsFirstIndex) {
        # For subsequent indexes in multi-index mode, we don't need to ask again
        # The files are already determined from the first index
        return
    }
    
    Write-Host "`nAdding Desktop Files" -ForegroundColor Cyan
    
    # Get source files using the same logic as the standalone function
    $sourceFiles = Get-DesktopFilesToCopy
    if (-not $sourceFiles -or $sourceFiles.Count -eq 0) {
        Write-Host "No files selected for copying to desktop." -ForegroundColor Yellow
        return
    }
    
    # Find desktop folders and copy files
    $desktopFolders = Find-DesktopFolders -MountPath $ImagePath
    
    if ($desktopFolders.Count -eq 0) {
        Write-Host "No desktop folders found in mounted image." -ForegroundColor Yellow
    } else {
        Copy-FilesToDesktopFolders -SourceFiles $sourceFiles -DesktopFolders $desktopFolders -IndexNumber "Current"
        Write-Host "✓ Desktop files added successfully" -ForegroundColor Green
    }
}

function Add-DriversToIndex {
    param(
        [string]$IndexNumber,
        [string]$DriverPath
    )
    
    try {
        Write-Host "`nMounting WIM/ESD index ${IndexNumber}..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$IndexNumber /MountDir:$Global:MountPath /CheckIntegrity
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${IndexNumber}"
        }
        
        Write-Host "Adding drivers from $DriverPath..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Adding Drivers"
        
        if ($Global:ForceUnsignedDrivers) {
            dism /Image:$Global:MountPath /Add-Driver /Driver:$DriverPath /Recurse /ForceUnsigned
        } else {
            dism /Image:$Global:MountPath /Add-Driver /Driver:$DriverPath /Recurse
        }
        
        Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Unmounting"
        dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "SUCCESS: Drivers added to index ${IndexNumber}" -ForegroundColor Green
        } else {
            throw "Failed to commit changes for index ${IndexNumber}"
        }
        
    } catch {
        Write-Host "ERROR processing index ${IndexNumber} : $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Attempting to discard changes and unmount..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    }
}

function Manage-PackagesAndFeatures {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nPackage and Feature Management Options:" -ForegroundColor Cyan
    Write-Host "1. Manage for specific index"
    Write-Host "2. Manage for all indexes"
    Write-Host "3. Manage Current Live System" -ForegroundColor Green
    Write-Host "4. Return to Main Menu" -ForegroundColor Red
    $choice = Read-Host "Select option (1-4)"
    
    $Global:EnabledFeatures = @()
    $Global:DisabledFeatures = @()
    $Global:AddedPackages = @()
    $Global:RemovedPackages = @()
    $Global:AddedAppxPackages = @()
    $Global:RemovedAppxPackages = @()
    $Global:SelectedDesktopFiles = @()
    $Global:ShouldAddDesktopFiles = $false
    
    switch ($choice) {
        "1" {
            $addDrivers = Read-Host "`nDo you want to add drivers to the image(s)? (y/N)"
            $shouldAddDrivers = $addDrivers.ToLower() -eq 'y'
            $driverPath = $null
            
            if ($shouldAddDrivers) {
                $driverPath = $Global:DriverPath
                if (-not $driverPath) {
                    do {
                        $path = Read-Host "Enter the path to your drivers folder"
                        $validatedPath = Test-DirectoryPath -Path $path -PathType "Driver"
                        if ($validatedPath) {
                            $driverPath = $validatedPath
                            $save = Read-Host "Save this driver path for future use? (Y/n)"
                            if ($save.ToLower() -ne 'n') {
                                $Global:DriverPath = $driverPath
                            }
                            break
                        } else {
                            Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
                        }
                    } while ($true)
                } else {
                    Write-Host "Using saved driver path: $driverPath" -ForegroundColor Green
                    $useExisting = Read-Host "Use this path? (Y/n)"
                    if ($useExisting.ToLower() -eq 'n') {
                        do {
                            $path = Read-Host "Enter the path to your drivers folder"
                            $validatedPath = Test-DirectoryPath -Path $path -PathType "Driver"
                            if ($validatedPath) {
                                $driverPath = $validatedPath
                                break
                            } else {
                                Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
                            }
                        } while ($true)
                    }
                }
            }
            
            $addDesktopFiles = Read-Host "`nDo you want to add desktop files to the image(s)? (y/N)"
            $shouldAddDesktopFiles = $addDesktopFiles.ToLower() -eq 'y'
            
            # Store globally for use across indexes
            $Global:ShouldAddDesktopFiles = $shouldAddDesktopFiles
            
            if ($shouldAddDesktopFiles) {
                # Get the files once and store them globally
                Write-Host "`nPreparing desktop files for installation..." -ForegroundColor Yellow
                $Global:SelectedDesktopFiles = Get-DesktopFilesToCopy
                if (-not $Global:SelectedDesktopFiles -or $Global:SelectedDesktopFiles.Count -eq 0) {
                    Write-Host "No desktop files selected. Skipping desktop file installation." -ForegroundColor Yellow
                    $Global:ShouldAddDesktopFiles = $false
                }
            }
            
            $Global:CurrentOperation = "Managing Packages and Features"
            $Global:CleanupRequired = $true
            
            try {
                Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
                try {
                    dism /Get-WimInfo /WimFile:$Global:InstallWimPath
                    if ($LASTEXITCODE -ne 0) {
                        Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
                        Read-Host "Press Enter to continue"
                        return
                    }
                } catch {
                    Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                    return
                }
                
                $indexNumber = Read-Host "`nEnter the index number"
                Manage-PackagesAndFeaturesForIndex -IndexNumber $indexNumber -AddDrivers $shouldAddDrivers -DriverPath $driverPath -IsMultiIndex $false
            } catch {
                Write-Host "Error during package and feature management: $($_.Exception.Message)" -ForegroundColor Red
            } finally {
                $Global:CurrentOperation = "None"
                $Global:CleanupRequired = $false
            }
        }
        "2" {
            $addDrivers = Read-Host "`nDo you want to add drivers to the image(s)? (y/N)"
            $shouldAddDrivers = $addDrivers.ToLower() -eq 'y'
            $driverPath = $null
            
            if ($shouldAddDrivers) {
                $driverPath = $Global:DriverPath
                if (-not $driverPath) {
                    do {
                        $path = Read-Host "Enter the path to your drivers folder"
                        $validatedPath = Test-DirectoryPath -Path $path -PathType "Driver"
                        if ($validatedPath) {
                            $driverPath = $validatedPath
                            $save = Read-Host "Save this driver path for future use? (Y/n)"
                            if ($save.ToLower() -ne 'n') {
                                $Global:DriverPath = $driverPath
                            }
                            break
                        } else {
                            Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
                        }
                    } while ($true)
                } else {
                    Write-Host "Using saved driver path: $driverPath" -ForegroundColor Green
                    $useExisting = Read-Host "Use this path? (Y/n)"
                    if ($useExisting.ToLower() -eq 'n') {
                        do {
                            $path = Read-Host "Enter the path to your drivers folder"
                            $validatedPath = Test-DirectoryPath -Path $path -PathType "Driver"
                            if ($validatedPath) {
                                $driverPath = $validatedPath
                                break
                            } else {
                                Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
                            }
                        } while ($true)
                    }
                }
            }
            
            $addDesktopFiles = Read-Host "`nDo you want to add desktop files to the image(s)? (y/N)"
            $shouldAddDesktopFiles = $addDesktopFiles.ToLower() -eq 'y'
            
            # Store globally for use across indexes
            $Global:ShouldAddDesktopFiles = $shouldAddDesktopFiles
            
            if ($shouldAddDesktopFiles) {
                # Get the files once and store them globally
                Write-Host "`nPreparing desktop files for installation..." -ForegroundColor Yellow
                $Global:SelectedDesktopFiles = Get-DesktopFilesToCopy
                if (-not $Global:SelectedDesktopFiles -or $Global:SelectedDesktopFiles.Count -eq 0) {
                    Write-Host "No desktop files selected. Skipping desktop file installation." -ForegroundColor Yellow
                    $Global:ShouldAddDesktopFiles = $false
                }
            }
            
            $Global:CurrentOperation = "Managing Packages and Features"
            $Global:CleanupRequired = $true
            
            try {
                $indexes = Get-WimIndexes
                
                Write-Host "`nSelect which Windows edition to use as base for package/feature selection:" -ForegroundColor Yellow
                Write-Host "(Different editions have different available packages and features)" -ForegroundColor Gray
                try {
                    dism /Get-WimInfo /WimFile:$Global:InstallWimPath
                } catch {
                    Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                    return
                }
                
                do {
                    $baseIndex = Read-Host "`nEnter the index number of the edition to use as base"
                    if ($indexes -contains $baseIndex) {
                        break
                    } else {
                        Write-Host "Invalid index. Please select from the available indexes above." -ForegroundColor Red
                    }
                } while ($true)
                
                Write-Host "`n--- Processing Base Index $baseIndex (Package/Feature Selection) ---" -ForegroundColor Magenta
                Manage-PackagesAndFeaturesForIndex -IndexNumber $baseIndex -AddDrivers $shouldAddDrivers -DriverPath $driverPath -IsMultiIndex $true -IsFirstIndex $true
                
                foreach ($index in $indexes) {
                    if ($index -ne $baseIndex) {
                        Write-Host "`n--- Processing Index ${index} ---" -ForegroundColor Magenta
                        Manage-PackagesAndFeaturesForIndex -IndexNumber $index -AddDrivers $shouldAddDrivers -DriverPath $driverPath -IsMultiIndex $true -IsFirstIndex $false
                    }
                }
            } catch {
                Write-Host "Error during package and feature management: $($_.Exception.Message)" -ForegroundColor Red
            } finally {
                $Global:CurrentOperation = "None"
                $Global:CleanupRequired = $false
            }
        }
        "3" {
            Write-Host "`nManaging Current Live System" -ForegroundColor Green
            Write-Host "============================" -ForegroundColor Green
            Write-Host "This will work directly on your currently running Windows installation." -ForegroundColor Yellow
            Write-Host "Changes will take effect immediately and cannot be easily undone!" -ForegroundColor Red
            Write-Host ""
            
            $confirm = Read-Host "Are you sure you want to proceed with live system management? (y/N)"
            if ($confirm.ToLower() -eq 'y') {
                Manage-LiveSystem
            } else {
                Write-Host "Live system management cancelled." -ForegroundColor Yellow
            }
        }
        "4" { return }
        default {
            Write-Host "Invalid selection." -ForegroundColor Red
            Read-Host "Press Enter to continue"
        }
    }
}

function Manage-PackagesAndFeaturesForIndex {
    param(
        [string]$IndexNumber,
        [bool]$AddDrivers,
        [string]$DriverPath,
        [bool]$IsMultiIndex,
        [bool]$IsFirstIndex = $true
    )
    
    try {
        Write-Host "`nMounting WIM/ESD index ${IndexNumber}..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$IndexNumber /MountDir:$Global:MountPath /CheckIntegrity
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${IndexNumber}"
        }
        
        if ($AddDrivers -and $DriverPath) {
            Write-Host "`nAdding drivers from $DriverPath..." -ForegroundColor Yellow
            $Global:CurrentOperation = "Adding Drivers"
            
            if ($Global:ForceUnsignedDrivers) {
                dism /Image:$Global:MountPath /Add-Driver /Driver:$DriverPath /Recurse /ForceUnsigned
            } else {
                dism /Image:$Global:MountPath /Add-Driver /Driver:$DriverPath /Recurse
            }
            
            if ($LASTEXITCODE -ne 0) {
                Write-Host "Warning: Driver installation may have had issues for index ${IndexNumber}" -ForegroundColor Yellow
            } else {
                Write-Host "✓ Drivers added successfully" -ForegroundColor Green
            }
        }
        
        if ($Global:ShouldAddDesktopFiles -and $Global:SelectedDesktopFiles.Count -gt 0) {
            Write-Host "`nAdding desktop files..." -ForegroundColor Yellow
            
            # Find desktop folders and copy files
            $desktopFolders = Find-DesktopFolders -MountPath $Global:MountPath
            
            if ($desktopFolders.Count -eq 0) {
                Write-Host "No desktop folders found in mounted image." -ForegroundColor Yellow
            } else {
                Copy-FilesToDesktopFolders -SourceFiles $Global:SelectedDesktopFiles -DesktopFolders $desktopFolders -IndexNumber $IndexNumber
                Write-Host "✓ Desktop files added successfully" -ForegroundColor Green
            }
        }
        
        if ($IsMultiIndex -and -not $IsFirstIndex) {
            if ($Global:AddedPackages.Count -gt 0) {
                Write-Host "`nApplying previously selected system packages to add..." -ForegroundColor Yellow
                foreach ($package in $Global:AddedPackages) {
                    Write-Host "Adding: $(Split-Path $package -Leaf)" -ForegroundColor Gray
                    try {
                        dism /Image:$Global:MountPath /Add-Package /PackagePath:$package
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host "✓ Successfully added: $(Split-Path $package -Leaf)" -ForegroundColor Green
                        } else {
                            Write-Host "✗ Failed to add: $(Split-Path $package -Leaf)" -ForegroundColor Red
                        }
                    } catch {
                        Write-Host "✗ Error adding $(Split-Path $package -Leaf) : $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
            
            if ($Global:RemovedPackages.Count -gt 0) {
                Write-Host "`nApplying previously selected system packages to remove..." -ForegroundColor Yellow
                foreach ($package in $Global:RemovedPackages) {
                    Write-Host "Removing: $package" -ForegroundColor Gray
                    try {
                        dism /Image:$Global:MountPath /Remove-Package /PackageName:$package
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host "✓ Successfully removed: $package" -ForegroundColor Green
                        } else {
                            Write-Host "✗ Failed to remove: $package" -ForegroundColor Red
                        }
                    } catch {
                        Write-Host "✗ Error removing $package : $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
            
            if ($Global:AddedAppxPackages.Count -gt 0) {
                Write-Host "`nApplying previously selected application packages to add..." -ForegroundColor Yellow
                foreach ($package in $Global:AddedAppxPackages) {
                    $fileType = if ($package -like "*.appxbundle") { "APPXBUNDLE" } else { "APPX" }
                    Write-Host "Adding ${fileType}: $(Split-Path $package -Leaf)" -ForegroundColor Gray
                    try {
                        if ($package -like "*.appxbundle") {
                            dism /Image:$Global:MountPath /Add-ProvisionedAppxPackage /PackagePath:$package /SkipLicense
                        } else {
                            # For APPX files, check for dependencies and license in multi-index mode
                            $packageDir = Split-Path $package -Parent
                            $dependencies = Get-ChildItem -Path $packageDir -Filter "*dependency*.appx" -ErrorAction SilentlyContinue
                            $frameworkDeps = Get-ChildItem -Path $packageDir -Filter "*framework*.appx" -ErrorAction SilentlyContinue
                            $allDeps = @()
                            if ($dependencies) { $allDeps += $dependencies.FullName }
                            if ($frameworkDeps) { $allDeps += $frameworkDeps.FullName }
                            
                            $licenseFile = Get-ChildItem -Path $packageDir -Filter "*license*.xml" -ErrorAction SilentlyContinue | Select-Object -First 1
                            
                            if ($allDeps.Count -gt 0 -or $licenseFile) {
                                $dismArgs = @("/Image:$Global:MountPath", "/Add-ProvisionedAppxPackage", "/FolderPath:$packageDir")
                                
                                foreach ($dep in $allDeps) {
                                    $dismArgs += "/DependencyPackagePath:$dep"
                                }
                                
                                if ($licenseFile) {
                                    $dismArgs += "/LicensePath:$($licenseFile.FullName)"
                                }
                                
                                dism @dismArgs
                            } else {
                                dism /Image:$Global:MountPath /Add-ProvisionedAppxPackage /PackagePath:$package /SkipLicense
                            }
                        }
                        
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host "✓ Successfully added: $(Split-Path $package -Leaf)" -ForegroundColor Green
                        } else {
                            Write-Host "✗ Failed to add: $(Split-Path $package -Leaf)" -ForegroundColor Red
                        }
                    } catch {
                        Write-Host "✗ Error adding $(Split-Path $package -Leaf) : $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
            
            if ($Global:RemovedAppxPackages.Count -gt 0) {
                Write-Host "`nApplying previously selected application packages to remove..." -ForegroundColor Yellow
                foreach ($package in $Global:RemovedAppxPackages) {
                    Write-Host "Removing application package: $package" -ForegroundColor Gray
                    try {
                        dism /Image:$Global:MountPath /Remove-ProvisionedAppxPackage /PackageName:$package
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host "✓ Successfully removed: $package" -ForegroundColor Green
                        } else {
                            Write-Host "✗ Failed to remove: $package" -ForegroundColor Red
                        }
                    } catch {
                        Write-Host "✗ Error removing $package : $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
            
            if ($Global:EnabledFeatures.Count -gt 0) {
                Write-Host "`nApplying previously selected features to enable..." -ForegroundColor Yellow
                foreach ($feature in $Global:EnabledFeatures) {
                    Write-Host "Enabling: $feature" -ForegroundColor Gray
                    try {
                        dism /Image:$Global:MountPath /Enable-Feature /FeatureName:$feature /All
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host "✓ Successfully enabled: $feature" -ForegroundColor Green
                        } else {
                            Write-Host "✗ Failed to enable: $feature" -ForegroundColor Red
                        }
                    } catch {
                        Write-Host "✗ Error enabling $feature : $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
            
            if ($Global:DisabledFeatures.Count -gt 0) {
                Write-Host "`nApplying previously selected features to disable..." -ForegroundColor Yellow
                foreach ($featureInfo in $Global:DisabledFeatures) {
                    $feature = $featureInfo.Name
                    $removePayload = $featureInfo.RemovePayload
                    
                    Write-Host "Disabling: $feature" -ForegroundColor Gray
                    try {
                        if ($removePayload) {
                            dism /Image:$Global:MountPath /Disable-Feature /FeatureName:$feature /Remove
                        } else {
                            dism /Image:$Global:MountPath /Disable-Feature /FeatureName:$feature
                        }
                        
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host "✓ Successfully disabled: $feature" -ForegroundColor Green
                        } else {
                            Write-Host "✗ Failed to disable: $feature" -ForegroundColor Red
                        }
                    } catch {
                        Write-Host "✗ Error disabling $feature : $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
        } else {
            do {
                Clear-Host
                Write-Host "Package and Feature Management - Index: ${IndexNumber}" -ForegroundColor Cyan
                Write-Host "=================================================" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "SYSTEM PACKAGE MANAGEMENT (.cab/.msu):" -ForegroundColor Magenta
                Write-Host "1. View all installed system packages"
                Write-Host "2. View specific system package info"
                Write-Host "3. Add system package(s)"
                Write-Host "4. Remove system package(s)"
                Write-Host ""
                Write-Host "APPLICATION PACKAGE MANAGEMENT (.appx/.appxbundle):" -ForegroundColor Magenta
                Write-Host "5. View all provisioned application packages"
                Write-Host "6. Add application package(s)"
                Write-Host "7. Remove application package(s)"
                Write-Host ""
                Write-Host "FEATURE MANAGEMENT:" -ForegroundColor Magenta
                Write-Host "8. View all available features"
                Write-Host "9. View specific feature info"
                Write-Host "10. Enable feature(s)"
                Write-Host "11. Disable feature(s)"
                Write-Host ""
                Write-Host "ACTIONS:" -ForegroundColor Magenta
                Write-Host "12. Commit changes and continue"
                Write-Host "13. Discard changes and exit"
                
                $managementChoice = Read-Host "`nSelect option (1-13)"
                
                switch ($managementChoice) {
                    "1" { 
                        Show-InstalledPackages -ImagePath $Global:MountPath
                        Read-Host "`nPress Enter to continue"
                    }
                    "2" {
                        $packageName = Read-Host "Enter system package name"
                        Show-PackageInfo -ImagePath $Global:MountPath -PackageName $packageName
                        Read-Host "`nPress Enter to continue"
                    }
                    "3" {
                        Add-Packages -ImagePath $Global:MountPath -IsMultiIndex $IsMultiIndex
                    }
                    "4" {
                        Remove-Packages -ImagePath $Global:MountPath -IsMultiIndex $IsMultiIndex
                    }
                    "5" { 
                        Show-ProvisionedAppxPackages -ImagePath $Global:MountPath
                        Read-Host "`nPress Enter to continue"
                    }
                    "6" {
                        Add-AppxPackages -ImagePath $Global:MountPath -IsMultiIndex $IsMultiIndex
                    }
                    "7" {
                        Remove-AppxPackages -ImagePath $Global:MountPath -IsMultiIndex $IsMultiIndex
                    }
                    "8" { 
                        Show-AvailableFeatures -ImagePath $Global:MountPath
                        Read-Host "`nPress Enter to continue"
                    }
                    "9" {
                        $featureName = Read-Host "Enter feature name"
                        Show-FeatureInfo -ImagePath $Global:MountPath -FeatureName $featureName
                        Read-Host "`nPress Enter to continue"
                    }
                    "10" {
                        Enable-Features -ImagePath $Global:MountPath -IsMultiIndex $IsMultiIndex
                    }
                    "11" {
                        Disable-Features -ImagePath $Global:MountPath -IsMultiIndex $IsMultiIndex
                    }
                    "12" {
                        Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
                        $Global:CurrentOperation = "Unmounting"
                        dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
                        
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host "SUCCESS: Changes committed for index ${IndexNumber}" -ForegroundColor Green
                        } else {
                            throw "Failed to commit changes for index ${IndexNumber}"
                        }
                        return
                    }
                    "13" {
                        Write-Host "Discarding changes and unmounting..." -ForegroundColor Yellow
                        dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                        Write-Host "Changes discarded." -ForegroundColor Yellow
                        return
                    }
                    default {
                        Write-Host "Invalid selection." -ForegroundColor Red
                        Start-Sleep -Seconds 1
                    }
                }
            } while ($true)
        }
        
        if ($IsMultiIndex -and -not $IsFirstIndex) {
            Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
            $Global:CurrentOperation = "Unmounting"
            dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "SUCCESS: Changes committed for index ${IndexNumber}" -ForegroundColor Green
            } else {
                throw "Failed to commit changes for index ${IndexNumber}"
            }
        }
        
    } catch {
        Write-Host "ERROR processing index ${IndexNumber} : $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Attempting to discard changes and unmount..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    }
}

function Get-KeyNameFromTargetEdition {
    param([string]$TargetEdition)
    
    # Map DISM target editions to exact key names we actually have
    switch ($TargetEdition) {
        # Core/Home editions
        "Core" { return "Windows 11 Home" }
        "CoreN" { return "Windows 11 Home N" }
        "CoreSingleLanguage" { return "Windows 11 Home Single Language" }
        "CoreCountrySpecific" { return "Windows 11 Home Country Specific" }
        
        # Professional editions  
        "Professional" { return "Windows 11 Pro" }
        "ProfessionalN" { return "Windows 11 Pro N" }
        "ProfessionalWorkstation" { return "Windows 11 Pro for Workstations" }
        "ProfessionalWorkstationN" { return "Windows 11 Pro N for Workstations" }
        "ProfessionalEducation" { return "Windows 11 Pro Education" }
        "ProfessionalEducationN" { return "Windows 11 Pro Education N" }
        
        # Education editions
        "Education" { return "Windows 11 Education" }
        "EducationN" { return "Windows 11 Education N" }
        
        # Enterprise editions
        "Enterprise" { return "Windows 11 Enterprise" }
        "EnterpriseN" { return "Windows 11 Enterprise N" }
        "EnterpriseG" { return "Windows 11 Enterprise G" }
        "EnterpriseGN" { return "Windows 11 Enterprise G N" }
        "EnterpriseS" { return "Windows 10 Enterprise S" }
        
        # IoT Enterprise
        "IoTEnterprise" { return "Windows IoT Enterprise LTSC 2024" }
        
        default { return $null }
    }
}

function Manage-LiveSystem {
    Write-Host "`nLive System Package and Feature Management" -ForegroundColor Green
    Write-Host "WARNING: Changes are applied immediately to your running system!" -ForegroundColor Red
    Write-Host ""
    
    do {
        Clear-Host
        Write-Host "Live System Management" -ForegroundColor Green
        Write-Host "======================" -ForegroundColor Green
        Write-Host ""
        Write-Host "SYSTEM PACKAGE MANAGEMENT (.cab/.msu):" -ForegroundColor Magenta
        Write-Host "1. View all installed system packages"
        Write-Host "2. View specific system package info"
        Write-Host "3. Add system package(s)"
        Write-Host "4. Remove system package(s)"
        Write-Host ""
        Write-Host "APPLICATION PACKAGE MANAGEMENT (.appx/.appxbundle):" -ForegroundColor Magenta
        Write-Host "5. View all provisioned application packages"
        Write-Host "6. Add application package(s)"
        Write-Host "7. Remove application package(s)"
        Write-Host ""
        Write-Host "FEATURE MANAGEMENT:" -ForegroundColor Magenta
        Write-Host "8. View all available features"
        Write-Host "9. View specific feature info"
        Write-Host "10. Enable feature(s)"
        Write-Host "11. Disable feature(s)"
        Write-Host ""
        Write-Host "ACTIONS:" -ForegroundColor Magenta
        Write-Host "12. Return to Package Management Menu"
        
        $managementChoice = Read-Host "`nSelect option (1-12)"
        
        switch ($managementChoice) {
            "1" { 
                Show-LiveInstalledPackages
                Read-Host "`nPress Enter to continue"
            }
            "2" {
                $packageName = Read-Host "Enter system package name"
                Show-LivePackageInfo -PackageName $packageName
                Read-Host "`nPress Enter to continue"
            }
            "3" {
                Add-LivePackages
            }
            "4" {
                Remove-LivePackages
            }
            "5" { 
                Show-LiveProvisionedAppxPackages
                Read-Host "`nPress Enter to continue"
            }
            "6" {
                Add-LiveAppxPackages
            }
            "7" {
                Remove-LiveAppxPackages
            }
            "8" { 
                Show-LiveAvailableFeatures
                Read-Host "`nPress Enter to continue"
            }
            "9" {
                $featureName = Read-Host "Enter feature name"
                Show-LiveFeatureInfo -FeatureName $featureName
                Read-Host "`nPress Enter to continue"
            }
            "10" {
                Enable-LiveFeatures
            }
            "11" {
                Disable-LiveFeatures
            }
            "12" {
                return
            }
            default {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($true)
}

function Show-LiveInstalledPackages {
    Write-Host "Getting installed system packages from live system..." -ForegroundColor Yellow
    try {
        Write-Host "`nInstalled System Packages (.cab/.msu) - Live System:" -ForegroundColor Cyan
        Write-Host "===================================================" -ForegroundColor Cyan
        dism /Online /Get-Packages
    } catch {
        Write-Host "Error getting live system packages: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-LivePackageInfo {
    param([string]$PackageName)
    
    try {
        Write-Host "`nDetailed information for package: $PackageName (Live System)" -ForegroundColor Cyan
        dism /Online /Get-PackageInfo /PackageName:$PackageName
    } catch {
        Write-Host "Error getting live package info: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Add-LivePackages {
    Write-Host "`nAdding System Packages to Live System" -ForegroundColor Cyan
    Write-Host "WARNING: Changes will be applied immediately!" -ForegroundColor Red
    
    $packagePath = $Global:PackagePath
    if (-not $packagePath) {
        do {
            $path = Read-Host "Enter the path to your system packages folder"
            $validatedPath = Test-DirectoryPath -Path $path -PathType "Package"
            if ($validatedPath) {
                $packagePath = $validatedPath
                break
            } else {
                Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
            }
        } while ($true)
    } else {
        Write-Host "Using saved package path: $packagePath" -ForegroundColor Green
        $useExisting = Read-Host "Use this path? (Y/n)"
        if ($useExisting.ToLower() -eq 'n') {
            do {
                $path = Read-Host "Enter the path to your system packages folder"
                $validatedPath = Test-DirectoryPath -Path $path -PathType "Package"
                if ($validatedPath) {
                    $packagePath = $validatedPath
                    break
                } else {
                    Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
                }
            } while ($true)
        }
    }
    
    $cabFiles = Get-ChildItem -Path $packagePath -Filter "*.cab" -Recurse -ErrorAction SilentlyContinue
    $msuFiles = Get-ChildItem -Path $packagePath -Filter "*.msu" -Recurse -ErrorAction SilentlyContinue
    
    $packages = @()
    if ($cabFiles) { $packages += $cabFiles.FullName }
    if ($msuFiles) { $packages += $msuFiles.FullName }
    
    if ($packages.Count -eq 0) {
        Write-Host "No .cab or .msu files found in: $packagePath" -ForegroundColor Red
        Read-Host "`nPress Enter to continue"
        return
    }
    
    Write-Host "Found $($packages.Count) system package file(s):" -ForegroundColor Green
    foreach ($pkg in $packages) {
        Write-Host "  - $(Split-Path $pkg -Leaf)" -ForegroundColor Gray
    }
    
    $proceed = Read-Host "`nProceed with adding these packages to the live system? (y/N)"
    if ($proceed.ToLower() -eq 'y') {
        Write-Host "`nAdding packages to live system..." -ForegroundColor Yellow
        foreach ($package in $packages) {
            Write-Host "Adding: $(Split-Path $package -Leaf)" -ForegroundColor Gray
            try {
                dism /Online /Add-Package /PackagePath:$package
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "✓ Successfully added: $(Split-Path $package -Leaf)" -ForegroundColor Green
                } else {
                    Write-Host "✗ Failed to add: $(Split-Path $package -Leaf)" -ForegroundColor Red
                }
            } catch {
                Write-Host "✗ Error adding $(Split-Path $package -Leaf) : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    Read-Host "`nPress Enter to continue"
}

function Remove-LivePackages {
    Write-Host "`nRemoving System Packages from Live System" -ForegroundColor Cyan
    Write-Host "WARNING: Changes will be applied immediately!" -ForegroundColor Red
    
    Write-Host "Getting installed packages for reference..." -ForegroundColor Yellow
    try {
        dism /Online /Get-Packages
    } catch {
        Write-Host "Error getting packages: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nEnter package names separated by commas:"
    $packageInput = Read-Host "Package names to remove"
    if (-not $packageInput) {
        return
    }
    
    $packages = $packageInput -split ',' | ForEach-Object { $_.Trim() }
    
    $confirm = Read-Host "`nAre you sure you want to remove these packages from the live system? (y/N)"
    if ($confirm.ToLower() -eq 'y') {
        Write-Host "`nRemoving packages from live system..." -ForegroundColor Yellow
        foreach ($package in $packages) {
            Write-Host "Removing: $package" -ForegroundColor Gray
            try {
                dism /Online /Remove-Package /PackageName:$package
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "✓ Successfully removed: $package" -ForegroundColor Green
                } else {
                    Write-Host "✗ Failed to remove: $package" -ForegroundColor Red
                }
            } catch {
                Write-Host "✗ Error removing $package : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    Read-Host "`nPress Enter to continue"
}

function Show-LiveProvisionedAppxPackages {
    Write-Host "Getting provisioned application packages from live system..." -ForegroundColor Yellow
    try {
        Write-Host "`nProvisioned Application Packages - Live System:" -ForegroundColor Cyan
        Write-Host "===============================================" -ForegroundColor Cyan
        dism /Online /Get-ProvisionedAppxPackages
    } catch {
        Write-Host "Error getting live application packages: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Add-LiveAppxPackages {
    Write-Host "`nAdding Application Packages to Live System" -ForegroundColor Cyan
    Write-Host "WARNING: Changes will be applied immediately to your running system!" -ForegroundColor Red
    
    $appxPackagePath = $Global:AppxPackagePath
    if (-not $appxPackagePath) {
        do {
            $path = Read-Host "Enter the path to your APPX/APPXBUNDLE packages folder"
            $validatedPath = Test-DirectoryPath -Path $path -PathType "APPX Package"
            if ($validatedPath) {
                $appxPackagePath = $validatedPath
                $save = Read-Host "Save this APPX package path for future use? (Y/n)"
                if ($save.ToLower() -ne 'n') {
                    $Global:AppxPackagePath = $appxPackagePath
                }
                break
            } else {
                Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
            }
        } while ($true)
    } else {
        Write-Host "Using saved APPX package path: ${appxPackagePath}" -ForegroundColor Green
        $useExisting = Read-Host "Use this path? (Y/n)"
        if ($useExisting.ToLower() -eq 'n') {
            do {
                $path = Read-Host "Enter the path to your APPX/APPXBUNDLE packages folder"
                $validatedPath = Test-DirectoryPath -Path $path -PathType "APPX Package"
                if ($validatedPath) {
                    $appxPackagePath = $validatedPath
                    break
                } else {
                    Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
                }
            } while ($true)
        }
    }
    
    Write-Host "`nScanning for APPX/APPXBUNDLE packages in: ${appxPackagePath}" -ForegroundColor Yellow
    $appxFiles = Get-ChildItem -Path $appxPackagePath -Filter "*.appx" -Recurse -ErrorAction SilentlyContinue
    $appxbundleFiles = Get-ChildItem -Path $appxPackagePath -Filter "*.appxbundle" -Recurse -ErrorAction SilentlyContinue
    
    $packages = @()
    if ($appxFiles) {
        $packages += $appxFiles.FullName
    }
    if ($appxbundleFiles) {
        $packages += $appxbundleFiles.FullName
    }
    
    $packages = $packages | Where-Object { $_ -and $_.Trim() }
    
    if ($packages.Count -eq 0) {
        Write-Host "No .appx or .appxbundle files found in: ${appxPackagePath}" -ForegroundColor Red
        Read-Host "`nPress Enter to continue"
        return
    }
    
    Write-Host "Found $($packages.Count) application package file(s):" -ForegroundColor Green
    foreach ($pkg in $packages) {
        $fileType = if ($pkg -like "*.appxbundle") { "APPXBUNDLE" } else { "APPX" }
        Write-Host "  - $(Split-Path $pkg -Leaf) [$fileType]" -ForegroundColor Gray
    }
    
    # Group packages by type for processing
    $appxPackages = $packages | Where-Object { $_ -like "*.appx" }
    $appxbundlePackages = $packages | Where-Object { $_ -like "*.appxbundle" }
    
    $proceed = Read-Host "`nProceed with adding these application packages to the live system? (y/N)"
    if ($proceed.ToLower() -ne 'y') {
        return
    }
    
    Write-Host "`nAdding application packages to live system..." -ForegroundColor Yellow
    
    # Process APPX files (may need dependencies)
    foreach ($package in $appxPackages) {
        Write-Host "`nProcessing APPX: $(Split-Path $package -Leaf)" -ForegroundColor Cyan
        
        # Check for dependencies and license
        $packageDir = Split-Path $package -Parent
        $packageName = [System.IO.Path]::GetFileNameWithoutExtension($package)
        
        # Look for dependencies
        $dependencies = Get-ChildItem -Path $packageDir -Filter "*dependency*.appx" -ErrorAction SilentlyContinue
        $frameworkDeps = Get-ChildItem -Path $packageDir -Filter "*framework*.appx" -ErrorAction SilentlyContinue
        $allDeps = @()
        if ($dependencies) { $allDeps += $dependencies.FullName }
        if ($frameworkDeps) { $allDeps += $frameworkDeps.FullName }
        
        # Look for license file
        $licenseFile = Get-ChildItem -Path $packageDir -Filter "*license*.xml" -ErrorAction SilentlyContinue | Select-Object -First 1
        
        try {
            if ($allDeps.Count -gt 0 -or $licenseFile) {
                Write-Host "  Using folder-based installation..." -ForegroundColor Gray
                
                # Use folder path method for complex installations
                $dismArgs = @("/Online", "/Add-ProvisionedAppxPackage", "/FolderPath:$packageDir")
                
                if ($allDeps.Count -gt 0) {
                    Write-Host "  Found dependencies: $($allDeps.Count)" -ForegroundColor Yellow
                    foreach ($dep in $allDeps) {
                        $dismArgs += "/DependencyPackagePath:$dep"
                        Write-Host "    - $(Split-Path $dep -Leaf)" -ForegroundColor Gray
                    }
                }
                
                if ($licenseFile) {
                    Write-Host "  Found license: $(Split-Path $licenseFile.FullName -Leaf)" -ForegroundColor Yellow
                    $dismArgs += "/LicensePath:$($licenseFile.FullName)"
                }
                
                dism @dismArgs
            } else {
                Write-Host "  Using direct package installation..." -ForegroundColor Gray
                dism /Online /Add-ProvisionedAppxPackage /PackagePath:$package /SkipLicense
            }
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  ✓ Successfully added: $(Split-Path $package -Leaf)" -ForegroundColor Green
            } else {
                Write-Host "  ✗ Failed to add: $(Split-Path $package -Leaf)" -ForegroundColor Red
            }
        } catch {
            Write-Host "  ✗ Error adding $(Split-Path $package -Leaf) : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # Process APPXBUNDLE files
    foreach ($package in $appxbundlePackages) {
        Write-Host "`nProcessing APPXBUNDLE: $(Split-Path $package -Leaf)" -ForegroundColor Cyan
        try {
            dism /Online /Add-ProvisionedAppxPackage /PackagePath:$package /SkipLicense
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  ✓ Successfully added: $(Split-Path $package -Leaf)" -ForegroundColor Green
            } else {
                Write-Host "  ✗ Failed to add: $(Split-Path $package -Leaf)" -ForegroundColor Red
            }
        } catch {
            Write-Host "  ✗ Error adding $(Split-Path $package -Leaf) : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host "`nLive system application package installation completed!" -ForegroundColor Green
    Write-Host "Note: Some applications may require a restart to be fully available." -ForegroundColor Yellow
    
    Read-Host "`nPress Enter to continue"
}

function Remove-LiveAppxPackages {
    Write-Host "`nRemoving Application Packages from Live System" -ForegroundColor Cyan
    Write-Host "WARNING: Changes will be applied immediately!" -ForegroundColor Red
    
    try {
        dism /Online /Get-ProvisionedAppxPackages
    } catch {
        Write-Host "Error getting packages: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nEnter application package names separated by commas:"
    $packageInput = Read-Host "Application package names to remove"
    if (-not $packageInput) {
        return
    }
    
    $packages = $packageInput -split ',' | ForEach-Object { $_.Trim() }
    
    $confirm = Read-Host "`nAre you sure you want to remove these packages from the live system? (y/N)"
    if ($confirm.ToLower() -eq 'y') {
        Write-Host "`nRemoving application packages from live system..." -ForegroundColor Yellow
        foreach ($package in $packages) {
            Write-Host "Removing: $package" -ForegroundColor Gray
            try {
                dism /Online /Remove-ProvisionedAppxPackage /PackageName:$package
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "✓ Successfully removed: $package" -ForegroundColor Green
                } else {
                    Write-Host "✗ Failed to remove: $package" -ForegroundColor Red
                }
            } catch {
                Write-Host "✗ Error removing $package : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    Read-Host "`nPress Enter to continue"
}

function Show-LiveAvailableFeatures {
    Write-Host "Getting available features from live system..." -ForegroundColor Yellow
    try {
        Write-Host "`nAvailable Windows Features - Live System:" -ForegroundColor Cyan
        Write-Host "=========================================" -ForegroundColor Cyan
        dism /Online /Get-Features /Format:Table
    } catch {
        Write-Host "Error getting live features: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-LiveFeatureInfo {
    param([string]$FeatureName)
    
    try {
        Write-Host "`nDetailed information for feature: $FeatureName (Live System)" -ForegroundColor Cyan
        dism /Online /Get-FeatureInfo /FeatureName:$FeatureName
    } catch {
        Write-Host "Error getting live feature info: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Enable-LiveFeatures {
    Write-Host "`nEnabling Windows Features on Live System" -ForegroundColor Cyan
    Write-Host "WARNING: Changes will be applied immediately!" -ForegroundColor Red
    
    try {
        dism /Online /Get-Features /Format:Table
    } catch {
        Write-Host "Error getting features: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nEnter feature names separated by commas:"
    $featureInput = Read-Host "Features to enable"
    if (-not $featureInput) {
        return
    }
    
    $features = $featureInput -split ',' | ForEach-Object { $_.Trim() }
    
    $confirm = Read-Host "`nAre you sure you want to enable these features on the live system? (y/N)"
    if ($confirm.ToLower() -eq 'y') {
        Write-Host "`nEnabling features on live system..." -ForegroundColor Yellow
        foreach ($feature in $features) {
            Write-Host "Enabling: $feature" -ForegroundColor Gray
            try {
                dism /Online /Enable-Feature /FeatureName:$feature /All
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "✓ Successfully enabled: $feature" -ForegroundColor Green
                } else {
                    Write-Host "✗ Failed to enable: $feature" -ForegroundColor Red
                }
            } catch {
                Write-Host "✗ Error enabling $feature : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    Read-Host "`nPress Enter to continue"
}

function Disable-LiveFeatures {
    Write-Host "`nDisabling Windows Features on Live System" -ForegroundColor Cyan
    Write-Host "WARNING: Changes will be applied immediately!" -ForegroundColor Red
    
    try {
        dism /Online /Get-Features /Format:Table
    } catch {
        Write-Host "Error getting features: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nEnter feature names separated by commas:"
    $featureInput = Read-Host "Features to disable"
    if (-not $featureInput) {
        return
    }
    
    $features = $featureInput -split ',' | ForEach-Object { $_.Trim() }
    
    $removePayload = Read-Host "Remove feature payload completely? (y/N)"
    $removeCompletely = $removePayload.ToLower() -eq 'y'
    
    $confirm = Read-Host "`nAre you sure you want to disable these features on the live system? (y/N)"
    if ($confirm.ToLower() -eq 'y') {
        Write-Host "`nDisabling features on live system..." -ForegroundColor Yellow
        foreach ($feature in $features) {
            Write-Host "Disabling: $feature" -ForegroundColor Gray
            try {
                if ($removeCompletely) {
                    dism /Online /Disable-Feature /FeatureName:$feature /Remove
                } else {
                    dism /Online /Disable-Feature /FeatureName:$feature
                }
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "✓ Successfully disabled: $feature" -ForegroundColor Green
                } else {
                    Write-Host "✗ Failed to disable: $feature" -ForegroundColor Red
                }
            } catch {
                Write-Host "✗ Error disabling $feature : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    Read-Host "`nPress Enter to continue"
}

function Manage-DesktopFiles {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "       Desktop Files Management" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "CURRENT CONFIGURATION:" -ForegroundColor Magenta
        Write-Host "Install Image Path: " -NoNewline
        Write-Host $Global:InstallWimPath -ForegroundColor Green
        
        $fileType = if ($Global:InstallWimPath -like "*.esd") { "ESD" } else { "WIM" }
        Write-Host "File Type: " -NoNewline
        Write-Host $fileType -ForegroundColor Green
        
        Write-Host "Mount Path: " -NoNewline
        Write-Host "$Global:MountPath (Ready)" -ForegroundColor Green
        
        Write-Host "Desktop Files Path: " -NoNewline
        if ($Global:DesktopFilesPath) {
            Write-Host $Global:DesktopFilesPath -ForegroundColor Green
        } else {
            Write-Host "Not Set (Will prompt when needed)" -ForegroundColor Yellow
        }
        
        try {
            $indexes = Get-WimIndexes
            Write-Host "Available Indexes: " -NoNewline
            if ($indexes.Count -gt 0) {
                Write-Host "$($indexes.Count) found" -ForegroundColor Green
            } else {
                Write-Host "None found" -ForegroundColor Red
            }
        } catch {
            Write-Host "Available Indexes: " -NoNewline
            Write-Host "Unable to read" -ForegroundColor Red
        }
        Write-Host ""
        
        Write-Host "DESKTOP FILES OPERATIONS:" -ForegroundColor Cyan
        Write-Host "1. Add Files to Desktop (Specific Index)" -ForegroundColor Yellow
        Write-Host "2. Add Files to Desktop (All Indexes)" -ForegroundColor Yellow
        Write-Host "3. View Desktop Files (Specific Index)" -ForegroundColor Green
        Write-Host "4. Remove Files from Specific Index" -ForegroundColor Red
        Write-Host "5. Remove Files from All Indexes" -ForegroundColor Red
        Write-Host "6. Return to Main Menu" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Select an option (1-6): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { Add-FilesToDesktopSpecificIndex }
            "2" { Add-FilesToDesktopAllIndexes }
            "3" { View-DesktopFiles }
            "4" { Remove-FilesFromSpecificIndex }
            "5" { Remove-FilesFromAllIndexes }
            "6" { return }
            default {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

function Add-FilesToDesktopSpecificIndex {
    # Check and set desktop files path BEFORE mounting
    if (-not $Global:DesktopFilesPath) {
        Write-Host "`nDesktop files path is not set. Please configure it first." -ForegroundColor Yellow
        $validPath = Get-DesktopFilesPathInteractive
        if (-not $validPath) {
            Write-Host "Cannot proceed without desktop files path." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    }
    
    Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
    dism /Get-WimInfo /WimFile:$Global:InstallWimPath
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexNumber = Read-Host "`nEnter the index number"
    Add-FilesToDesktopForIndex -IndexNumber $indexNumber
    Read-Host "Press Enter to continue"
}

function Add-FilesToDesktopAllIndexes {
    # Check and set desktop files path BEFORE mounting
    if (-not $Global:DesktopFilesPath) {
        Write-Host "`nDesktop files path is not set. Please configure it first." -ForegroundColor Yellow
        $validPath = Get-DesktopFilesPathInteractive
        if (-not $validPath) {
            Write-Host "Cannot proceed without desktop files path." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    }
    
    $indexes = Get-WimIndexes
    if ($indexes.Count -eq 0) {
        Write-Host "No indexes found in the image." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nThis will add files to desktop folders in ALL indexes." -ForegroundColor Yellow
    Write-Host "Found $($indexes.Count) indexes to process." -ForegroundColor Cyan
    
    $proceed = Read-Host "Proceed with adding files to all indexes? (Y/n)"
    if ($proceed.ToLower() -eq 'n') {
        return
    }
    
    # Get the files to copy ONCE before processing any indexes
    $sourceFiles = Get-DesktopFilesToCopy
    if (-not $sourceFiles -or $sourceFiles.Count -eq 0) {
        Write-Host "No files selected for copying." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Now process each index with the already-selected files
    foreach ($index in $indexes) {
        Write-Host "`n--- Processing Index ${index} ---" -ForegroundColor Magenta
        Add-FilesToDesktopForIndexWithFiles -IndexNumber $index -SourceFiles $sourceFiles
    }
    
    Write-Host "`nCompleted adding files to all indexes!" -ForegroundColor Green
    Read-Host "Press Enter to continue"
}

function Add-FilesToDesktopForIndexWithFiles {
    param(
        [string]$IndexNumber,
        [array]$SourceFiles
    )
    
    try {
        Write-Host "`nMounting WIM/ESD index ${IndexNumber}..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$IndexNumber /MountDir:$Global:MountPath /CheckIntegrity
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${IndexNumber}"
        }
        
        # Find desktop folders and copy files
        $desktopFolders = Find-DesktopFolders -MountPath $Global:MountPath
        
        if ($desktopFolders.Count -eq 0) {
            Write-Host "No desktop folders found in mounted image." -ForegroundColor Yellow
        } else {
            Copy-FilesToDesktopFolders -SourceFiles $SourceFiles -DesktopFolders $desktopFolders -IndexNumber $IndexNumber
        }
        
        Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Unmounting"
        dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "SUCCESS: Desktop files added to index ${IndexNumber}" -ForegroundColor Green
        } else {
            throw "Failed to commit changes for index ${IndexNumber}"
        }
        
    } catch {
        Write-Host "ERROR processing index ${IndexNumber} : $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Attempting to discard changes and unmount..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    }
}

function Get-DesktopFilesPathInteractive {
    do {
        $path = Read-Host "Enter the path to your desktop files folder"
        $validatedPath = Test-DirectoryPath -Path $path -PathType "Desktop Files"
        if ($validatedPath) {
            $Global:DesktopFilesPath = $validatedPath
            $save = Read-Host "Save this desktop files path for future use? (Y/n)"
            if ($save.ToLower() -ne 'n') {
                Write-Host "Desktop files path saved for future use." -ForegroundColor Green
            }
            return $validatedPath
        } else {
            Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
            $retry = Read-Host "Try again? (Y/n)"
            if ($retry.ToLower() -eq 'n') {
                return $null
            }
        }
    } while ($true)
}

function View-DesktopFiles {
    Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
    dism /Get-WimInfo /WimFile:$Global:InstallWimPath
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexNumber = Read-Host "`nEnter the index number to view desktop files"
    
    $Global:CurrentOperation = "Viewing Desktop Files"
    $Global:CleanupRequired = $true
    
    try {
        Write-Host "`nMounting WIM/ESD index ${indexNumber} (Read-Only)..." -ForegroundColor Yellow
        
        # Direct DISM call to show output properly
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$indexNumber /MountDir:$Global:MountPath /CheckIntegrity /ReadOnly
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${indexNumber}"
        }
        
        $desktopFolders = Find-DesktopFolders -MountPath $Global:MountPath
        
        if ($desktopFolders.Count -eq 0) {
            Write-Host "`nNo desktop folders found in this index." -ForegroundColor Yellow
        } else {
            Write-Host "`nDesktop Files in Index ${indexNumber}:" -ForegroundColor Cyan
            Write-Host "===================================" -ForegroundColor Cyan
            
            foreach ($desktopFolder in $desktopFolders) {
                $folderName = ($desktopFolder -split '\\')[-2..-1] -join '\'
                Write-Host "`nFolder: ${folderName}" -ForegroundColor Green
                
                $files = Get-ChildItem -Path $desktopFolder -Recurse -File -ErrorAction SilentlyContinue
                if ($files.Count -eq 0) {
                    Write-Host "  (No files)" -ForegroundColor Gray
                } else {
                    foreach ($file in $files) {
                        $relativePath = $file.FullName.Substring($desktopFolder.Length + 1)
                        $sizeKB = [math]::Round($file.Length / 1KB, 2)
                        Write-Host "  ${relativePath} ($sizeKB KB)" -ForegroundColor White
                    }
                }
            }
        }
        
        Read-Host "Press Enter to continue"
        
        Write-Host "`nUnmounting image..." -ForegroundColor Yellow
        dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        
    } catch {
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
    
    Read-Host "Press Enter to continue"
}

function Remove-FilesFromSpecificIndex {
    Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
    dism /Get-WimInfo /WimFile:$Global:InstallWimPath
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexNumber = Read-Host "`nEnter the index number"
    
    $Global:CurrentOperation = "Removing Desktop Files"
    $Global:CleanupRequired = $true
    
    try {
        Remove-DesktopFilesFromIndex -IndexNumber $indexNumber -IsMultiIndex $false
    } catch {
        Write-Host "Error during desktop file removal: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
    
    Read-Host "Press Enter to continue"
}

function Remove-FilesFromAllIndexes {
    $indexes = Get-WimIndexes
    if ($indexes.Count -eq 0) {
        Write-Host "No indexes found in the image." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nThis will remove files from desktop folders in ALL indexes." -ForegroundColor Yellow
    Write-Host "Found $($indexes.Count) indexes to process." -ForegroundColor Cyan
    
    $proceed = Read-Host "Proceed with removing files from all indexes? (Y/n)"
    if ($proceed.ToLower() -eq 'n') {
        return
    }
    
    # Show available indexes and let user choose the base index
    Write-Host "`nSelect which Windows edition to use as base for file selection:" -ForegroundColor Yellow
    Write-Host "(Different editions may have different files in desktop folders)" -ForegroundColor Gray
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
    } catch {
        Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        $baseIndex = Read-Host "`nEnter the index number of the edition to use as base"
        if ($indexes -contains $baseIndex) {
            break
        } else {
            Write-Host "Invalid index. Please select from the available indexes above." -ForegroundColor Red
        }
    } while ($true)
    
    $Global:CurrentOperation = "Removing Desktop Files"
    $Global:CleanupRequired = $true
    
    # Clear the global array
    $Global:RemovedDesktopFiles = @()
    
    try {
        Write-Host "`n--- Processing Base Index $baseIndex (File Selection) ---" -ForegroundColor Magenta
        Remove-DesktopFilesFromIndex -IndexNumber $baseIndex -IsMultiIndex $true -IsFirstIndex $true
        
        foreach ($index in $indexes) {
            if ($index -ne $baseIndex) {
                Write-Host "`n--- Processing Index ${index} ---" -ForegroundColor Magenta
                Remove-DesktopFilesFromIndex -IndexNumber $index -IsMultiIndex $true -IsFirstIndex $false
            }
        }
    } catch {
        Write-Host "Error during desktop file removal: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
    
    Read-Host "Press Enter to continue"
}

function Remove-DesktopFilesFromIndex {
    param(
        [string]$IndexNumber,
        [bool]$IsMultiIndex,
        [bool]$IsFirstIndex = $true
    )
    
    try {
        Write-Host "`nMounting WIM/ESD index ${IndexNumber}..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$IndexNumber /MountDir:$Global:MountPath /CheckIntegrity
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${IndexNumber}"
        }
        
        if ($IsMultiIndex -and -not $IsFirstIndex) {
            # Apply previously selected files to remove
            if ($Global:RemovedDesktopFiles.Count -gt 0) {
                Write-Host "`nApplying previously selected desktop files to remove..." -ForegroundColor Yellow
                
                $successCount = 0
                $failCount = 0
                
                foreach ($fileInfo in $Global:RemovedDesktopFiles) {
                    # Reconstruct the path in the mounted image
                    $desktopFolder = $Global:MountPath + "\Users\Default\Desktop"
                    $targetFile = Join-Path $desktopFolder $fileInfo.RelativePath
                    
                    try {
                        if (Test-Path $targetFile) {
                            Remove-Item $targetFile -Force
                            Write-Host "  ✓ Removed: $($fileInfo.RelativePath)" -ForegroundColor Green
                            $successCount++
                        } else {
                            Write-Host "  ⚠ Not found: $($fileInfo.RelativePath)" -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Host "  ✗ Error removing $($fileInfo.RelativePath): $($_.Exception.Message)" -ForegroundColor Red
                        $failCount++
                    }
                }
                
                Write-Host "`nSummary for index ${IndexNumber}: ${successCount} removed, ${failCount} failed" -ForegroundColor Cyan
            }
        } else {
            # First index - get user selection
            $desktopFolders = Find-DesktopFolders -MountPath $Global:MountPath
            
            if ($desktopFolders.Count -eq 0) {
                Write-Host "`nNo desktop folders found in this index." -ForegroundColor Yellow
                dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                return
            }
            
            # Collect all files and show them
            $allFiles = @()
            $fileIndex = 1
            
            Write-Host "`nDesktop Files in Index ${IndexNumber}:" -ForegroundColor Cyan
            Write-Host "===================================" -ForegroundColor Cyan
            
            foreach ($desktopFolder in $desktopFolders) {
                $folderName = ($desktopFolder -split '\\')[-2..-1] -join '\'
                $files = Get-ChildItem -Path $desktopFolder -Recurse -File -ErrorAction SilentlyContinue
                
                foreach ($file in $files) {
                    $relativePath = $file.FullName.Substring($desktopFolder.Length + 1)
                    $sizeKB = [math]::Round($file.Length / 1KB, 2)
                    Write-Host "${fileIndex}. [$folderName] ${relativePath} ($sizeKB KB)" -ForegroundColor White
                    
                    $allFiles += @{
                        Index = $fileIndex
                        FullPath = $file.FullName
                        RelativePath = $relativePath
                        DesktopFolder = $desktopFolder
                        FolderName = $folderName
                    }
                    $fileIndex++
                }
            }
            
            if ($allFiles.Count -eq 0) {
                Write-Host "No files found in desktop folders." -ForegroundColor Yellow
                dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                return
            }
            
            Write-Host "`nEnter file numbers to remove (comma-separated, e.g., 1,3,5-7):" -ForegroundColor Yellow
            Write-Host "Or enter 'ALL' to remove all files:" -ForegroundColor Yellow
            $selection = Read-Host "Files to remove"
            
            if (-not $selection.Trim()) {
                Write-Host "No selection made. Cancelling operation." -ForegroundColor Yellow
                dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                return
            }
            
            $filesToRemove = @()
            
            if ($selection.ToUpper() -eq 'ALL') {
                $filesToRemove = $allFiles
            } else {
                # Parse selection
                $parts = $selection -split ','
                
                foreach ($part in $parts) {
                    $part = $part.Trim()
                    if ($part -match '^(\d+)-(\d+)$') {
                        # Range like 5-7
                        $start = [int]$matches[1]
                        $end = [int]$matches[2]
                        for ($i = $start; $i -le $end; $i++) {
                            $file = $allFiles | Where-Object { $_.Index -eq $i }
                            if ($file) { $filesToRemove += $file }
                        }
                    } elseif ($part -match '^\d+$') {
                        # Single number
                        $num = [int]$part
                        $file = $allFiles | Where-Object { $_.Index -eq $num }
                        if ($file) { $filesToRemove += $file }
                    }
                }
            }
            
            if ($filesToRemove.Count -eq 0) {
                Write-Host "No valid files selected for removal." -ForegroundColor Yellow
                dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                return
            }
            
            if ($IsMultiIndex) {
                # Store for use with other indexes
                $Global:RemovedDesktopFiles = $filesToRemove
                $applyToAll = Read-Host "`nApply these desktop file removals to all indexes? (Y/n)"
                if ($applyToAll.ToLower() -eq 'n') {
                    $Global:RemovedDesktopFiles = @()
                }
            }
            
            # Remove files from this index
            $successCount = 0
            $failCount = 0
            
            Write-Host "Removing $($filesToRemove.Count) file(s) from index ${IndexNumber}..." -ForegroundColor Yellow
            
            foreach ($fileInfo in $filesToRemove) {
                try {
                    if (Test-Path $fileInfo.FullPath) {
                        Remove-Item $fileInfo.FullPath -Force
                        Write-Host "  ✓ Removed: $($fileInfo.RelativePath)" -ForegroundColor Green
                        $successCount++
                    } else {
                        Write-Host "  ⚠ Not found: $($fileInfo.RelativePath)" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "  ✗ Error removing $($fileInfo.RelativePath): $($_.Exception.Message)" -ForegroundColor Red
                    $failCount++
                }
            }
            
            Write-Host "`nSummary for index ${IndexNumber}: ${successCount} removed, ${failCount} failed" -ForegroundColor Cyan
        }
        
        Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Unmounting"
        dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "SUCCESS: Changes committed for index ${IndexNumber}" -ForegroundColor Green
        } else {
            throw "Failed to commit changes for index ${IndexNumber}"
        }
        
    } catch {
        Write-Host "ERROR processing index ${IndexNumber} : $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Attempting to discard changes and unmount..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    }
}

function Remove-SelectedFilesFromIndex {
    param([string]$IndexNumber, [array]$FilesToRemove)
    
    try {
        Write-Host "`nMounting WIM/ESD index ${IndexNumber}..." -ForegroundColor Yellow
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$IndexNumber /MountDir:$Global:MountPath /CheckIntegrity
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${IndexNumber}"
        }
        
        $successCount = 0
        $failCount = 0
        
        Write-Host "Removing $($FilesToRemove.Count) file(s) from index ${IndexNumber}..." -ForegroundColor Yellow
        
        foreach ($fileInfo in $FilesToRemove) {
            # Reconstruct the path in the mounted image
            $desktopFolder = $Global:MountPath + "\Users\Default\Desktop"
            $targetFile = Join-Path $desktopFolder $fileInfo.RelativePath
            
            try {
                if (Test-Path $targetFile) {
                    Remove-Item $targetFile -Force
                    Write-Host "  ✓ Removed: $($fileInfo.RelativePath)" -ForegroundColor Green
                    $successCount++
                } else {
                    Write-Host "  ⚠ Not found: $($fileInfo.RelativePath)" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "  ✗ Error removing $($fileInfo.RelativePath): $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        Write-Host "`nSummary for index ${IndexNumber}: ${successCount} removed, ${failCount} failed" -ForegroundColor Cyan
        
        Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
        dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Changes committed successfully for index ${IndexNumber}" -ForegroundColor Green
        } else {
            throw "Failed to commit changes for index ${IndexNumber}"
        }
        
    } catch {
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    }
}

function Show-AvailableFeatures {
    param([string]$ImagePath)
    
    Write-Host "Getting available features (this may take a moment)..." -ForegroundColor Yellow
    try {
        Write-Host "`nAvailable Windows Features:" -ForegroundColor Cyan
        Write-Host "============================" -ForegroundColor Cyan
        dism /Image:$ImagePath /Get-Features /Format:Table
    } catch {
        Write-Host "Error getting features: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-FeatureInfo {
   param([string]$ImagePath, [string]$FeatureName)
   
   try {
       Write-Host "`nDetailed information for feature: $FeatureName" -ForegroundColor Cyan
       dism /Image:$ImagePath /Get-FeatureInfo /FeatureName:$FeatureName
   } catch {
       Write-Host "Error getting feature info: $($_.Exception.Message)" -ForegroundColor Red
   }
}

function Enable-Features {
    param([string]$ImagePath, [bool]$IsMultiIndex)
    
    Write-Host "`nEnabling Windows Features" -ForegroundColor Cyan
    Write-Host "Getting available features for reference..." -ForegroundColor Yellow
    try {
        Write-Host "`nAvailable Windows Features:" -ForegroundColor Cyan
        Write-Host "============================" -ForegroundColor Cyan
        dism /Image:$ImagePath /Get-Features /Format:Table
    } catch {
        Write-Host "Error getting features: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nEnter feature names separated by commas (or press Enter to skip):"
    Write-Host "Example: Microsoft-Windows-Subsystem-Linux,VirtualMachinePlatform"
   
    $featureInput = Read-Host "Features to enable"
    if (-not $featureInput) {
        return
    }
   
    $features = $featureInput -split ',' | ForEach-Object { $_.Trim() }
    
    if ($IsMultiIndex) {
        $Global:EnabledFeatures += $features
        $applyToAll = Read-Host "`nApply these features to all indexes? (Y/n)"
        if ($applyToAll.ToLower() -eq 'n') {
            $Global:EnabledFeatures = @()
        }
    }
   
    Write-Host "`nEnabling features..." -ForegroundColor Yellow
    foreach ($feature in $features) {
        Write-Host "Enabling: $feature" -ForegroundColor Gray
        try {
            dism /Image:$ImagePath /Enable-Feature /FeatureName:$feature /All
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully enabled: $feature" -ForegroundColor Green
            } else {
                Write-Host "✗ Failed to enable: $feature" -ForegroundColor Red
            }
        } catch {
            Write-Host "✗ Error enabling $feature : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
   
    Read-Host "`nPress Enter to continue"
}

function Disable-Features {
    param([string]$ImagePath, [bool]$IsMultiIndex)
    
    Write-Host "`nDisabling Windows Features" -ForegroundColor Cyan
    Write-Host "Getting available features for reference..." -ForegroundColor Yellow
    try {
        Write-Host "`nAvailable Windows Features:" -ForegroundColor Cyan
        Write-Host "============================" -ForegroundColor Cyan
        dism /Image:$ImagePath /Get-Features /Format:Table
    } catch {
        Write-Host "Error getting features: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nEnter feature names separated by commas (or press Enter to skip):"
    Write-Host "Example: Internet-Explorer-Optional-amd64,WindowsMediaPlayer"
    
    $featureInput = Read-Host "Features to disable"
    if (-not $featureInput) {
        return
    }
    
    $features = $featureInput -split ',' | ForEach-Object { $_.Trim() }
    
    $removePayload = Read-Host "Remove feature payload completely? (y/N)"
    $removeCompletely = $removePayload.ToLower() -eq 'y'
    
    if ($IsMultiIndex) {
        foreach ($feature in $features) {
            $Global:DisabledFeatures += @{
                Name = $feature
                RemovePayload = $removeCompletely
            }
        }
        $applyToAll = Read-Host "`nApply these features to all indexes? (Y/n)"
        if ($applyToAll.ToLower() -eq 'n') {
            $Global:DisabledFeatures = @()
        }
    }
    
    Write-Host "`nDisabling features..." -ForegroundColor Yellow
    foreach ($feature in $features) {
        Write-Host "Disabling: $feature" -ForegroundColor Gray
        try {
            if ($removeCompletely) {
                dism /Image:$ImagePath /Disable-Feature /FeatureName:$feature /Remove
            } else {
                dism /Image:$ImagePath /Disable-Feature /FeatureName:$feature
            }
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully disabled: $feature" -ForegroundColor Green
            } else {
                Write-Host "✗ Failed to disable: $feature" -ForegroundColor Red
            }
        } catch {
            Write-Host "✗ Error disabling $feature : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Read-Host "`nPress Enter to continue"
}

function Show-InstalledPackages {
    param([string]$ImagePath)
    
    Write-Host "Getting installed system packages (this may take a moment)..." -ForegroundColor Yellow
    try {
        Write-Host "`nInstalled System Packages (.cab/.msu):" -ForegroundColor Cyan
        Write-Host "=======================================" -ForegroundColor Cyan
        dism /Image:$ImagePath /Get-Packages
    } catch {
        Write-Host "Error getting system packages: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-PackageInfo {
    param([string]$ImagePath, [string]$PackageName)
    
    try {
        Write-Host "`nDetailed information for package: $PackageName" -ForegroundColor Cyan
        dism /Image:$ImagePath /Get-PackageInfo /PackageName:$PackageName
    } catch {
        Write-Host "Error getting package info: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Add-Packages {
    param([string]$ImagePath, [bool]$IsMultiIndex)
    
    Write-Host "`nAdding System Packages" -ForegroundColor Cyan
    $packagePath = $Global:PackagePath
    if (-not $packagePath) {
        do {
            $path = Read-Host "Enter the path to your system packages folder"
            $validatedPath = Test-DirectoryPath -Path $path -PathType "Package"
            if ($validatedPath) {
                $packagePath = $validatedPath
                $save = Read-Host "Save this package path for future use? (Y/n)"
                if ($save.ToLower() -ne 'n') {
                    $Global:PackagePath = $packagePath
                }
                break
            } else {
                Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
            }
        } while ($true)
    } else {
        Write-Host "Using saved package path: $packagePath" -ForegroundColor Green
        $useExisting = Read-Host "Use this path? (Y/n)"
        if ($useExisting.ToLower() -eq 'n') {
            do {
                $path = Read-Host "Enter the path to your system packages folder"
                $validatedPath = Test-DirectoryPath -Path $path -PathType "Package"
                if ($validatedPath) {
                    $packagePath = $validatedPath
                    break
                } else {
                    Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
                }
            } while ($true)
        }
    }
    
    Write-Host "`nScanning for system packages in: $packagePath" -ForegroundColor Yellow
    $cabFiles = Get-ChildItem -Path $packagePath -Filter "*.cab" -Recurse -ErrorAction SilentlyContinue
    $msuFiles = Get-ChildItem -Path $packagePath -Filter "*.msu" -Recurse -ErrorAction SilentlyContinue
    
    $packages = @()
    if ($cabFiles) {
        $packages += $cabFiles.FullName
    }
    if ($msuFiles) {
        $packages += $msuFiles.FullName
    }
    
    $packages = $packages | Where-Object { $_ -and $_.Trim() }
    
    if ($packages.Count -eq 0) {
        Write-Host "No .cab or .msu files found in: $packagePath" -ForegroundColor Red
        Read-Host "`nPress Enter to continue"
        return
    }
    
    Write-Host "Found $($packages.Count) system package file(s):" -ForegroundColor Green
    foreach ($pkg in $packages) {
        Write-Host "  - $(Split-Path $pkg -Leaf)" -ForegroundColor Gray
    }
    
    $proceed = Read-Host "`nProceed with adding these system packages? (Y/n)"
    if ($proceed.ToLower() -eq 'n') {
        return
    }
    
    if ($IsMultiIndex) {
        $Global:AddedPackages += $packages
        $applyToAll = Read-Host "`nApply these system packages to all indexes? (Y/n)"
        if ($applyToAll.ToLower() -eq 'n') {
            $Global:AddedPackages = @()
        }
    }
   
    Write-Host "`nAdding system packages..." -ForegroundColor Yellow
    foreach ($package in $packages) {
        Write-Host "Adding: $(Split-Path $package -Leaf)" -ForegroundColor Gray
        try {
            dism /Image:$ImagePath /Add-Package /PackagePath:$package
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully added: $(Split-Path $package -Leaf)" -ForegroundColor Green
            } else {
                Write-Host "✗ Failed to add: $(Split-Path $package -Leaf)" -ForegroundColor Red
            }
        } catch {
            Write-Host "✗ Error adding $(Split-Path $package -Leaf) : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
   
    Read-Host "`nPress Enter to continue"
}

function Remove-Packages {
    param([string]$ImagePath, [bool]$IsMultiIndex)
    
    Write-Host "`nRemoving System Packages" -ForegroundColor Cyan
    Write-Host "Getting installed system packages for reference..." -ForegroundColor Yellow
    try {
        Write-Host "`nInstalled System Packages:" -ForegroundColor Cyan
        Write-Host "===========================" -ForegroundColor Cyan
        dism /Image:$ImagePath /Get-Packages
    } catch {
        Write-Host "Error getting system packages: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nEnter system package names separated by commas (or press Enter to skip):"
    Write-Host "Example: Package_for_KB123456~31bf3856ad364e35~amd64~~1.0.0.0"
    
    $packageInput = Read-Host "System package names to remove"
    if (-not $packageInput) {
        return
    }
    
    $packages = $packageInput -split ',' | ForEach-Object { $_.Trim() }
    
    if ($IsMultiIndex) {
        $Global:RemovedPackages += $packages
        $applyToAll = Read-Host "`nApply these system package removals to all indexes? (Y/n)"
        if ($applyToAll.ToLower() -eq 'n') {
            $Global:RemovedPackages = @()
        }
    }
    
    Write-Host "`nRemoving system packages..." -ForegroundColor Yellow
    foreach ($package in $packages) {
        Write-Host "Removing: $package" -ForegroundColor Gray
        try {
            dism /Image:$Global:MountPath /Remove-Package /PackageName:$package
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully removed: $package" -ForegroundColor Green
            } else {
                Write-Host "✗ Failed to remove: $package" -ForegroundColor Red
            }
        } catch {
            Write-Host "✗ Error removing $package : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Read-Host "`nPress Enter to continue"
}

function Remove-WimIndex {
   if (-not (Test-RequiredPaths)) {
       Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
       Read-Host "Press Enter to continue"
       return
   }
   
   if (-not (Test-MountPathEmpty)) {
       Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
       Read-Host "Press Enter to continue"
       return
   }
   
   Write-Host "Available indexes in WIM/ESD:" -ForegroundColor Cyan
   try {
       dism /Get-WimInfo /WimFile:$Global:InstallWimPath
   } catch {
       Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
       Read-Host "Press Enter to continue"
       return
   }
   
   $availableIndexes = Get-WimIndexes
   if ($availableIndexes.Count -eq 0) {
       Write-Host "No indexes found in file." -ForegroundColor Red
       Read-Host "Press Enter to continue"
       return
   }
   
   Write-Host "`nYou can enter multiple indexes separated by commas (e.g., 1,3,6)" -ForegroundColor Yellow
   $indexInput = Read-Host "`nEnter the index number(s) to remove"
   
   if (-not $indexInput.Trim()) {
       Write-Host "No indexes specified." -ForegroundColor Yellow
       Read-Host "Press Enter to continue"
       return
   }
   
   $indexesToRemove = @()
   $invalidIndexes = @()
   
   foreach ($indexStr in ($indexInput -split ',' | ForEach-Object { $_.Trim() })) {
       if ($indexStr -match '^\d+$') {
           $indexNum = [int]$indexStr
           if ($availableIndexes -contains $indexNum.ToString()) {
               $indexesToRemove += $indexNum
           } else {
               $invalidIndexes += $indexStr
           }
       } else {
           $invalidIndexes += $indexStr
       }
   }
   
   if ($invalidIndexes.Count -gt 0) {
       Write-Host "Invalid or non-existent indexes: $($invalidIndexes -join ', ')" -ForegroundColor Red
   }
   
   if ($indexesToRemove.Count -eq 0) {
       Write-Host "No valid indexes to remove." -ForegroundColor Red
       Read-Host "Press Enter to continue"
       return
   }
   
   $indexesToRemove = $indexesToRemove | Sort-Object -Unique -Descending
   
   Write-Host "`nIndexes to be removed: $($indexesToRemove -join ', ')" -ForegroundColor Yellow
   $confirm = Read-Host "Are you sure you want to remove these indexes? (y/N)"
   
   if ($confirm.ToLower() -eq 'y') {
       $successCount = 0
       $failCount = 0
       
       foreach ($index in $indexesToRemove) {
           try {
               Write-Host "`nRemoving index ${index}..." -ForegroundColor Yellow
               dism /Delete-Image /ImageFile:$Global:InstallWimPath /Index:$index
               
               if ($LASTEXITCODE -eq 0) {
                   Write-Host "✓ Successfully removed index ${index}" -ForegroundColor Green
                   $successCount++
               } else {
                   Write-Host "✗ Failed to remove index ${index}" -ForegroundColor Red
                   $failCount++
               }
           } catch {
               Write-Host "✗ Error removing index ${index} : $($_.Exception.Message)" -ForegroundColor Red
               $failCount++
           }
       }
       
       Write-Host "`nSummary:" -ForegroundColor Cyan
       Write-Host "Successfully removed: ${successCount} indexes" -ForegroundColor Green
       if ($failCount -gt 0) {
           Write-Host "Failed to remove: ${failCount} indexes" -ForegroundColor Red
       }
   } else {
       Write-Host "Operation cancelled." -ForegroundColor Yellow
   }
   
   Read-Host "Press Enter to continue"
}

function Add-WimIndex {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "        Add WIM/ESD Index" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "CURRENT CONFIGURATION:" -ForegroundColor Magenta
        Write-Host "Target Image: " -NoNewline
        Write-Host $Global:InstallWimPath -ForegroundColor Green
        
        $fileType = if ($Global:InstallWimPath -like "*.esd") { "ESD" } else { "WIM" }
        Write-Host "Target Type: " -NoNewline
        Write-Host $fileType -ForegroundColor Green
        
        Write-Host "Mount Path: " -NoNewline
        Write-Host "$Global:MountPath (Ready)" -ForegroundColor Green
        
        try {
            $indexes = Get-WimIndexes
            Write-Host "Current Indexes: " -NoNewline
            Write-Host "$($indexes.Count) found" -ForegroundColor Green
        } catch {
            Write-Host "Current Indexes: " -NoNewline
            Write-Host "Unable to read" -ForegroundColor Red
        }
        Write-Host ""
        
        Write-Host "ADD INDEX OPTIONS:" -ForegroundColor Cyan
        Write-Host "1. Export from another WIM file" -ForegroundColor Yellow
        Write-Host "2. Export from ESD file" -ForegroundColor Yellow
        Write-Host "3. Return to Index Management" -ForegroundColor Red
        Write-Host ""
        Write-Host "Select an option (1-3): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { Export-FromWim; return }
            "2" { Export-FromEsd; return }
            "3" { return }
            default { 
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

function Manage-Indexes {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "       Index Management" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "CURRENT CONFIGURATION:" -ForegroundColor Magenta
        Write-Host "Install Image Path: " -NoNewline
        Write-Host $Global:InstallWimPath -ForegroundColor Green
        
        $fileType = if ($Global:InstallWimPath -like "*.esd") { "ESD" } else { "WIM" }
        Write-Host "File Type: " -NoNewline
        Write-Host $fileType -ForegroundColor Green
        
        Write-Host "Mount Path: " -NoNewline
        Write-Host "$Global:MountPath (Ready)" -ForegroundColor Green
        try {
            $indexes = Get-WimIndexes
            Write-Host "Current Indexes: " -NoNewline
            if ($indexes.Count -gt 0) {
                Write-Host "$($indexes.Count) found" -ForegroundColor Green
            } else {
                Write-Host "None found" -ForegroundColor Red
            }
        } catch {
            Write-Host "Current Indexes: " -NoNewline
            Write-Host "Unable to read" -ForegroundColor Red
        }
        Write-Host ""
        
        Write-Host "INDEX OPERATIONS:" -ForegroundColor Cyan
        Write-Host "1. View Current Indexes" -ForegroundColor Yellow
        Write-Host "2. Add Index" -ForegroundColor Yellow
        Write-Host "3. Remove Index" -ForegroundColor Yellow
        Write-Host "4. Return to Main Menu" -ForegroundColor Red
        Write-Host ""
        Write-Host "Select an option (1-4): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" {
                Write-Host "`nCurrent Indexes:" -ForegroundColor Cyan
                Write-Host "=================" -ForegroundColor Cyan
                try {
                    dism /Get-WimInfo /WimFile:$Global:InstallWimPath
                } catch {
                    Write-Host "Error getting file information." -ForegroundColor Red
                }
                Read-Host "`nPress Enter to continue"
            }
            "2" { Add-WimIndex }
            "3" { Remove-WimIndex }
            "4" { return }
            default {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

function Manage-Conversions {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "      File Conversion" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "CURRENT CONFIGURATION:" -ForegroundColor Magenta
        Write-Host "Install Image Path: " -NoNewline
        Write-Host $Global:InstallWimPath -ForegroundColor Green
        
        $fileType = if ($Global:InstallWimPath -like "*.esd") { "ESD" } else { "WIM" }
        Write-Host "Current File Type: " -NoNewline
        Write-Host $fileType -ForegroundColor Green
        
        Write-Host "Mount Path: " -NoNewline
        Write-Host "$Global:MountPath (Ready)" -ForegroundColor Green
        
        Write-Host "Backup Path: " -NoNewline
        if ($Global:BackupPath) {
            Write-Host $Global:BackupPath -ForegroundColor Green
        } else {
            Write-Host "Not Set (Will use default)" -ForegroundColor Yellow
        }
        try {
            $fileInfo = Get-Item $Global:InstallWimPath -ErrorAction SilentlyContinue
            if ($fileInfo) {
                $sizeGB = [math]::Round($fileInfo.Length / 1GB, 2)
                Write-Host "Current File Size: " -NoNewline
                Write-Host "$sizeGB GB" -ForegroundColor Cyan
            }
        } catch {
            Write-Host "Current File Size: " -NoNewline
            Write-Host "Unable to read" -ForegroundColor Red
        }
        Write-Host ""
        
        Write-Host "CONVERSION OPTIONS:" -ForegroundColor Cyan
        Write-Host "1. Convert to WIM" -ForegroundColor $(if ($fileType -eq "ESD") { "Yellow" } else { "DarkGray" })
        Write-Host "2. Convert to ESD" -ForegroundColor $(if ($fileType -eq "WIM") { "Yellow" } else { "DarkGray" })
        Write-Host "3. Recompress current format" -ForegroundColor Yellow
        Write-Host "4. Return to Main Menu" -ForegroundColor Red
        Write-Host ""
        
        if ($fileType -eq "WIM") {
            Write-Host "Note: File is already in WIM format" -ForegroundColor Gray
        } elseif ($fileType -eq "ESD") {
            Write-Host "Note: File is already in ESD format" -ForegroundColor Gray
        }
        Write-Host ""
        Write-Host "Select an option (1-4): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { 
                if ($fileType -eq "wim") {
                    Write-Host "File is already in WIM format." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue"
                } else {
                    Convert-EsdToWim
                }
            }
            "2" {
                if ($fileType -eq "esd") {
                    Write-Host "File is already in ESD format." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue"
                } else {
                    Convert-WimToEsd
                }
            }
            "3" {
                if ($fileType -eq "WIM") {
                    Convert-WimToWim
                } else {
                    Convert-EsdToEsd
                }
            }
            "4" { return }
            default {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

function Manage-ISOCreation {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "       ISO Creation" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "CURRENT CONFIGURATION:" -ForegroundColor Magenta
        Write-Host "Source Path: " -NoNewline
        Write-Host $Global:ExtractedISOPath -ForegroundColor Green
        Write-Host "ISO Output Path: " -NoNewline
        if ($Global:ISOOutputPath) {
            Write-Host $Global:ISOOutputPath -ForegroundColor Green
        } else {
            Write-Host "Not Set (Will prompt when needed)" -ForegroundColor Yellow
        }
        $uefiBootFile = Join-Path $Global:ExtractedISOPath "efi\Microsoft\boot\efisys.bin"
        $biosBootFile = Join-Path $Global:ExtractedISOPath "boot\etfsboot.com"
        $isolinuxBootFile = Join-Path $Global:ExtractedISOPath "isolinux\isolinux.bin"
        
        Write-Host "`nAVAILABLE BOOT MODES:" -ForegroundColor Magenta
        Write-Host "UEFI Boot Support: " -NoNewline
        if (Test-Path $uefiBootFile) {
            Write-Host "Available (efisys.bin found)" -ForegroundColor Green
        } else {
            Write-Host "Not Available" -ForegroundColor Red
        }
        
        Write-Host "BIOS Boot Support: " -NoNewline
        if (Test-Path $biosBootFile) {
            Write-Host "Available (etfsboot.com found)" -ForegroundColor Green
        } elseif (Test-Path $isolinuxBootFile) {
            Write-Host "Available (isolinux.bin found)" -ForegroundColor Green
        } else {
            Write-Host "Not Available" -ForegroundColor Red
        }
        
        $oscdimgPath = Get-OSCDImgPath
        Write-Host "OSCDIMG Tool: " -NoNewline
        if ($oscdimgPath) {
            Write-Host "Available" -ForegroundColor Green
        } else {
            Write-Host "Not Found (Install Windows ADK)" -ForegroundColor Red
        }
        Write-Host ""
        
        Write-Host "ISO CREATION OPTIONS:" -ForegroundColor Cyan
        
        $uefiAvailable = Test-Path $uefiBootFile
        $biosAvailable = (Test-Path $biosBootFile) -or (Test-Path $isolinuxBootFile)
        
        Write-Host "1. Create UEFI-only ISO" -ForegroundColor $(if ($uefiAvailable -and $oscdimgPath) { "Yellow" } else { "DarkGray" })
        Write-Host "2. Create BIOS-only ISO" -ForegroundColor $(if ($biosAvailable -and $oscdimgPath) { "Yellow" } else { "DarkGray" })
        Write-Host "3. Create Hybrid ISO (UEFI + BIOS)" -ForegroundColor $(if ($uefiAvailable -and $biosAvailable -and $oscdimgPath) { "Yellow" } else { "DarkGray" })
        Write-Host "4. Detect and Recommend Best Option" -ForegroundColor $(if ($oscdimgPath) { "Cyan" } else { "DarkGray" })
        Write-Host "5. Return to Main Menu" -ForegroundColor Red
        Write-Host ""
        
        if (-not $oscdimgPath) {
            Write-Host "WARNING: OSCDIMG tool not found. Please install Windows ADK first." -ForegroundColor Red
            Write-Host ""
        }
        
        Write-Host "Select an option (1-5): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { 
                if ($uefiAvailable -and $oscdimgPath) {
                    Create-ISO -BootMode "UEFI"
                } else {
                    Write-Host "UEFI boot files not available or OSCDIMG tool missing." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                }
            }
            "2" {
                if ($biosAvailable -and $oscdimgPath) {
                    Create-ISO -BootMode "BIOS"
                } else {
                    Write-Host "BIOS boot files not available or OSCDIMG tool missing." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                }
            }
            "3" {
                if ($uefiAvailable -and $biosAvailable -and $oscdimgPath) {
                    Create-ISO -BootMode "Hybrid"
                } else {
                    Write-Host "Both UEFI and BIOS boot files required for hybrid ISO, or OSCDIMG tool missing." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                }
            }
            "4" {
                if ($oscdimgPath) {
                    Recommend-ISOType
                } else {
                    Write-Host "OSCDIMG tool not found. Please install Windows ADK first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                }
            }
            "5" { return }
            default {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

function Get-BootParameters {
    param([string]$BootMode)
    
    $uefiBootFile = Join-Path $Global:ExtractedISOPath "efi\Microsoft\boot\efisys.bin"
    $biosBootFile = Join-Path $Global:ExtractedISOPath "boot\etfsboot.com"
    $isolinuxBootFile = Join-Path $Global:ExtractedISOPath "isolinux\isolinux.bin"
    
    switch ($BootMode) {
        "UEFI" {
            if (Test-Path $uefiBootFile) {
                return "-bootdata:2#p0,e,b$uefiBootFile#pEF,e,b$uefiBootFile"
            }
        }
        "BIOS" {
            if (Test-Path $biosBootFile) {
                return "-bootdata:1#p0,e,b$biosBootFile"
            } elseif (Test-Path $isolinuxBootFile) {
                return "-bootdata:1#p0,e,b$isolinuxBootFile"
            }
        }
        "Hybrid" {
            $biosFile = $null
            if (Test-Path $biosBootFile) {
                $biosFile = $biosBootFile
            } elseif (Test-Path $isolinuxBootFile) {
                $biosFile = $isolinuxBootFile
            }
            
            if ((Test-Path $uefiBootFile) -and $biosFile) {
                return "-bootdata:2#p0,e,b$biosFile#pEF,e,b$uefiBootFile"
            }
        }
    }
    
    return $null
}

function Get-ISOOutputPath {
    $outputPath = $Global:ISOOutputPath
    if (-not $outputPath) {
        do {
            $outputPath = Read-Host "Enter the output path for the new ISO file (including filename.iso)"
            $outputPath = $outputPath.Trim('"', "'")
            $outputPath = [System.Environment]::ExpandEnvironmentVariables($outputPath)
            
            $directory = Split-Path $outputPath -Parent
            if (Test-Path $directory) {
                break
            } else {
                Write-Host "Output directory does not exist: ${directory}" -ForegroundColor Red
                Write-Host "Please enter a path with an existing directory." -ForegroundColor Red
            }
        } while ($true)
    } else {
        Write-Host "Using preset output path: $outputPath" -ForegroundColor Green
        $usePreset = Read-Host "Use this path? (Y/n)"
        if ($usePreset.ToLower() -eq 'n') {
            do {
                $outputPath = Read-Host "Enter the output path for the new ISO file (including filename.iso)"
                $outputPath = $outputPath.Trim('"', "'")
                $outputPath = [System.Environment]::ExpandEnvironmentVariables($outputPath)
                
                $directory = Split-Path $outputPath -Parent
                if (Test-Path $directory) {
                    break
                } else {
                    Write-Host "Output directory does not exist: ${directory}" -ForegroundColor Red
                    Write-Host "Please enter a path with an existing directory." -ForegroundColor Red
                }
            } while ($true)
        }
    }
    
    return $outputPath
}

function Recommend-ISOType {
    Write-Host "`nAnalyzing available boot options..." -ForegroundColor Yellow
    
    $uefiBootFile = Join-Path $Global:ExtractedISOPath "efi\Microsoft\boot\efisys.bin"
    $biosBootFile = Join-Path $Global:ExtractedISOPath "boot\etfsboot.com"
    $isolinuxBootFile = Join-Path $Global:ExtractedISOPath "isolinux\isolinux.bin"
    
    $uefiAvailable = Test-Path $uefiBootFile
    $biosAvailable = (Test-Path $biosBootFile) -or (Test-Path $isolinuxBootFile)
    
    Write-Host "`nBoot Analysis Results:" -ForegroundColor Cyan
    Write-Host "======================" -ForegroundColor Cyan
    Write-Host "UEFI Support: " -NoNewline
    Write-Host $(if ($uefiAvailable) { "✓ Available" } else { "✗ Not Available" }) -ForegroundColor $(if ($uefiAvailable) { "Green" } else { "Red" })
    Write-Host "BIOS Support: " -NoNewline
    Write-Host $(if ($biosAvailable) { "✓ Available" } else { "✗ Not Available" }) -ForegroundColor $(if ($biosAvailable) { "Green" } else { "Red" })
    
    Write-Host "`nRECOMMENDATION:" -ForegroundColor Magenta
    if ($uefiAvailable -and $biosAvailable) {
        Write-Host "Create HYBRID ISO - Maximum compatibility with both modern UEFI and legacy BIOS systems" -ForegroundColor Green
        $recommended = "Hybrid"
    } elseif ($uefiAvailable) {
        Write-Host "Create UEFI-ONLY ISO - Modern systems only (recommended for Windows 10/11)" -ForegroundColor Yellow
        $recommended = "UEFI"
    } elseif ($biosAvailable) {
        Write-Host "Create BIOS-ONLY ISO - Legacy systems only" -ForegroundColor Yellow
        $recommended = "BIOS"
    } else {
        Write-Host "No valid boot files found - Cannot create bootable ISO" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $proceed = Read-Host "`nProceed with recommended $recommended ISO? (Y/n)"
    if ($proceed.ToLower() -ne 'n') {
        Create-ISO -BootMode $recommended
    }
}

function Test-ImageFile {
    param(
        [string]$Path,
        [string]$Type = "Image"
    )
    
    if (-not $Path) {
        Write-Host "$Type path is not set." -ForegroundColor Red
        return $false
    }
    
    if (-not (Test-Path $Path)) {
        Write-Host "$Type file not found: $Path" -ForegroundColor Red
        return $false
    }
    
    $extension = [System.IO.Path]::GetExtension($Path).ToLower()
    if ($extension -notin @(".wim", ".esd")) {
        Write-Host "$Type file should have .wim or .esd extension: $Path" -ForegroundColor Yellow
    }
    
    try {
        $output = dism /Get-WimInfo /WimFile:$Path 2>&1
        if ($LASTEXITCODE -ne 0) {
            if ($output -match "file cannot be found|not found|access.*denied") {
                Write-Host "$Type file cannot be accessed or found: $Path" -ForegroundColor Red
            } elseif ($output -match "invalid|corrupt|format") {
                Write-Host "$Type file appears to be invalid or corrupted: $Path" -ForegroundColor Red
            } else {
                Write-Host "$Type file validation failed: $Path" -ForegroundColor Red
                Write-Host "Error details: $($output -join ' ')" -ForegroundColor Gray
            }
            return $false
        }
        return $true
    } catch {
        Write-Host "Error validating $Type file: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Export-FromWim {
    if (-not (Test-ImageFile -Path $Global:InstallWimPath -Type "Install WIM")) {
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        $sourceWim = ""
        $sourceWimPath = Read-Host "Enter path to source extracted ISO folder"
        $sourceWimPath = $sourceWimPath.Trim('"', "'")
        $sourceWimPath = [System.Environment]::ExpandEnvironmentVariables($sourceWimPath)
        
        $validatedPath = Test-DirectoryPath -Path $sourceWimPath -PathType "Source ISO"
        if ($validatedPath) {
            $sourceWim = Join-Path $validatedPath "sources\install.wim"
            if (Test-Path $sourceWim) {
                Write-Host "Successfully found source install.wim at: $sourceWim" -ForegroundColor Green
                break
            } else {
                Write-Host "install.wim not found in sources folder. Please check the path." -ForegroundColor Red
                Write-Host "Expected: $sourceWim" -ForegroundColor Gray
            }
        } else {
            Write-Host "Please enter a valid existing directory path to the extracted ISO folder." -ForegroundColor Red
        }
        
        $retry = Read-Host "Try again? (Y/n)"
        if ($retry.ToLower() -eq 'n') {
            return
        }
    } while ($true)
    
    if (-not (Test-ImageFile -Path $sourceWim -Type "Source WIM")) {
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nAvailable indexes in source WIM:" -ForegroundColor Yellow
    Write-Host "Source: $sourceWim" -ForegroundColor Gray
    try {
        dism /Get-WimInfo /WimFile:$sourceWim
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Failed to read source WIM information." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    } catch {
        Write-Host "Error reading source WIM: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $availableIndexes = @()
    try {
        $wimInfo = dism /Get-WimInfo /WimFile:$sourceWim
        foreach ($line in $wimInfo) {
            if ($line -match "Index : (\d+)") {
                $availableIndexes += $matches[1]
            }
        }
    } catch {
        Write-Host "Error getting source WIM indexes." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if ($availableIndexes.Count -eq 0) {
        Write-Host "No indexes found in source WIM." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nYou can enter multiple indexes separated by commas (e.g., 1,3,6)" -ForegroundColor Yellow
    $indexInput = Read-Host "`nEnter source index number(s)"
    
    if (-not $indexInput.Trim()) {
        Write-Host "No indexes specified." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexesToExport = @()
    $invalidIndexes = @()
    
    foreach ($indexStr in ($indexInput -split ',' | ForEach-Object { $_.Trim() })) {
        if ($indexStr -match '^\d+$') {
            if ($availableIndexes -contains $indexStr) {
                $indexesToExport += $indexStr
            } else {
                $invalidIndexes += $indexStr
            }
        } else {
            $invalidIndexes += $indexStr
        }
    }
    
    if ($invalidIndexes.Count -gt 0) {
        Write-Host "Invalid or non-existent indexes: $($invalidIndexes -join ', ')" -ForegroundColor Red
    }
    
    if ($indexesToExport.Count -eq 0) {
        Write-Host "No valid indexes to export." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexesToExport = $indexesToExport | Sort-Object -Unique
    
    Write-Host "`nIndexes to be exported: $($indexesToExport -join ', ')" -ForegroundColor Yellow
    
    $customizeNames = $false
    $wimlibAvailable = Get-WimlibImagexPath
    
    if ($wimlibAvailable) {
        Write-Host "`nWimlib is available - full metadata customization supported" -ForegroundColor Green
        $keepDefaults = Read-Host "Would you like to customize the names, descriptions, and FLAGS? (y/N)"
        $customizeNames = $keepDefaults.ToLower() -eq 'y'
    } else {
        Write-Host "`nWimlib not available - only name customization will be supported" -ForegroundColor Yellow
        $keepDefaults = Read-Host "Would you like to customize the names? (y/N)"
        $customizeNames = $keepDefaults.ToLower() -eq 'y'
    }
    
    # Store custom names/metadata
    $customMetadata = @{}
    if ($customizeNames) {
        foreach ($sourceIndex in $indexesToExport) {
            $indexInfo = dism /Get-WimInfo /WimFile:$sourceWim /Index:$sourceIndex 2>$null
            $defaultName = "Exported Index ${sourceIndex}"
            $defaultFlags = "Unknown"
            if ($indexInfo) {
                $nameMatch = $indexInfo | Select-String "Name\s*:\s*(.+)"
                if ($nameMatch) {
                    $defaultName = $nameMatch.Matches[0].Groups[1].Value.Trim()
                }
            }
            
            # Try to get current FLAGS if wimlib is available
            if ($wimlibAvailable) {
                try {
                    $wimlibInfo = & $wimlibAvailable info $sourceWim $sourceIndex 2>$null
                    foreach ($line in $wimlibInfo) {
                        if ($line -match "^FLAGS\s*:\s*(.+)") {
                            $defaultFlags = $matches[1].Trim()
                            break
                        }
                    }
                } catch {
                    # Ignore wimlib errors, keep default
                }
            }
            
            Write-Host "`nCustomizing index ${sourceIndex} (Default: '$defaultName'):" -ForegroundColor Cyan
            $userName = Read-Host "Enter new name (Press Enter for default)"
            if (-not $userName.Trim()) { $userName = $defaultName }
            
            if ($wimlibAvailable) {
                $userDesc = Read-Host "Enter description (Press Enter for same as name)"
                if (-not $userDesc.Trim()) { $userDesc = $userName }
                
                Write-Host ""
                Write-Host "FLAGS corresponds to the Windows edition (e.g., Professional, Enterprise, Education)" -ForegroundColor Cyan
                Write-Host "Current FLAGS: $defaultFlags" -ForegroundColor Gray
                $userFlags = Read-Host "Enter FLAGS/edition (Press Enter for current)"
                if (-not $userFlags.Trim()) { $userFlags = $defaultFlags }
                
                $customMetadata[$sourceIndex] = @{
                    OriginalName = $defaultName
                    ExportName = $userName
                    NewName = $userName
                    NewDescription = $userDesc
                    NewFlags = $userFlags
                    UseWimlib = $true
                }
            } else {
                $customMetadata[$sourceIndex] = @{
                    OriginalName = $defaultName
                    ExportName = $userName
                    UseWimlib = $false
                }
            }
        }
    }
    
    try {
        $successCount = 0
        $failCount = 0
        
        foreach ($sourceIndex in $indexesToExport) {
            $indexInfo = dism /Get-WimInfo /WimFile:$sourceWim /Index:$sourceIndex 2>$null
            $defaultName = "Exported Index ${sourceIndex}"
            if ($indexInfo) {
                $nameMatch = $indexInfo | Select-String "Name\s*:\s*(.+)"
                if ($nameMatch) {
                    $defaultName = $nameMatch.Matches[0].Groups[1].Value.Trim()
                }
            }
            
            # Determine export name
            $exportName = $defaultName
            if ($customizeNames -and $customMetadata.ContainsKey($sourceIndex)) {
                $exportName = $customMetadata[$sourceIndex].ExportName
            }
            
            Write-Host "`nExporting index ${sourceIndex}..." -ForegroundColor Yellow
            Write-Host "Export name: ${exportName}" -ForegroundColor Cyan
            Write-Host "From: $sourceWim (Index: $sourceIndex)" -ForegroundColor Gray
            Write-Host "To: $Global:InstallWimPath" -ForegroundColor Gray
            
            # Export with the chosen name
            dism /Export-Image /SourceImageFile:$sourceWim /SourceIndex:$sourceIndex /DestinationImageFile:$Global:InstallWimPath /DestinationName:$exportName /CheckIntegrity
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully exported: ${exportName}" -ForegroundColor Green
                $successCount++
                
                # Update additional metadata with wimlib if available and user customized
                if ($customizeNames -and $customMetadata.ContainsKey($sourceIndex)) {
                    $metadata = $customMetadata[$sourceIndex]
                    if ($metadata.UseWimlib) {
                        Write-Host "Updating full metadata with wimlib..." -ForegroundColor Cyan
                        $updateSuccess = Update-ExportedIndexMetadata -IndexName $exportName -NewName $metadata.NewName -NewDescription $metadata.NewDescription -NewFlags $metadata.NewFlags
                        if (-not $updateSuccess) {
                            Write-Host "Warning: Export succeeded but metadata update failed" -ForegroundColor Yellow
                        }
                    }
                }
            } else {
                Write-Host "✗ Failed to export index ${sourceIndex}" -ForegroundColor Red
                $failCount++
            }
        }
        
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "Successfully exported: ${successCount} indexes" -ForegroundColor Green
        if ($failCount -gt 0) {
            Write-Host "Failed to export: ${failCount} indexes" -ForegroundColor Red
        }
        
        if ($customizeNames) {
            if ($wimlibAvailable) {
                Write-Host "Full metadata (name, description, FLAGS, EDITIONID, DISPLAYNAME, DISPLAYDESCRIPTION) updated using wimlib" -ForegroundColor Cyan
            } else {
                Write-Host "Only names were customized (wimlib required for description and FLAGS updates)" -ForegroundColor Yellow
            }
        }
        
    } catch {
        Write-Host "Error during export: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Read-Host "Press Enter to continue"
}

function Export-FromEsd {
    if (-not (Test-ImageFile -Path $Global:InstallWimPath -Type "Install WIM")) {
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        $sourcePath = Read-Host "Enter path to extracted ISO folder or ESD file"
        $sourcePath = $sourcePath.Trim('"', "'")
        $sourcePath = [System.Environment]::ExpandEnvironmentVariables($sourcePath)
        
        if (Test-Path $sourcePath) {
            $sourceEsd = ""
            
            if (Test-Path $sourcePath -PathType Container) {
                $possibleEsd = Join-Path $sourcePath "sources\install.esd"
                if (Test-Path $possibleEsd) {
                    $sourceEsd = $possibleEsd
                    Write-Host "Found install.esd at: $sourceEsd" -ForegroundColor Green
                } else {
                    Write-Host "install.esd not found in sources folder." -ForegroundColor Red
                    Write-Host "Expected: $possibleEsd" -ForegroundColor Gray
                }
            } else {
                if ($sourcePath -like "*.esd") {
                    $sourceEsd = $sourcePath
                    Write-Host "Found ESD file: $sourceEsd" -ForegroundColor Green
                } else {
                    Write-Host "File does not have .esd extension. Please select an ESD file or extracted ISO folder." -ForegroundColor Red
                }
            }
            
            if ($sourceEsd) {
                break
            }
        } else {
            Write-Host "Path not found: $sourcePath" -ForegroundColor Red
        }
        
        $retry = Read-Host "Try again? (Y/n)"
        if ($retry.ToLower() -eq 'n') {
            return
        }
    } while ($true)
    
    Write-Host "`nValidating ESD file..." -ForegroundColor Yellow
    try {
        $esdInfo = dism /Get-WimInfo /WimFile:$sourceEsd
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Failed to read ESD file information. The file may be corrupted or invalid." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    } catch {
        Write-Host "Error reading ESD file: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nAvailable indexes in source ESD:" -ForegroundColor Yellow
    Write-Host "Source: $sourceEsd" -ForegroundColor Gray
    try {
        dism /Get-WimInfo /WimFile:$sourceEsd
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Failed to read source ESD information." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    } catch {
        Write-Host "Error reading source ESD: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $availableIndexes = @()
    try {
        $esdInfo = dism /Get-WimInfo /WimFile:$sourceEsd
        foreach ($line in $esdInfo) {
            if ($line -match "Index : (\d+)") {
                $availableIndexes += $matches[1]
            }
        }
    } catch {
        Write-Host "Error getting source ESD indexes." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if ($availableIndexes.Count -eq 0) {
        Write-Host "No indexes found in source ESD." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nYou can enter multiple indexes separated by commas (e.g., 1,3,6)" -ForegroundColor Yellow
    $indexInput = Read-Host "`nEnter source index number(s)"
    
    if (-not $indexInput.Trim()) {
        Write-Host "No indexes specified." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexesToExport = @()
    $invalidIndexes = @()
    
    foreach ($indexStr in ($indexInput -split ',' | ForEach-Object { $_.Trim() })) {
        if ($indexStr -match '^\d+$') {
            if ($availableIndexes -contains $indexStr) {
                $indexesToExport += $indexStr
            } else {
                $invalidIndexes += $indexStr
            }
        } else {
            $invalidIndexes += $indexStr
        }
    }
    
    if ($invalidIndexes.Count -gt 0) {
        Write-Host "Invalid or non-existent indexes: $($invalidIndexes -join ', ')" -ForegroundColor Red
    }
    
    if ($indexesToExport.Count -eq 0) {
        Write-Host "No valid indexes to export." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexesToExport = $indexesToExport | Sort-Object -Unique
    
    Write-Host "`nIndexes to be exported: $($indexesToExport -join ', ')" -ForegroundColor Yellow
    
    $customizeNames = $false
    $wimlibAvailable = Get-WimlibImagexPath
    
    if ($wimlibAvailable) {
        Write-Host "`nWimlib is available - full metadata customization supported" -ForegroundColor Green
        $keepDefaults = Read-Host "Would you like to customize the names, descriptions, and FLAGS? (y/N)"
        $customizeNames = $keepDefaults.ToLower() -eq 'y'
    } else {
        Write-Host "`nWimlib not available - only name customization will be supported" -ForegroundColor Yellow
        $keepDefaults = Read-Host "Would you like to customize the names? (y/N)"
        $customizeNames = $keepDefaults.ToLower() -eq 'y'
    }
    
    # Store custom names/metadata
    $customMetadata = @{}
    if ($customizeNames) {
        foreach ($sourceIndex in $indexesToExport) {
            $indexInfo = dism /Get-WimInfo /WimFile:$sourceEsd /Index:$sourceIndex 2>$null
            $defaultName = "Exported Index ${sourceIndex}"
            $defaultFlags = "Unknown"
            if ($indexInfo) {
                $nameMatch = $indexInfo | Select-String "Name\s*:\s*(.+)"
                if ($nameMatch) {
                    $defaultName = $nameMatch.Matches[0].Groups[1].Value.Trim()
                }
            }
            
            # Try to get current FLAGS if wimlib is available
            if ($wimlibAvailable) {
                try {
                    $wimlibInfo = & $wimlibAvailable info $sourceEsd $sourceIndex 2>$null
                    foreach ($line in $wimlibInfo) {
                        if ($line -match "^FLAGS\s*:\s*(.+)") {
                            $defaultFlags = $matches[1].Trim()
                            break
                        }
                    }
                } catch {
                    # Ignore wimlib errors, keep default
                }
            }
            
            Write-Host "`nCustomizing index ${sourceIndex} (Default: '$defaultName'):" -ForegroundColor Cyan
            $userName = Read-Host "Enter new name (Press Enter for default)"
            if (-not $userName.Trim()) { $userName = $defaultName }
            
            if ($wimlibAvailable) {
                $userDesc = Read-Host "Enter description (Press Enter for same as name)"
                if (-not $userDesc.Trim()) { $userDesc = $userName }
                
                Write-Host ""
                Write-Host "FLAGS corresponds to the Windows edition (e.g., Professional, Enterprise, Education)" -ForegroundColor Cyan
                Write-Host "Current FLAGS: $defaultFlags" -ForegroundColor Gray
                $userFlags = Read-Host "Enter FLAGS/edition (Press Enter for current)"
                if (-not $userFlags.Trim()) { $userFlags = $defaultFlags }
                
                $customMetadata[$sourceIndex] = @{
                    OriginalName = $defaultName
                    ExportName = $userName
                    NewName = $userName
                    NewDescription = $userDesc
                    NewFlags = $userFlags
                    UseWimlib = $true
                }
            } else {
                $customMetadata[$sourceIndex] = @{
                    OriginalName = $defaultName
                    ExportName = $userName
                    UseWimlib = $false
                }
            }
        }
    }
    
    try {
        $successCount = 0
        $failCount = 0
        
        foreach ($sourceIndex in $indexesToExport) {
            $indexInfo = dism /Get-WimInfo /WimFile:$sourceEsd /Index:$sourceIndex 2>$null
            $defaultName = "Exported Index ${sourceIndex}"
            if ($indexInfo) {
                $nameMatch = $indexInfo | Select-String "Name\s*:\s*(.+)"
                if ($nameMatch) {
                    $defaultName = $nameMatch.Matches[0].Groups[1].Value.Trim()
                }
            }
            
            # Determine export name
            $exportName = $defaultName
            if ($customizeNames -and $customMetadata.ContainsKey($sourceIndex)) {
                $exportName = $customMetadata[$sourceIndex].ExportName
            }
            
            Write-Host "`nExporting index ${sourceIndex} from ESD..." -ForegroundColor Yellow
            Write-Host "Export name: ${exportName}" -ForegroundColor Cyan
            Write-Host "From: $sourceEsd (Index: $sourceIndex)" -ForegroundColor Gray
            Write-Host "To: $Global:InstallWimPath" -ForegroundColor Gray
            Write-Host "Note: This process may take several minutes." -ForegroundColor Cyan

            # Export with the chosen name
            dism /Export-Image /SourceImageFile:$sourceEsd /SourceIndex:$sourceIndex /DestinationImageFile:$Global:InstallWimPath /DestinationName:$exportName /CheckIntegrity
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully exported ${exportName} from ESD" -ForegroundColor Green
                $successCount++
                
                # Update additional metadata with wimlib if available and user customized
                if ($customizeNames -and $customMetadata.ContainsKey($sourceIndex)) {
                    $metadata = $customMetadata[$sourceIndex]
                    if ($metadata.UseWimlib) {
                        Write-Host "Updating full metadata with wimlib..." -ForegroundColor Cyan
                        $updateSuccess = Update-ExportedIndexMetadata -IndexName $exportName -NewName $metadata.NewName -NewDescription $metadata.NewDescription -NewFlags $metadata.NewFlags
                        if (-not $updateSuccess) {
                            Write-Host "Warning: Export succeeded but metadata update failed" -ForegroundColor Yellow
                        }
                    }
                }
            } else {
                Write-Host "✗ Failed to export index ${sourceIndex}" -ForegroundColor Red
                $failCount++
            }
        }
        
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "Successfully exported: ${successCount} indexes" -ForegroundColor Green
        if ($failCount -gt 0) {
            Write-Host "Failed to export: ${failCount} indexes" -ForegroundColor Red
        }
        
        if ($customizeNames) {
            if ($wimlibAvailable) {
                Write-Host "Full metadata (name, description, FLAGS, EDITIONID, DISPLAYNAME, DISPLAYDESCRIPTION) updated using wimlib" -ForegroundColor Cyan
            } else {
                Write-Host "Only names were customized (wimlib required for description and FLAGS updates)" -ForegroundColor Yellow
            }
        }
        
    } catch {
        Write-Host "Error during export: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Read-Host "Press Enter to continue"
}

function Convert-WimToEsd {
    if (-not (Test-RequiredPaths) -or -not (Test-MountPathEmpty)) {
        Write-Host "Please ensure all required paths are set and mount path is empty." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-ImageFile -Path $Global:InstallWimPath -Type "Install WIM")) {
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nWIM to ESD Conversion" -ForegroundColor Cyan
    Write-Host "=====================" -ForegroundColor Cyan
    Write-Host "`nCurrent WIM Information:" -ForegroundColor Yellow
    
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
    } catch {
        Write-Host "Error getting WIM information." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nBackup Options:" -ForegroundColor Magenta
    Write-Host "The conversion process is safe and doesn't modify the original file." -ForegroundColor Gray
    Write-Host "However, you can create a backup for extra safety." -ForegroundColor Gray

    $createBackup = Read-Host "`nCreate backup of original file before conversion? (y/N)"
    $shouldBackup = $createBackup.ToLower() -eq 'y'
    $backupPath = $null

    if ($shouldBackup) {
        if ($Global:BackupPath) {
            $fileName = Split-Path $Global:InstallWimPath -Leaf
            $defaultBackupPath = Join-Path $Global:BackupPath ($fileName + ".backup")
        } else {
            $defaultBackupPath = $Global:InstallWimPath + ".backup"
        }
        
        Write-Host "`nDefault backup path: ${defaultBackupPath}" -ForegroundColor Green
        
        do {
            $customPath = Read-Host "Enter new path for backup (Press Enter for default)"
            
            if (-not $customPath) {
                $finalBackupPath = $defaultBackupPath
            } else {
                $customPath = $customPath.Trim('"', "'")
                $customPath = [System.Environment]::ExpandEnvironmentVariables($customPath)
                
                if ($customPath -match '\.(wim|esd|backup)$') {
                    $directory = Split-Path $customPath -Parent
                    if ($directory -and (Test-Path $directory)) {
                        $finalBackupPath = $customPath
                    } elseif (-not $directory) {
                        $finalBackupPath = Join-Path (Get-Location) $customPath
                    } else {
                        Write-Host "Directory does not exist: ${directory}" -ForegroundColor Red
                        continue
                    }
                } else {
                    if (Test-Path $customPath) {
                        $fileName = Split-Path $Global:InstallWimPath -Leaf
                        $finalBackupPath = Join-Path $customPath ($fileName + ".backup")
                    } else {
                        Write-Host "Directory does not exist: ${customPath}" -ForegroundColor Red
                        continue
                    }
                }
            }
            
            Write-Host "Backup will be created at: ${finalBackupPath}" -ForegroundColor Cyan
            if (Test-Path $finalBackupPath) {
                Write-Host "Warning: Backup file already exists!" -ForegroundColor Yellow
                $overwriteBackup = Read-Host "Overwrite existing backup? (y/N)"
                if ($overwriteBackup.ToLower() -ne 'y') {
                    Write-Host "Skipping backup creation (existing backup will not be overwritten)." -ForegroundColor Yellow
                    $shouldBackup = $false
                    break
                }
            }
            
            $backupPath = $finalBackupPath
            break
            
        } while ($true)
    }
    
    Write-Host "`nCompression Options (ESD Format):" -ForegroundColor Magenta
    Write-Host "1. Maximum (Normal compression, Balanced conversion)" -ForegroundColor Yellow
    Write-Host "2. Fast (Quick compression, Faster conversion)" -ForegroundColor Yellow
    Write-Host "3. None (No compression, Fastest conversion)" -ForegroundColor Yellow
    Write-Host "4. Recovery (Best compression, Slowest conversion)" -ForegroundColor Green
    
    do {
        $choice = Read-Host "`nSelect compression type (1-4, default is 4)"
        if (-not $choice) { $choice = "4" }
        
        switch ($choice) {
            "1" { $compressionType = "maximum"; break }
            "2" { $compressionType = "fast"; break }
            "3" { $compressionType = "none"; break }
            "4" { $compressionType = "recovery"; break }
            default { Write-Host "Invalid selection. Please choose 1-4." -ForegroundColor Red; continue }
        }
        break
    } while ($true)
    
    $targetPath = $Global:InstallWimPath -replace '\.wim$', '.esd'
    
    Write-Host "`nConversion Details:" -ForegroundColor Cyan
    Write-Host "Source: $Global:InstallWimPath" -ForegroundColor Gray
    Write-Host "Target: ${targetPath}" -ForegroundColor Gray
    Write-Host "Compression: ${compressionType}" -ForegroundColor Gray
    if ($shouldBackup) {
        Write-Host "Backup will be created: ${backupPath}" -ForegroundColor Gray
    } else {
        Write-Host "No backup will be created" -ForegroundColor Gray
    }
    
    $proceed = Read-Host "`nProceed with conversion? (Y/n)"
    if ($proceed.ToLower() -eq 'n') { return }
    
    try {
        if (Test-Path $targetPath) {
            $overwrite = Read-Host "ESD file already exists. Overwrite? (y/N)"
            if ($overwrite.ToLower() -ne 'y') { return }
            Remove-Item $targetPath -Force
        }
        
        if ($shouldBackup) {
            Write-Host "`nCreating backup of original file..." -ForegroundColor Yellow
            try {
                Copy-Item $Global:InstallWimPath $backupPath -Force
                Write-Host "✓ Original file backed up to: ${backupPath}" -ForegroundColor Green
            } catch {
                Write-Host "ERROR: Could not create backup: $($_.Exception.Message)" -ForegroundColor Red
                $continueAnyway = Read-Host "Continue without backup? (y/N)"
                if ($continueAnyway.ToLower() -ne 'y') { return }
                $shouldBackup = $false
            }
        }
        
        $indexes = Get-WimIndexes
        if ($indexes.Count -eq 0) {
            Write-Host "No indexes found in source file." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        Write-Host "`nStarting conversion process..." -ForegroundColor Yellow
        Write-Host "This may take significant time depending on indexes and compression." -ForegroundColor Cyan
        
        $currentIndex = 0
        foreach ($index in $indexes) {
            $currentIndex++
            Write-Host "`nExporting index ${index} ($currentIndex of $($indexes.Count))..." -ForegroundColor Yellow
            
            $indexInfo = dism /Get-WimInfo /WimFile:$Global:InstallWimPath /Index:$index 2>$null
            $indexName = "Index ${index}"
            if ($indexInfo) {
                $nameMatch = $indexInfo | Select-String "Name\s*:\s*(.+)"
                if ($nameMatch) { $indexName = $nameMatch.Matches[0].Groups[1].Value.Trim() }
            }
            
            Write-Host "Processing: ${indexName}" -ForegroundColor Gray
            
            dism /Export-Image /SourceImageFile:$Global:InstallWimPath /SourceIndex:$index /DestinationImageFile:$targetPath /Compress:$compressionType /CheckIntegrity
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully exported: ${indexName}" -ForegroundColor Green
            } else {
                throw "Failed to export index ${index}"
            }
        }
        
        Write-Host "`nVerifying converted ESD file..." -ForegroundColor Yellow
        if (Test-ImageFile -Path $targetPath -Type "Converted ESD") {
            Write-Host "✓ ESD file verification successful" -ForegroundColor Green
            $Global:InstallWimPath = $targetPath
            Write-Host "✓ Updated install path to use ESD file" -ForegroundColor Green
            Write-Host "`nSUCCESS: WIM to ESD conversion completed!" -ForegroundColor Green
            Write-Host "Your ISO structure now uses install.esd instead of install.wim" -ForegroundColor Cyan
            
            if ($shouldBackup) {
                Write-Host "Original file has been safely backed up to: ${backupPath}" -ForegroundColor Cyan
            }
            
            $backupText = if ($shouldBackup) { "(backup will remain) " } else { "(no backup exists) " }
            $removeOriginal = Read-Host "`nRemove original file? ${backupText}(y/N)"
            if ($removeOriginal.ToLower() -eq 'y') {
                $originalToRemove = $Global:InstallWimPath -replace '\.esd$', '.wim'
                if (Test-Path $originalToRemove) {
                    Remove-Item $originalToRemove -Force
                    Write-Host "✓ Original install.wim removed" -ForegroundColor Green
                }
            }
        } else {
            throw "ESD file verification failed"
        }
        
    } catch {
        Write-Host "`nERROR: Conversion failed - $($_.Exception.Message)" -ForegroundColor Red
        if (Test-Path $targetPath) {
            Write-Host "Cleaning up incomplete ESD file..." -ForegroundColor Yellow
            try { Remove-Item $targetPath -Force } catch { 
                Write-Host "Warning: Could not remove incomplete ESD file: ${targetPath}" -ForegroundColor Yellow 
            }
        }
        Write-Host "Original WIM file remains unchanged." -ForegroundColor Green
    }
    
    Read-Host "Press Enter to continue"
}

function Convert-EsdToWim {
    if (-not (Test-RequiredPaths) -or -not (Test-MountPathEmpty)) {
        Write-Host "Please ensure all required paths are set and mount path is empty." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-ImageFile -Path $Global:InstallWimPath -Type "Install ESD")) {
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nESD to WIM Conversion" -ForegroundColor Cyan
    Write-Host "=====================" -ForegroundColor Cyan
    Write-Host "`nCurrent ESD Information:" -ForegroundColor Yellow
    
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
    } catch {
        Write-Host "Error getting ESD information." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nBackup Options:" -ForegroundColor Magenta
    Write-Host "The conversion process is safe and doesn't modify the original file." -ForegroundColor Gray
    Write-Host "However, you can create a backup for extra safety." -ForegroundColor Gray

    $createBackup = Read-Host "`nCreate backup of original file before conversion? (y/N)"
    $shouldBackup = $createBackup.ToLower() -eq 'y'
    $backupPath = $null

    if ($shouldBackup) {
        if ($Global:BackupPath) {
            $fileName = Split-Path $Global:InstallWimPath -Leaf
            $defaultBackupPath = Join-Path $Global:BackupPath ($fileName + ".backup")
        } else {
            $defaultBackupPath = $Global:InstallWimPath + ".backup"
        }
        
        Write-Host "`nDefault backup path: ${defaultBackupPath}" -ForegroundColor Green
        
        do {
            $customPath = Read-Host "Enter new path for backup (Press Enter for default)"
            
            if (-not $customPath) {
                $finalBackupPath = $defaultBackupPath
            } else {
                $customPath = $customPath.Trim('"', "'")
                $customPath = [System.Environment]::ExpandEnvironmentVariables($customPath)
                
                if ($customPath -match '\.(wim|esd|backup)$') {
                    $directory = Split-Path $customPath -Parent
                    if ($directory -and (Test-Path $directory)) {
                        $finalBackupPath = $customPath
                    } elseif (-not $directory) {
                        $finalBackupPath = Join-Path (Get-Location) $customPath
                    } else {
                        Write-Host "Directory does not exist: ${directory}" -ForegroundColor Red
                        continue
                    }
                } else {
                    if (Test-Path $customPath) {
                        $fileName = Split-Path $Global:InstallWimPath -Leaf
                        $finalBackupPath = Join-Path $customPath ($fileName + ".backup")
                    } else {
                        Write-Host "Directory does not exist: ${customPath}" -ForegroundColor Red
                        continue
                    }
                }
            }
            
            Write-Host "Backup will be created at: ${finalBackupPath}" -ForegroundColor Cyan
            if (Test-Path $finalBackupPath) {
                Write-Host "Warning: Backup file already exists!" -ForegroundColor Yellow
                $overwriteBackup = Read-Host "Overwrite existing backup? (y/N)"
                if ($overwriteBackup.ToLower() -ne 'y') {
                    Write-Host "Skipping backup creation (existing backup will not be overwritten)." -ForegroundColor Yellow
                    $shouldBackup = $false
                    break
                }
            }
            
            $backupPath = $finalBackupPath
            break
            
        } while ($true)
    }
    
    Write-Host "`nCompression Options (WIM Format):" -ForegroundColor Magenta
    Write-Host "1. Maximum (Normal compression, Balanced conversion)" -ForegroundColor Yellow
    Write-Host "2. Fast (Quick compression, Faster conversion)" -ForegroundColor Yellow
    Write-Host "3. None (No compression, Fastest conversion)" -ForegroundColor Yellow
    Write-Host "4. Recovery (Best compression, Slowest conversion)" -ForegroundColor Green
    
    do {
        $choice = Read-Host "`nSelect compression type (1-4, default is 4)"
        if (-not $choice) { $choice = "4" }
        
        switch ($choice) {
            "1" { $compressionType = "maximum"; break }
            "2" { $compressionType = "fast"; break }
            "3" { $compressionType = "none"; break }
            "4" { $compressionType = "recovery"; break }
            default { Write-Host "Invalid selection. Please choose 1-4." -ForegroundColor Red; continue }
        }
        break
    } while ($true)
    
    $targetPath = $Global:InstallWimPath -replace '\.esd$', '.wim'
    
    Write-Host "`nConversion Details:" -ForegroundColor Cyan
    Write-Host "Source: $Global:InstallWimPath" -ForegroundColor Gray
    Write-Host "Target: ${targetPath}" -ForegroundColor Gray
    Write-Host "Compression: ${compressionType}" -ForegroundColor Gray
    if ($shouldBackup) {
        Write-Host "Backup will be created: ${backupPath}" -ForegroundColor Gray
    } else {
        Write-Host "No backup will be created" -ForegroundColor Gray
    }
    
    $proceed = Read-Host "`nProceed with conversion? (Y/n)"
    if ($proceed.ToLower() -eq 'n') { return }
    
    try {
        if (Test-Path $targetPath) {
            $overwrite = Read-Host "WIM file already exists. Overwrite? (y/N)"
            if ($overwrite.ToLower() -ne 'y') { return }
            Remove-Item $targetPath -Force
        }
        
        if ($shouldBackup) {
            Write-Host "`nCreating backup of original file..." -ForegroundColor Yellow
            try {
                Copy-Item $Global:InstallWimPath $backupPath -Force
                Write-Host "✓ Original file backed up to: ${backupPath}" -ForegroundColor Green
            } catch {
                Write-Host "ERROR: Could not create backup: $($_.Exception.Message)" -ForegroundColor Red
                $continueAnyway = Read-Host "Continue without backup? (y/N)"
                if ($continueAnyway.ToLower() -ne 'y') { return }
                $shouldBackup = $false
            }
        }
        
        $indexes = Get-WimIndexes
        if ($indexes.Count -eq 0) {
            Write-Host "No indexes found in source file." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        Write-Host "`nStarting conversion process..." -ForegroundColor Yellow
        Write-Host "This may take significant time depending on indexes and compression." -ForegroundColor Cyan
        
        $currentIndex = 0
        foreach ($index in $indexes) {
            $currentIndex++
            Write-Host "`nExporting index ${index} ($currentIndex of $($indexes.Count))..." -ForegroundColor Yellow
            
            $indexInfo = dism /Get-WimInfo /WimFile:$Global:InstallWimPath /Index:$index 2>$null
            $indexName = "Index ${index}"
            if ($indexInfo) {
                $nameMatch = $indexInfo | Select-String "Name\s*:\s*(.+)"
                if ($nameMatch) { $indexName = $nameMatch.Matches[0].Groups[1].Value.Trim() }
            }
            
            Write-Host "Processing: ${indexName}" -ForegroundColor Gray
            
            dism /Export-Image /SourceImageFile:$Global:InstallWimPath /SourceIndex:$index /DestinationImageFile:$targetPath /Compress:$compressionType /CheckIntegrity
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully exported: ${indexName}" -ForegroundColor Green
            } else {
                throw "Failed to export index ${index}"
            }
        }
        
        Write-Host "`nVerifying converted WIM file..." -ForegroundColor Yellow
        if (Test-ImageFile -Path $targetPath -Type "Converted WIM") {
            Write-Host "✓ WIM file verification successful" -ForegroundColor Green
            $Global:InstallWimPath = $targetPath
            Write-Host "✓ Updated install path to use WIM file" -ForegroundColor Green
            Write-Host "`nSUCCESS: ESD to WIM conversion completed!" -ForegroundColor Green
            Write-Host "Your ISO structure now uses install.wim instead of install.esd" -ForegroundColor Cyan
            
            if ($shouldBackup) {
                Write-Host "Original file has been safely backed up to: ${backupPath}" -ForegroundColor Cyan
            }
            
            $backupText = if ($shouldBackup) { "(backup will remain) " } else { "(no backup exists) " }
            $removeOriginal = Read-Host "`nRemove original file? ${backupText}(y/N)"
            if ($removeOriginal.ToLower() -eq 'y') {
                $originalToRemove = $Global:InstallWimPath -replace '\.wim$', '.esd'
                if (Test-Path $originalToRemove) {
                    Remove-Item $originalToRemove -Force
                    Write-Host "✓ Original install.esd removed" -ForegroundColor Green
                }
            }
        } else {
            throw "WIM file verification failed"
        }
        
    } catch {
        Write-Host "`nERROR: Conversion failed - $($_.Exception.Message)" -ForegroundColor Red
        if (Test-Path $targetPath) {
            Write-Host "Cleaning up incomplete WIM file..." -ForegroundColor Yellow
            try { Remove-Item $targetPath -Force } catch { 
                Write-Host "Warning: Could not remove incomplete WIM file: ${targetPath}" -ForegroundColor Yellow 
            }
        }
        Write-Host "Original ESD file remains unchanged." -ForegroundColor Green
    }
    
    Read-Host "Press Enter to continue"
}

function Convert-WimToWim {
    if (-not (Test-RequiredPaths) -or -not (Test-MountPathEmpty)) {
        Write-Host "Please ensure all required paths are set and mount path is empty." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-ImageFile -Path $Global:InstallWimPath -Type "Install WIM")) {
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nWIM Recompression" -ForegroundColor Cyan
    Write-Host "==================" -ForegroundColor Cyan
    Write-Host "`nCurrent WIM Information:" -ForegroundColor Yellow
    
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
    } catch {
        Write-Host "Error getting WIM information." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nBackup Options:" -ForegroundColor Magenta
    Write-Host "The recompression process is safe and doesn't modify the original file." -ForegroundColor Gray
    Write-Host "However, you can create a backup for extra safety." -ForegroundColor Gray

    $createBackup = Read-Host "`nCreate backup of original file before recompression? (y/N)"
    $shouldBackup = $createBackup.ToLower() -eq 'y'
    $backupPath = $null

    if ($shouldBackup) {
        if ($Global:BackupPath) {
            $fileName = Split-Path $Global:InstallWimPath -Leaf
            $defaultBackupPath = Join-Path $Global:BackupPath ($fileName + ".backup")
        } else {
            $defaultBackupPath = $Global:InstallWimPath + ".backup"
        }
        
        Write-Host "`nDefault backup path: ${defaultBackupPath}" -ForegroundColor Green
        
        do {
            $customPath = Read-Host "Enter new path for backup (Press Enter for default)"
            
            if (-not $customPath) {
                $finalBackupPath = $defaultBackupPath
            } else {
                $customPath = $customPath.Trim('"', "'")
                $customPath = [System.Environment]::ExpandEnvironmentVariables($customPath)
                
                if ($customPath -match '\.(wim|esd|backup)$') {
                    $directory = Split-Path $customPath -Parent
                    if ($directory -and (Test-Path $directory)) {
                        $finalBackupPath = $customPath
                    } elseif (-not $directory) {
                        $finalBackupPath = Join-Path (Get-Location) $customPath
                    } else {
                        Write-Host "Directory does not exist: ${directory}" -ForegroundColor Red
                        continue
                    }
                } else {
                    if (Test-Path $customPath) {
                        $fileName = Split-Path $Global:InstallWimPath -Leaf
                        $finalBackupPath = Join-Path $customPath ($fileName + ".backup")
                    } else {
                        Write-Host "Directory does not exist: ${customPath}" -ForegroundColor Red
                        continue
                    }
                }
            }
            
            Write-Host "Backup will be created at: ${finalBackupPath}" -ForegroundColor Cyan
            if (Test-Path $finalBackupPath) {
                Write-Host "Warning: Backup file already exists!" -ForegroundColor Yellow
                $overwriteBackup = Read-Host "Overwrite existing backup? (y/N)"
                if ($overwriteBackup.ToLower() -ne 'y') {
                    Write-Host "Skipping backup creation (existing backup will not be overwritten)." -ForegroundColor Yellow
                    $shouldBackup = $false
                    break
                }
            }
            
            $backupPath = $finalBackupPath
            break
            
        } while ($true)
    }
    
    Write-Host "`nCompression Options (WIM Format):" -ForegroundColor Magenta
    Write-Host "1. Maximum (Normal compression, Balanced conversion)" -ForegroundColor Yellow
    Write-Host "2. Fast (Quick compression, Faster conversion)" -ForegroundColor Yellow
    Write-Host "3. None (No compression, Fastest conversion)" -ForegroundColor Yellow
    Write-Host "4. Recovery (Best compression, Slowest conversion)" -ForegroundColor Green
    
    do {
        $choice = Read-Host "`nSelect compression type (1-4, default is 4)"
        if (-not $choice) { $choice = "4" }
        
        switch ($choice) {
            "1" { $compressionType = "maximum"; break }
            "2" { $compressionType = "fast"; break }
            "3" { $compressionType = "none"; break }
            "4" { $compressionType = "recovery"; break }
            default { Write-Host "Invalid selection. Please choose 1-4." -ForegroundColor Red; continue }
        }
        break
    } while ($true)
    
    $targetPath = $Global:InstallWimPath -replace '\.wim$', '.new.wim'
    
    Write-Host "`nRecompression Details:" -ForegroundColor Cyan
    Write-Host "Source: $Global:InstallWimPath" -ForegroundColor Gray
    Write-Host "Target: ${targetPath}" -ForegroundColor Gray
    Write-Host "Compression: ${compressionType}" -ForegroundColor Gray
    if ($shouldBackup) {
        Write-Host "Backup will be created: ${backupPath}" -ForegroundColor Gray
    } else {
        Write-Host "No backup will be created" -ForegroundColor Gray
    }
    
    $proceed = Read-Host "`nProceed with recompression? (Y/n)"
    if ($proceed.ToLower() -eq 'n') { return }
    
    try {
        if (Test-Path $targetPath) {
            Remove-Item $targetPath -Force
        }
        
        if ($shouldBackup) {
            Write-Host "`nCreating backup of original file..." -ForegroundColor Yellow
            try {
                Copy-Item $Global:InstallWimPath $backupPath -Force
                Write-Host "✓ Original file backed up to: ${backupPath}" -ForegroundColor Green
            } catch {
                Write-Host "ERROR: Could not create backup: $($_.Exception.Message)" -ForegroundColor Red
                $continueAnyway = Read-Host "Continue without backup? (y/N)"
                if ($continueAnyway.ToLower() -ne 'y') { return }
                $shouldBackup = $false
            }
        }
        
        $indexes = Get-WimIndexes
        if ($indexes.Count -eq 0) {
            Write-Host "No indexes found in source file." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        Write-Host "`nStarting recompression process..." -ForegroundColor Yellow
        Write-Host "This may take significant time depending on indexes and compression." -ForegroundColor Cyan
        
        $currentIndex = 0
        foreach ($index in $indexes) {
            $currentIndex++
            Write-Host "`nExporting index ${index} ($currentIndex of $($indexes.Count))..." -ForegroundColor Yellow
            
            $indexInfo = dism /Get-WimInfo /WimFile:$Global:InstallWimPath /Index:$index 2>$null
            $indexName = "Index ${index}"
            if ($indexInfo) {
                $nameMatch = $indexInfo | Select-String "Name\s*:\s*(.+)"
                if ($nameMatch) { $indexName = $nameMatch.Matches[0].Groups[1].Value.Trim() }
            }
            
            Write-Host "Processing: ${indexName}" -ForegroundColor Gray
            
            dism /Export-Image /SourceImageFile:$Global:InstallWimPath /SourceIndex:$index /DestinationImageFile:$targetPath /Compress:$compressionType /CheckIntegrity
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully exported: ${indexName}" -ForegroundColor Green
            } else {
                throw "Failed to export index ${index}"
            }
        }
        
        Write-Host "`nVerifying recompressed WIM file..." -ForegroundColor Yellow
        if (Test-ImageFile -Path $targetPath -Type "Recompressed WIM") {
            Write-Host "✓ WIM file verification successful" -ForegroundColor Green
            Write-Host "`nSUCCESS: WIM recompression completed!" -ForegroundColor Green
            Write-Host "Your WIM file has been recompressed with the new compression settings" -ForegroundColor Cyan
            
            if ($shouldBackup) {
                Write-Host "Original file has been safely backed up to: ${backupPath}" -ForegroundColor Cyan
            }
            
            $backupText = if ($shouldBackup) { "(backup will remain) " } else { "(no backup exists) " }
            $removeOriginal = Read-Host "`nReplace original install.wim with recompressed version? ${backupText}(Y/n)"
            if ($removeOriginal.ToLower() -ne 'n') {
                Remove-Item $Global:InstallWimPath -Force
                Move-Item $targetPath $Global:InstallWimPath -Force
                Write-Host "✓ Original install.wim replaced with recompressed version" -ForegroundColor Green
            } else {
                Write-Host "Recompressed file saved as: ${targetPath}" -ForegroundColor Cyan
                Write-Host "Original file remains unchanged" -ForegroundColor Green
                Write-Host "You can manually replace it later if desired." -ForegroundColor Gray
            }
        } else {
            throw "WIM file verification failed"
        }
        
    } catch {
        Write-Host "`nERROR: Recompression failed - $($_.Exception.Message)" -ForegroundColor Red
        if (Test-Path $targetPath) {
            Write-Host "Cleaning up incomplete WIM file..." -ForegroundColor Yellow
            try { Remove-Item $targetPath -Force } catch { 
                Write-Host "Warning: Could not remove incomplete WIM file: ${targetPath}" -ForegroundColor Yellow 
            }
        }
        Write-Host "Original WIM file remains unchanged." -ForegroundColor Green
    }
    
    Read-Host "Press Enter to continue"
}

function Convert-EsdToEsd {
    if (-not (Test-RequiredPaths) -or -not (Test-MountPathEmpty)) {
        Write-Host "Please ensure all required paths are set and mount path is empty." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-ImageFile -Path $Global:InstallWimPath -Type "Install ESD")) {
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nESD Recompression" -ForegroundColor Cyan
    Write-Host "==================" -ForegroundColor Cyan
    Write-Host "`nCurrent ESD Information:" -ForegroundColor Yellow
    
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
    } catch {
        Write-Host "Error getting ESD information." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nBackup Options:" -ForegroundColor Magenta
    Write-Host "The recompression process is safe and doesn't modify the original file." -ForegroundColor Gray
    Write-Host "However, you can create a backup for extra safety." -ForegroundColor Gray

    $createBackup = Read-Host "`nCreate backup of original file before recompression? (y/N)"
    $shouldBackup = $createBackup.ToLower() -eq 'y'
    $backupPath = $null

    if ($shouldBackup) {
        if ($Global:BackupPath) {
            $fileName = Split-Path $Global:InstallWimPath -Leaf
            $defaultBackupPath = Join-Path $Global:BackupPath ($fileName + ".backup")
        } else {
            $defaultBackupPath = $Global:InstallWimPath + ".backup"
        }
        
        Write-Host "`nDefault backup path: ${defaultBackupPath}" -ForegroundColor Green
        
        do {
            $customPath = Read-Host "Enter new path for backup (Press Enter for default)"
            
            if (-not $customPath) {
                $finalBackupPath = $defaultBackupPath
            } else {
                $customPath = $customPath.Trim('"', "'")
                $customPath = [System.Environment]::ExpandEnvironmentVariables($customPath)
                
                if ($customPath -match '\.(wim|esd|backup)$') {
                    $directory = Split-Path $customPath -Parent
                    if ($directory -and (Test-Path $directory)) {
                        $finalBackupPath = $customPath
                    } elseif (-not $directory) {
                        $finalBackupPath = Join-Path (Get-Location) $customPath
                    } else {
                        Write-Host "Directory does not exist: ${directory}" -ForegroundColor Red
                        continue
                    }
                } else {
                    if (Test-Path $customPath) {
                        $fileName = Split-Path $Global:InstallWimPath -Leaf
                        $finalBackupPath = Join-Path $customPath ($fileName + ".backup")
                    } else {
                        Write-Host "Directory does not exist: ${customPath}" -ForegroundColor Red
                        continue
                    }
                }
            }
            
            Write-Host "Backup will be created at: ${finalBackupPath}" -ForegroundColor Cyan
            if (Test-Path $finalBackupPath) {
                Write-Host "Warning: Backup file already exists!" -ForegroundColor Yellow
                $overwriteBackup = Read-Host "Overwrite existing backup? (y/N)"
                if ($overwriteBackup.ToLower() -ne 'y') {
                    Write-Host "Skipping backup creation (existing backup will not be overwritten)." -ForegroundColor Yellow
                    $shouldBackup = $false
                    break
                }
            }
            
            $backupPath = $finalBackupPath
            break
            
        } while ($true)
    }
    
    Write-Host "`nCompression Options (ESD Format):" -ForegroundColor Magenta
    Write-Host "1. Maximum (Normal compression, Balanced conversion)" -ForegroundColor Yellow
    Write-Host "2. Fast (Quick compression, Faster conversion)" -ForegroundColor Yellow
    Write-Host "3. None (No compression, Fastest conversion)" -ForegroundColor Yellow
    Write-Host "4. Recovery (Best compression, Slowest conversion)" -ForegroundColor Green
    
    do {
        $choice = Read-Host "`nSelect compression type (1-4, default is 4)"
        if (-not $choice) { $choice = "4" }
        
        switch ($choice) {
            "1" { $compressionType = "maximum"; break }
            "2" { $compressionType = "fast"; break }
            "3" { $compressionType = "none"; break }
            "4" { $compressionType = "recovery"; break }
            default { Write-Host "Invalid selection. Please choose 1-4." -ForegroundColor Red; continue }
        }
        break
    } while ($true)
    
    $targetPath = $Global:InstallWimPath -replace '\.esd$', '.new.esd'
    
    Write-Host "`nRecompression Details:" -ForegroundColor Cyan
    Write-Host "Source: $Global:InstallWimPath" -ForegroundColor Gray
    Write-Host "Target: ${targetPath}" -ForegroundColor Gray
    Write-Host "Compression: ${compressionType}" -ForegroundColor Gray
    if ($shouldBackup) {
        Write-Host "Backup will be created: ${backupPath}" -ForegroundColor Gray
    } else {
        Write-Host "No backup will be created" -ForegroundColor Gray
    }
    
    $proceed = Read-Host "`nProceed with recompression? (Y/n)"
    if ($proceed.ToLower() -eq 'n') { return }
    
    try {
        if (Test-Path $targetPath) {
            Remove-Item $targetPath -Force
        }

        if ($shouldBackup) {
            Write-Host "`nCreating backup of original file..." -ForegroundColor Yellow
            try {
                Copy-Item $Global:InstallWimPath $backupPath -Force
                Write-Host "✓ Original file backed up to: ${backupPath}" -ForegroundColor Green
            } catch {
                Write-Host "ERROR: Could not create backup: $($_.Exception.Message)" -ForegroundColor Red
                $continueAnyway = Read-Host "Continue without backup? (y/N)"
                if ($continueAnyway.ToLower() -ne 'y') { return }
                $shouldBackup = $false
            }
        }
        
        $indexes = Get-WimIndexes
        if ($indexes.Count -eq 0) {
            Write-Host "No indexes found in source file." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        Write-Host "`nStarting recompression process..." -ForegroundColor Yellow
        Write-Host "This may take significant time depending on indexes and compression." -ForegroundColor Cyan
        
        $currentIndex = 0
        foreach ($index in $indexes) {
            $currentIndex++
            Write-Host "`nExporting index ${index} ($currentIndex of $($indexes.Count))..." -ForegroundColor Yellow
            
            $indexInfo = dism /Get-WimInfo /WimFile:$Global:InstallWimPath /Index:$index 2>$null
            $indexName = "Index ${index}"
            if ($indexInfo) {
                $nameMatch = $indexInfo | Select-String "Name\s*:\s*(.+)"
                if ($nameMatch) { $indexName = $nameMatch.Matches[0].Groups[1].Value.Trim() }
            }
            
            Write-Host "Processing: ${indexName}" -ForegroundColor Gray
            
            dism /Export-Image /SourceImageFile:$Global:InstallWimPath /SourceIndex:$index /DestinationImageFile:$targetPath /Compress:$compressionType /CheckIntegrity
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully exported: ${indexName}" -ForegroundColor Green
            } else {
                throw "Failed to export index ${index}"
            }
        }
        
        Write-Host "`nVerifying recompressed ESD file..." -ForegroundColor Yellow
        if (Test-ImageFile -Path $targetPath -Type "Recompressed ESD") {
            Write-Host "✓ ESD file verification successful" -ForegroundColor Green
            Write-Host "`nSUCCESS: ESD recompression completed!" -ForegroundColor Green
            Write-Host "Your ESD file has been recompressed with the new compression settings" -ForegroundColor Cyan
            
            if ($shouldBackup) {
                Write-Host "Original file has been safely backed up to: ${backupPath}" -ForegroundColor Cyan
            }
            
            $backupText = if ($shouldBackup) { "(backup will remain) " } else { "(no backup exists) " }
            $removeOriginal = Read-Host "`nReplace original install.esd with recompressed version? ${backupText}(Y/n)"
            if ($removeOriginal.ToLower() -ne 'n') {
                Remove-Item $Global:InstallWimPath -Force
                Move-Item $targetPath $Global:InstallWimPath -Force
                Write-Host "✓ Original install.esd replaced with recompressed version" -ForegroundColor Green
            } else {
                Write-Host "Recompressed file saved as: ${targetPath}" -ForegroundColor Cyan
                Write-Host "Original file remains unchanged" -ForegroundColor Green
                Write-Host "You can manually replace it later if desired." -ForegroundColor Gray
            }
        } else {
            throw "ESD file verification failed"
        }
        
    } catch {
        Write-Host "`nERROR: Recompression failed - $($_.Exception.Message)" -ForegroundColor Red
        if (Test-Path $targetPath) {
            Write-Host "Cleaning up incomplete ESD file..." -ForegroundColor Yellow
            try { Remove-Item $targetPath -Force } catch { 
                Write-Host "Warning: Could not remove incomplete ESD file: ${targetPath}" -ForegroundColor Yellow 
            }
        }
        Write-Host "Original ESD file remains unchanged." -ForegroundColor Green
    }
    
    Read-Host "Press Enter to continue"
}

function Manage-Editions {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "       Edition Management" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "CURRENT CONFIGURATION:" -ForegroundColor Magenta
        Write-Host "Install Image Path: " -NoNewline
        Write-Host $Global:InstallWimPath -ForegroundColor Green
        
        $fileType = if ($Global:InstallWimPath -like "*.esd") { "ESD" } else { "WIM" }
        Write-Host "File Type: " -NoNewline
        Write-Host $fileType -ForegroundColor Green
        
        Write-Host "Mount Path: " -NoNewline
        Write-Host "$Global:MountPath (Ready)" -ForegroundColor Green
        
        Write-Host "Wimlib Status: " -NoNewline
        $wimlibAvailable = $Global:WimlibPath -and (Test-Path (Join-Path $Global:WimlibPath "wimlib-imagex.exe"))
        if ($wimlibAvailable) {
            Write-Host "Available" -ForegroundColor Green
        } else {
            Write-Host "Not configured (Setup in Configuration menu)" -ForegroundColor Red
        }
        
        # Get current index count
        try {
            $indexes = Get-WimIndexes
            Write-Host "Available Indexes: " -NoNewline
            if ($indexes.Count -gt 0) {
                Write-Host "$($indexes.Count) found" -ForegroundColor Green
            } else {
                Write-Host "None found" -ForegroundColor Red
            }
        } catch {
            Write-Host "Available Indexes: " -NoNewline
            Write-Host "Unable to read" -ForegroundColor Red
        }
        Write-Host ""
        
        Write-Host "EDITION UPGRADE IMPORTANT NOTES:" -ForegroundColor Red
        Write-Host "• This is for OFFLINE Windows images only (WIM/ESD files)" -ForegroundColor Gray
        Write-Host "• Don't use this on an image that has already been upgraded to a higher edition" -ForegroundColor Gray
        Write-Host "• Recommended to use on the LOWEST edition available in the edition family" -ForegroundColor Gray
        Write-Host "• Edition upgrades change the image permanently" -ForegroundColor Gray
        Write-Host "• Product key can only be set AFTER the edition upgrade completes" -ForegroundColor Gray
        Write-Host ""
        
        Write-Host "EDITION OPERATIONS:" -ForegroundColor Cyan
        Write-Host "1. View Current Edition (Mounted Image)" -ForegroundColor Yellow
        Write-Host "2. View Available Target Editions (Mounted Image)" -ForegroundColor Yellow
        Write-Host "3. Upgrade Edition (Mounted Image)" -ForegroundColor Yellow
        Write-Host "4. View Index Metadata (requires wimlib)" -ForegroundColor $(if ($wimlibAvailable) { "Cyan" } else { "DarkGray" })
        Write-Host "5. Update Index Metadata (requires wimlib)" -ForegroundColor $(if ($wimlibAvailable) { "Cyan" } else { "DarkGray" })
        Write-Host "6. Set Product Keys (Generic/RTM)" -ForegroundColor Yellow
        Write-Host "7. Return to Main Menu" -ForegroundColor Red
        Write-Host ""
        Write-Host "Select an option (1-7): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { Get-CurrentEdition }
            "2" { Get-TargetEditions }
            "3" { Set-Edition }
            "4" { 
                if ($wimlibAvailable) {
                    View-IndexMetadata
                } else {
                    Write-Host "Wimlib is not configured." -ForegroundColor Red
                    Write-Host "Please set up wimlib in the Configuration menu first." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue"
                }
            }
            "5" { 
                if ($wimlibAvailable) {
                    Update-IndexMetadataWithWimlib
                } else {
                    Write-Host "Wimlib is not configured." -ForegroundColor Red
                    Write-Host "Please set up wimlib in the Configuration menu first." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue"
                }
            }
            "6" { Manage-ProductKeys }
            "7" { return }
            default {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

function Get-CurrentEdition {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    } catch {
        Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexNumber = Read-Host "`nEnter the index number to check current edition"
    
    $Global:CurrentOperation = "Getting Current Edition"
    $Global:CleanupRequired = $true
    
    try {
        Write-Host "`nMounting WIM/ESD index ${indexNumber}..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$indexNumber /MountDir:$Global:MountPath /CheckIntegrity /ReadOnly
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${indexNumber}"
        }
        
        Write-Host "`nCurrent Edition Information:" -ForegroundColor Cyan
        Write-Host "============================" -ForegroundColor Cyan
        $Global:CurrentOperation = "Getting Current Edition"
        dism /Image:$Global:MountPath /Get-CurrentEdition
        
        Write-Host "`nUnmounting image..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Unmounting"
        dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Image unmounted successfully" -ForegroundColor Green
        } else {
            Write-Host "✗ Warning: Image unmount may have had issues" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "ERROR: Failed to get current edition - $($_.Exception.Message)" -ForegroundColor Red
        
        Write-Host "Attempting to unmount image..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
    
    Read-Host "Press Enter to continue"
}

function Get-TargetEditions {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    } catch {
        Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexNumber = Read-Host "`nEnter the index number to check available target editions"
    
    $Global:CurrentOperation = "Getting Target Editions"
    $Global:CleanupRequired = $true
    
    try {
        Write-Host "`nMounting WIM/ESD index ${indexNumber}..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$indexNumber /MountDir:$Global:MountPath /CheckIntegrity /ReadOnly
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${indexNumber}"
        }
        
        Write-Host "`nCurrent Edition:" -ForegroundColor Cyan
        Write-Host "=================" -ForegroundColor Cyan
        dism /Image:$Global:MountPath /Get-CurrentEdition
        
        Write-Host "`nAvailable Target Editions:" -ForegroundColor Cyan
        Write-Host "===========================" -ForegroundColor Cyan
        $Global:CurrentOperation = "Getting Target Editions"
        dism /Image:$Global:MountPath /Get-TargetEditions
        
        Write-Host "`nUnmounting image..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Unmounting"
        dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Image unmounted successfully" -ForegroundColor Green
        } else {
            Write-Host "✗ Warning: Image unmount may have had issues" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "ERROR: Failed to get target editions - $($_.Exception.Message)" -ForegroundColor Red
        
        Write-Host "Attempting to unmount image..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
    
    Read-Host "Press Enter to continue"
}

function Set-Edition {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nEdition Upgrade Process:" -ForegroundColor Magenta
    Write-Host "1. Mount the image" -ForegroundColor Cyan
    Write-Host "2. Upgrade edition with product key in single command" -ForegroundColor Cyan
    Write-Host "3. Update metadata (name, description, flags)" -ForegroundColor Cyan
    Write-Host "4. Commit changes" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    } catch {
        Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexNumber = Read-Host "`nEnter the index number to upgrade"
    
    $Global:CurrentOperation = "Setting Edition (Offline)"
    $Global:CleanupRequired = $true
    
    try {
        Write-Host "`nMounting WIM/ESD index ${indexNumber}..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$indexNumber /MountDir:$Global:MountPath /CheckIntegrity
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${indexNumber}"
        }
        
        Write-Host "`nCurrent Edition:" -ForegroundColor Cyan
        Write-Host "=================" -ForegroundColor Cyan
        dism /Image:$Global:MountPath /Get-CurrentEdition
        
        Write-Host "`nAvailable Target Editions:" -ForegroundColor Cyan
        Write-Host "===========================" -ForegroundColor Cyan
        dism /Image:$Global:MountPath /Get-TargetEditions
        
        # STEP 1: Get edition and product key details
        Write-Host "`nSTEP 1: Edition Upgrade Setup" -ForegroundColor Yellow
        Write-Host "==============================" -ForegroundColor Yellow
        Write-Host "We'll upgrade the edition and set the product key in a single command for better compatibility." -ForegroundColor Gray
        Write-Host ""
        
        $targetEdition = Read-Host "Enter the target edition ID (e.g., Professional, Enterprise)"
        if (-not $targetEdition) {
            Write-Host "No target edition specified. Cancelling operation." -ForegroundColor Yellow
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
            return
        }
        
        Write-Host ""
        Write-Host "Product Key Setup:" -ForegroundColor Cyan
        Write-Host "You can type the product key in lowercase - it will be automatically converted to uppercase." -ForegroundColor Gray
        
        # Check for recommended RTM key based on target edition
        $recommendedKeyName = Get-KeyNameFromTargetEdition -TargetEdition $targetEdition
        $productKey = ""
        
        if ($recommendedKeyName) {
            # Search for exact matches in both RTM and KMS collections
            $rtmMatch = Find-ExactMatchOnly -SearchTerm $recommendedKeyName -KeyCollection $Global:RTMKeys -KeyType "RTM"
            $kmsMatch = Find-ExactMatchOnly -SearchTerm $recommendedKeyName -KeyCollection $Global:KMSClientKeys -KeyType "KMS Client"
            
            $recommendations = @()
            if ($rtmMatch) { $recommendations += $rtmMatch }
            if ($kmsMatch) { $recommendations += $kmsMatch }
            
            if ($recommendations.Count -gt 0) {
                Write-Host ""
                Write-Host "Recommended Product Keys for $targetEdition :" -ForegroundColor Green
                Write-Host "===============================================" -ForegroundColor Green
                $i = 1
                foreach ($rec in $recommendations) {
                    Write-Host "$i. $($rec.Type): $($rec.EditionName)" -ForegroundColor Cyan
                    Write-Host "   Key: $($rec.Key)" -ForegroundColor Green
                    Write-Host "   Reason: $($rec.Reason)" -ForegroundColor Gray
                    Write-Host ""
                    $i++
                }
                
                Write-Host "Options:" -ForegroundColor Yellow
                Write-Host "1-$($recommendations.Count). Use recommended key" -ForegroundColor Green
                Write-Host "$($recommendations.Count + 1). Enter custom key" -ForegroundColor Yellow
                Write-Host "$($recommendations.Count + 2). Continue without key" -ForegroundColor Red
                
                do {
                    $choice = Read-Host "Select option"
                    $choiceNum = 0
                    if ([int]::TryParse($choice, [ref]$choiceNum)) {
                        if ($choiceNum -ge 1 -and $choiceNum -le $recommendations.Count) {
                            $productKey = $recommendations[$choiceNum - 1].Key
                            Write-Host "Selected: $($recommendations[$choiceNum - 1].EditionName) ($($recommendations[$choiceNum - 1].Type))" -ForegroundColor Green
                            break
                        } elseif ($choiceNum -eq ($recommendations.Count + 1)) {
                            $productKeyInput = Read-Host "Enter the product key for the new edition (format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)"
                            if ($productKeyInput -and $productKeyInput.Length -ge 25) {
                                $productKey = $productKeyInput.ToUpper()
                                Write-Host "Product key to be used: $productKey" -ForegroundColor Green
                                
                                $confirmKey = Read-Host "Is this product key correct? (Y/n)"
                                if ($confirmKey.ToLower() -eq 'n') {
                                    Write-Host "Operation cancelled due to product key confirmation." -ForegroundColor Yellow
                                    dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                                    return
                                }
                            }
                            break
                        } elseif ($choiceNum -eq ($recommendations.Count + 2)) {
                            $productKey = ""
                            break
                        }
                    }
                    Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                } while ($true)
            } else {
                # No recommendations found, fall back to manual entry
                $productKeyInput = Read-Host "Enter the product key for the new edition (format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)"
                if ($productKeyInput -and $productKeyInput.Length -ge 25) {
                    $productKey = $productKeyInput.ToUpper()
                    Write-Host "Product key to be used: $productKey" -ForegroundColor Green
                    
                    $confirmKey = Read-Host "Is this product key correct? (Y/n)"
                    if ($confirmKey.ToLower() -eq 'n') {
                        Write-Host "Operation cancelled due to product key confirmation." -ForegroundColor Yellow
                        dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                        return
                    }
                } else {
                    Write-Host "Invalid or missing product key. Edition upgrade requires a valid product key." -ForegroundColor Red
                    $proceed = Read-Host "Continue without product key? (y/N)"
                    if ($proceed.ToLower() -ne 'y') {
                        dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                        return
                    }
                }
            }
        } else {
            # No mapping available for this target edition, fall back to manual entry
            Write-Host "No recommended keys available for edition: $targetEdition" -ForegroundColor Yellow
            $productKeyInput = Read-Host "Enter the product key for the new edition (format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)"
            if ($productKeyInput -and $productKeyInput.Length -ge 25) {
                $productKey = $productKeyInput.ToUpper()
                Write-Host "Product key to be used: $productKey" -ForegroundColor Green
                
                $confirmKey = Read-Host "Is this product key correct? (Y/n)"
                if ($confirmKey.ToLower() -eq 'n') {
                    Write-Host "Operation cancelled due to product key confirmation." -ForegroundColor Yellow
                    dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                    return
                }
            } else {
                Write-Host "Invalid or missing product key. Edition upgrade requires a valid product key." -ForegroundColor Red
                $proceed = Read-Host "Continue without product key? (y/N)"
                if ($proceed.ToLower() -ne 'y') {
                    dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                    return
                }
            }
        }
        
        # STEP 2: Unified edition upgrade command
        Write-Host "`nSTEP 2: Edition Upgrade" -ForegroundColor Yellow
        Write-Host "========================" -ForegroundColor Yellow
        Write-Host "Upgrading to edition: $targetEdition" -ForegroundColor Cyan
        
        if ($productKey) {
            Write-Host "Using: dism /Image:$Global:MountPath /Set-Edition:$targetEdition /AcceptEula /ProductKey:$productKey" -ForegroundColor Gray
            Write-Host "This operation may take several minutes..." -ForegroundColor Cyan
            
            $Global:CurrentOperation = "Setting Edition with Product Key"
            dism /Image:$Global:MountPath /Set-Edition:$targetEdition /AcceptEula /ProductKey:$productKey
        } else {
            Write-Host "Using: dism /Image:$Global:MountPath /Set-Edition:$targetEdition /AcceptEula" -ForegroundColor Gray
            Write-Host "This operation may take several minutes..." -ForegroundColor Cyan
            
            $Global:CurrentOperation = "Setting Edition without Product Key"
            dism /Image:$Global:MountPath /Set-Edition:$targetEdition /AcceptEula
        }
        
        $editionUpgradeSuccessful = $false
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Successfully upgraded to edition: $targetEdition" -ForegroundColor Green
            if ($productKey) {
                Write-Host "✓ Product key applied successfully" -ForegroundColor Green
            }
            $editionUpgradeSuccessful = $true
        } else {
            Write-Host "✗ Failed to upgrade to edition: $targetEdition" -ForegroundColor Red
            if ($productKey) {
                Write-Host "The unified command failed. This could be due to edition/key compatibility." -ForegroundColor Yellow
            }
            Write-Host "Discarding changes..." -ForegroundColor Yellow
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
            return
        }
        
        # STEP 3: Commit changes and update metadata
        if ($editionUpgradeSuccessful) {
            Write-Host "`nSTEP 3: Committing Changes" -ForegroundColor Yellow
            Write-Host "===========================" -ForegroundColor Yellow
            Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
            $Global:CurrentOperation = "Unmounting"
            dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Changes committed successfully" -ForegroundColor Green
                Write-Host "`n✓ Edition upgrade completed successfully!" -ForegroundColor Green
                
                # STEP 4: Optional metadata update with wimlib
                $wimlibAvailable = Get-WimlibImagexPath
                if ($wimlibAvailable) {
                    Write-Host "`nSTEP 4: Update Index Metadata (Optional)" -ForegroundColor Yellow
                    Write-Host "=========================================" -ForegroundColor Yellow
                    Write-Host "Would you like to update the index name and metadata to reflect the new edition?" -ForegroundColor Cyan
                    Write-Host "This will update the name, description, and FLAGS to match the upgraded edition." -ForegroundColor Gray
                    
                    $updateMetadata = Read-Host "Update index metadata? (Y/n)"
                    if ($updateMetadata.ToLower() -ne 'n') {
                        
                        Write-Host "`nEnter new metadata for the upgraded edition:" -ForegroundColor Cyan
                        Write-Host "Leave fields blank to keep current values" -ForegroundColor Gray
                        Write-Host ""
                        
                        $suggestedName = if ($targetEdition -eq "Professional") { 
                            "Windows 11 Pro" 
                        } elseif ($targetEdition -eq "Enterprise") { 
                            "Windows 11 Enterprise" 
                        } elseif ($targetEdition -eq "Education") { 
                            "Windows 11 Education" 
                        } else { 
                            "Windows 11 $targetEdition" 
                        }
                        
                        Write-Host "Suggested name: $suggestedName" -ForegroundColor Green
                        
                        $newName = Read-Host "New image name"
                        $newDescription = Read-Host "New image description"
                        
                        Write-Host ""
                        Write-Host "Note: FLAGS corresponds to the Windows edition (e.g., HomeN, Professional, ProfessionalWorkstation)" -ForegroundColor Cyan
                        Write-Host "Suggested FLAGS: $targetEdition" -ForegroundColor Green
                        $newFlags = Read-Host "New image FLAGS/edition"
                        
                        if ($newName -or $newDescription -or $newFlags) {
                            Write-Host "`nUpdating image metadata with wimlib..." -ForegroundColor Yellow
                            
                            try {
                                $args = @("info", $Global:InstallWimPath, $indexNumber)
                                
                                if ($newName) {
                                    $args += $newName
                                    if ($newDescription) {
                                        $args += $newDescription
                                    }
                                }
                                
                                if ($newFlags) {
                                    $args += "--image-property"
                                    $args += "FLAGS=$newFlags"
                                    $args += "--image-property" 
                                    $args += "EDITIONID=$newFlags"
                                }
                                
                                if ($newName) {
                                    $args += "--image-property"
                                    $args += "DISPLAYNAME=$newName"
                                }
                                
                                if ($newDescription) {
                                    $args += "--image-property"
                                    $args += "DISPLAYDESCRIPTION=$newDescription"
                                }
                                
                                $wimlibPath = Get-WimlibImagexPath
                                Write-Host "Command: wimlib-imagex $($args -join ' ')" -ForegroundColor Gray
                                & $wimlibPath @args
                                
                                if ($LASTEXITCODE -eq 0) {
                                    Write-Host "✓ Index metadata updated successfully!" -ForegroundColor Green
                                    
                                    $updates = @()
                                    if ($newName) { $updates += "Name: $newName"; $updates += "DisplayName: $newName" }
                                    if ($newDescription) { $updates += "Description: $newDescription"; $updates += "DisplayDescription: $newDescription" }
                                    if ($newFlags) { $updates += "Flags: $newFlags" }
                                    
                                    Write-Host "Updated properties:" -ForegroundColor Cyan
                                    foreach ($update in $updates) {
                                        Write-Host "  ✓ $update" -ForegroundColor Green
                                    }
                                } else {
                                    Write-Host "✗ Failed to update metadata" -ForegroundColor Red
                                }
                                
                            } catch {
                                Write-Host "✗ Error updating metadata: $($_.Exception.Message)" -ForegroundColor Red
                            }
                        } else {
                            Write-Host "No metadata changes specified." -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Host "`nTo update the index name to match the new edition:" -ForegroundColor Cyan
                    Write-Host "• Use the 'Update Index Metadata' option in the Edition Management menu" -ForegroundColor Gray
                    Write-Host "• This requires wimlib to be configured in the Configuration menu" -ForegroundColor Gray
                }
                
                Write-Host ""
                Write-Host "Edition upgrade process completed!" -ForegroundColor Green
            } else {
                throw "Failed to commit edition upgrade changes"
            }
        }
        
    } catch {
        Write-Host "ERROR: Failed during edition upgrade process - $($_.Exception.Message)" -ForegroundColor Red
        
        Write-Host "Attempting to discard changes and unmount..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
    
    Read-Host "Press Enter to continue"
}

function Get-WimlibImagexPath {
    if ($Global:WimlibPath) {
        $wimlibExe = Join-Path $Global:WimlibPath "wimlib-imagex.exe"
        if (Test-Path $wimlibExe) {
            return $wimlibExe
        } else {
            $Global:WimlibPath = $null
        }
    }
    
    try {
        $result = Get-Command "wimlib-imagex.exe" -ErrorAction SilentlyContinue
        if ($result) {
            $Global:WimlibPath = Split-Path $result.Source -Parent
            return $result.Source
        }
    } catch {}
    
    return $null
}

function Update-IndexMetadataWithWimlib {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    # Check if wimlib-imagex is available
    $wimlibPath = Get-WimlibImagexPath
    if (-not $wimlibPath) {
        Write-Host "Wimlib is not configured." -ForegroundColor Red
        Write-Host "Please set up wimlib in the Configuration menu first." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    $fileType = if ($Global:InstallWimPath -like "*.esd") { "ESD" } else { "WIM" }
    
    Write-Host "`nUpdate Index Metadata" -ForegroundColor Cyan
    Write-Host "=====================" -ForegroundColor Cyan
    Write-Host "Using wimlib for fast metadata updates" -ForegroundColor Gray
    Write-Host "File: $Global:InstallWimPath ($fileType format)" -ForegroundColor Green
    Write-Host ""
    
    if ($fileType -eq "ESD") {
        Write-Host "Note: ESD files are supported, but encrypted ESD files are not." -ForegroundColor Yellow
        Write-Host "If this fails, the file may be encrypted." -ForegroundColor Yellow
        Write-Host ""
    }
    
    # Show current indexes using wimlib-imagex info
    Write-Host "Available indexes:" -ForegroundColor Yellow
    try {
        & $wimlibPath info $Global:InstallWimPath
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to read image information"
        }
    } catch {
        Write-Host "Error getting image information with wimlib-imagex." -ForegroundColor Red
        if ($fileType -eq "ESD") {
            Write-Host "This ESD file may be encrypted or corrupted." -ForegroundColor Yellow
        }
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexNumber = Read-Host "`nEnter index number to update"
    
    # Validate index exists and show current metadata
    Write-Host "`nCurrent metadata for index ${indexNumber}:" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    try {
        & $wimlibPath info $Global:InstallWimPath $indexNumber
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Invalid index number: ${indexNumber}" -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    } catch {
        Write-Host "Error getting current metadata for index ${indexNumber}" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nEnter new metadata:" -ForegroundColor Yellow
    Write-Host "===================" -ForegroundColor Yellow
    Write-Host "Leave fields blank to keep current values" -ForegroundColor Gray
    Write-Host ""
    
    $newName = Read-Host "New image name"
    $newDescription = Read-Host "New image description"
    
    Write-Host ""
    Write-Host "Note: FLAGS corresponds to the Windows edition (e.g., HomeN, ProfessionalWorkstation)" -ForegroundColor Cyan
    $newFlags = Read-Host "New image FLAGS/edition"
    
    if (-not $newName -and -not $newDescription -and -not $newFlags) {
        Write-Host "No changes specified. Operation cancelled." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nUpdating image metadata..." -ForegroundColor Yellow
    
    try {
        # Build single command: wimlib-imagex info WIMFILE IMAGE [NEW_NAME [NEW_DESC]] [OPTIONS]
        $args = @("info", $Global:InstallWimPath, $indexNumber)
        
        # Add name if provided
        if ($newName) {
            $args += $newName
        }
        
        # Add description if provided (name must be provided first)
        if ($newDescription) {
            if (-not $newName) {
                Write-Host "Error: Cannot set description without also setting name." -ForegroundColor Red
                Write-Host "Please provide both name and description." -ForegroundColor Yellow
                Read-Host "Press Enter to continue"
                return
            }
            $args += $newDescription
        }
        
        # Add FLAGS property if provided
        if ($newFlags) {
            $args += "--image-property"
            $args += "FLAGS=$newFlags"
        }
        
        # Add DISPLAYNAME property if name was provided
        if ($newName) {
            $args += "--image-property"
            $args += "DISPLAYNAME=$newName"
        }
        
        # Add DISPLAYDESCRIPTION property if description was provided
        if ($newDescription) {
            $args += "--image-property"
            $args += "DISPLAYDESCRIPTION=$newDescription"
        }
        
        Write-Host "Command: wimlib-imagex $($args -join ' ')" -ForegroundColor Gray
        & $wimlibPath @args
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "`n✓ Image metadata updated successfully!" -ForegroundColor Green
            
            # Show what was updated
            $updates = @()
            if ($newName) { $updates += "Name: $newName"; $updates += "DisplayName: $newName" }
            if ($newDescription) { $updates += "Description: $newDescription"; $updates += "DisplayDescription: $newDescription" }
            if ($newFlags) { $updates += "Flags: $newFlags" }
            
            Write-Host "Updated properties:" -ForegroundColor Cyan
            foreach ($update in $updates) {
                Write-Host "  ✓ $update" -ForegroundColor Green
            }
            
            Write-Host ""
            Write-Host "Updated metadata:" -ForegroundColor Cyan
            Write-Host "=================" -ForegroundColor Cyan
            & $wimlibPath info $Global:InstallWimPath $indexNumber
        } else {
            Write-Host "`n✗ Failed to update metadata" -ForegroundColor Red
        }
        
    } catch {
        Write-Host "`n✗ Error updating metadata: $($_.Exception.Message)" -ForegroundColor Red
        if ($fileType -eq "ESD") {
            Write-Host "This ESD file may be encrypted or have other issues." -ForegroundColor Yellow
        }
    }
    
    Read-Host "Press Enter to continue"
}

function View-IndexMetadata {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    # Check if wimlib-imagex is available
    $wimlibPath = Get-WimlibImagexPath
    if (-not $wimlibPath) {
        Write-Host "Wimlib is not configured." -ForegroundColor Red
        Write-Host "Please set up wimlib in the Configuration menu first." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    $fileType = if ($Global:InstallWimPath -like "*.esd") { "ESD" } else { "WIM" }
    
    Write-Host "`nView Index Metadata" -ForegroundColor Cyan
    Write-Host "===================" -ForegroundColor Cyan
    Write-Host "File: $Global:InstallWimPath ($fileType format)" -ForegroundColor Green
    Write-Host ""
    
    # Show all indexes with detailed information using wimlib-imagex info
    Write-Host "All indexes with detailed metadata:" -ForegroundColor Yellow
    try {
        & $wimlibPath info $Global:InstallWimPath
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to read image information"
        }
    } catch {
        Write-Host "Error getting image information with wimlib-imagex." -ForegroundColor Red
        if ($fileType -eq "ESD") {
            Write-Host "This ESD file may be encrypted or corrupted." -ForegroundColor Yellow
        }
    }
    
    Read-Host "Press Enter to continue"
}

function Get-BackupPath {
    param([string]$SourceFormat, [string]$OperationType)
    
    Write-Host "`nBackup Options:" -ForegroundColor Magenta
    Write-Host "The $($OperationType.ToLower()) process is safe and doesn't modify the original $($SourceFormat.ToUpper()) file." -ForegroundColor Gray
    Write-Host "However, you can create a backup for extra safety." -ForegroundColor Gray
    
    $createBackup = Read-Host "`nCreate backup of original $($SourceFormat.ToUpper()) before $($OperationType.ToLower())? (y/N)"
    if ($createBackup.ToLower() -ne 'y') { return $null }
    
    $defaultBackupPath = if ($Global:BackupPath) {
        $fileName = Split-Path $Global:InstallWimPath -Leaf
        Join-Path $Global:BackupPath ($fileName + ".backup")
    } else {
        $Global:InstallWimPath -replace "\.$SourceFormat`$", ".$SourceFormat.backup"
    }
    
    Write-Host "`nDefault backup path: ${defaultBackupPath}" -ForegroundColor Green
    
    do {
        $customPath = Read-Host "Enter new path for backup (Press Enter for default)"
        
        if (-not $customPath) {
            if ($Global:BackupPath) {
                $backupDir = Split-Path $defaultBackupPath -Parent
                if (-not (Test-Path $backupDir)) {
                    try {
                        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
                        Write-Host "Created backup directory: $backupDir" -ForegroundColor Green
                    } catch {
                        Write-Host "Warning: Could not create backup directory: $backupDir" -ForegroundColor Yellow
                        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
                        continue
                    }
                }
            }
            return $defaultBackupPath
        } else {
            $customPath = $customPath.Trim('"', "'")
            $customPath = [System.Environment]::ExpandEnvironmentVariables($customPath)
            
            if ($customPath -match '\.(wim|esd|backup)$') {
                $directory = Split-Path $customPath -Parent
                if ($directory -and (Test-Path $directory)) {
                    return $customPath
                } elseif (-not $directory) {
                    return Join-Path (Get-Location) $customPath
                } else {
                    Write-Host "Directory does not exist: ${directory}" -ForegroundColor Red
                }
            } else {
                if (Test-Path $customPath) {
                    $fileName = Split-Path $Global:InstallWimPath -Leaf
                    return Join-Path $customPath ($fileName + ".backup")
                } else {
                    Write-Host "Directory does not exist: ${customPath}" -ForegroundColor Red
                }
            }
        }
    } while ($true)
}

function Get-CompressionChoice {
    param([string]$TargetFormat)
    
    Write-Host "`nCompression Options ($($TargetFormat.ToUpper()) Format):" -ForegroundColor Magenta
    Write-Host "1. Maximum (Normal compression, Balanced conversion)" -ForegroundColor Yellow
    Write-Host "2. Fast (Quick compression, Faster conversion)" -ForegroundColor Yellow
    Write-Host "3. None (No compression, Fastest conversion)" -ForegroundColor Yellow
    Write-Host "4. Recovery (Best compression, Slowest conversion)" -ForegroundColor Green
    
    do {
        $choice = Read-Host "`nSelect compression type (1-4, default is 4)"
        if (-not $choice) { $choice = "4" }
        
        switch ($choice) {
            "1" { return "maximum" }
            "2" { return "fast" }
            "3" { return "none" }
            "4" { return "recovery" }
            default { Write-Host "Invalid selection. Please choose 1-4." -ForegroundColor Red }
        }
    } while ($true)
}

function Show-ConversionDetails {
    param([string]$SourcePath, [string]$TargetPath, [string]$CompressionType, [string]$BackupPath)
    
    Write-Host "`nConversion Details:" -ForegroundColor Cyan
    Write-Host "Source: $SourcePath" -ForegroundColor Gray
    Write-Host "Target: $TargetPath" -ForegroundColor Gray
    Write-Host "Compression: $CompressionType" -ForegroundColor Gray
    Write-Host $(if ($BackupPath) { "Backup will be created: $BackupPath" } else { "No backup will be created" }) -ForegroundColor Gray
}

function Invoke-ImageConversion {
    param([string]$SourcePath, [string]$TargetPath, [string]$CompressionType)
    
    $indexes = Get-WimIndexes
    if ($indexes.Count -eq 0) {
        Write-Host "No indexes found in source file." -ForegroundColor Red
        return $false
    }
    
    Write-Host "`nStarting conversion process..." -ForegroundColor Yellow
    Write-Host "This may take significant time depending on indexes and compression." -ForegroundColor Cyan
    
    $currentIndex = 0
    foreach ($index in $indexes) {
        $currentIndex++
        Write-Host "`nExporting index ${index} ($currentIndex of $($indexes.Count))..." -ForegroundColor Yellow
        
        try {
            $indexInfo = dism /Get-WimInfo /WimFile:$SourcePath /Index:$index 2>$null
            $indexName = "Index ${index}"
            if ($indexInfo) {
                $nameMatch = $indexInfo | Select-String "Name\s*:\s*(.+)"
                if ($nameMatch) { $indexName = $nameMatch.Matches[0].Groups[1].Value.Trim() }
            }
            
            Write-Host "Processing: ${indexName}" -ForegroundColor Gray
            
            dism /Export-Image /SourceImageFile:$SourcePath /SourceIndex:$index /DestinationImageFile:$TargetPath /Compress:$CompressionType /CheckIntegrity
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully exported: ${indexName}" -ForegroundColor Green
            } else {
                throw "Failed to export index ${index}"
            }
            
        } catch {
            Write-Host "✗ Error exporting index ${index} : $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
    
    return $true
}

function Handle-OriginalFileCleanup {
    param([string]$SourceFormat, [string]$TargetFormat, [string]$BackupPath, [bool]$IsRecompression)
    
    if ($BackupPath) {
        Write-Host "Original $($SourceFormat.ToUpper()) has been safely backed up to: $BackupPath" -ForegroundColor Cyan
    }
    
    if (-not $IsRecompression) {
        $backupText = if ($BackupPath) { "(backup will remain) " } else { "(no backup exists) " }
        $removeOriginal = Read-Host "`nRemove original install.$SourceFormat? ${backupText}(y/N)"
        
        if ($removeOriginal.ToLower() -eq 'y') {
            $originalToRemove = $Global:InstallWimPath -replace "\.$TargetFormat`$", ".$SourceFormat"
            if (Test-Path $originalToRemove) {
                Remove-Item $originalToRemove -Force
                Write-Host "✓ Original install.$SourceFormat removed" -ForegroundColor Green
            }
        }
    }
}

function Manage-Drivers {
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "       Driver Management" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        $requiredPathsSet = Test-RequiredPaths
        $mountPathEmpty = Test-MountPathEmpty
        
        Write-Host "CURRENT CONFIGURATION:" -ForegroundColor Magenta
        Write-Host "Install Image Path: " -NoNewline
        if ($Global:InstallWimPath) {
            Write-Host $Global:InstallWimPath -ForegroundColor Green
        } else {
            Write-Host "Not Set" -ForegroundColor Red
        }
        Write-Host "Mount Path: " -NoNewline
        if ($Global:MountPath) {
            if ($mountPathEmpty) {
                Write-Host "$Global:MountPath (Ready)" -ForegroundColor Green
            } else {
                Write-Host "$Global:MountPath (IN USE - May need cleanup)" -ForegroundColor Red
            }
        } else {
            Write-Host "Not Set" -ForegroundColor Red
        }
        Write-Host "Driver Path: " -NoNewline
        if ($Global:DriverPath) {
            Write-Host $Global:DriverPath -ForegroundColor Green
        } else {
            Write-Host "Not Set (Will prompt when needed)" -ForegroundColor Yellow
        }
        Write-Host "Force Unsigned Drivers: " -NoNewline
        if ($Global:ForceUnsignedDrivers) {
            Write-Host "Enabled" -ForegroundColor Green
        } else {
            Write-Host "Disabled" -ForegroundColor Red
        }
        Write-Host ""
        
        Write-Host "DRIVER OPERATIONS:" -ForegroundColor Cyan
        Write-Host "1. Add Drivers" -ForegroundColor $(if ($requiredPathsSet -and $mountPathEmpty) { "Yellow" } else { "DarkGray" })
        Write-Host "2. Remove Drivers" -ForegroundColor $(if ($requiredPathsSet -and $mountPathEmpty) { "Yellow" } else { "DarkGray" })
        Write-Host "3. Export Drivers from Current System" -ForegroundColor Yellow
        Write-Host "4. Export Drivers from Mounted Image" -ForegroundColor $(if ($requiredPathsSet -and $mountPathEmpty) { "Yellow" } else { "DarkGray" })
        Write-Host "5. Return to Main Menu" -ForegroundColor Red
        Write-Host ""
        Write-Host "Select an option (1-5): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { 
                if (-not $requiredPathsSet) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } elseif (-not $mountPathEmpty) {
                    Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Add-DriversToWim
                }
            }
            "2" { 
                if (-not $requiredPathsSet) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } elseif (-not $mountPathEmpty) {
                    Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Remove-DriversFromWim
                }
            }
            "3" { Export-DriversFromSystem }
            "4" { 
                if (-not $requiredPathsSet) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } elseif (-not $mountPathEmpty) {
                    Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Export-DriversFromMountedImage
                }
            }
            "5" { return }
            default {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

function Remove-DriversFromWim {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nDriver Removal Options:" -ForegroundColor Cyan
    Write-Host "1. Remove from specific index"
    Write-Host "2. Remove from all indexes"
    $choice = Read-Host "Select option (1-2)"
    $Global:RemovedDrivers = @()
    
    $Global:CurrentOperation = "Removing Drivers"
    $Global:CleanupRequired = $true
    
    try {
        if ($choice -eq "1") {
            Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
            try {
                dism /Get-WimInfo /WimFile:$Global:InstallWimPath
                if ($LASTEXITCODE -ne 0) {
                    Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                    return
                }
            } catch {
                Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
                Read-Host "Press Enter to continue"
                return
            }
            
            $indexNumber = Read-Host "`nEnter the index number"
            Remove-DriversFromIndex -IndexNumber $indexNumber -IsMultiIndex $false
        } elseif ($choice -eq "2") {
            $indexes = Get-WimIndexes
            Write-Host "`nSelect which Windows edition to use as base for driver selection:" -ForegroundColor Yellow
            Write-Host "(Different editions may have different installed drivers)" -ForegroundColor Gray
            try {
                dism /Get-WimInfo /WimFile:$Global:InstallWimPath
            } catch {
                Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
                Read-Host "Press Enter to continue"
                return
            }
            
            do {
                $baseIndex = Read-Host "`nEnter the index number of the edition to use as base"
                if ($indexes -contains $baseIndex) {
                    break
                } else {
                    Write-Host "Invalid index. Please select from the available indexes above." -ForegroundColor Red
                }
            } while ($true)
            
            Write-Host "`n--- Processing Base Index $baseIndex (Driver Selection) ---" -ForegroundColor Magenta
            Remove-DriversFromIndex -IndexNumber $baseIndex -IsMultiIndex $true -IsFirstIndex $true
            
            foreach ($index in $indexes) {
                if ($index -ne $baseIndex) {
                    Write-Host "`n--- Processing Index ${index} ---" -ForegroundColor Magenta
                    Remove-DriversFromIndex -IndexNumber $index -IsMultiIndex $true -IsFirstIndex $false
                }
            }
        } else {
            Write-Host "Invalid selection." -ForegroundColor Red
        }
    } catch {
        Write-Host "Error during driver removal: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
    
    Read-Host "Press Enter to continue"
}

function Remove-DriversFromIndex {
    param(
        [string]$IndexNumber,
        [bool]$IsMultiIndex,
        [bool]$IsFirstIndex = $true
    )
    
    try {
        Write-Host "`nMounting WIM/ESD index ${IndexNumber}..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$IndexNumber /MountDir:$Global:MountPath /CheckIntegrity
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${IndexNumber}"
        }
        
        if ($IsMultiIndex -and -not $IsFirstIndex) {
            if ($Global:RemovedDrivers.Count -gt 0) {
                Write-Host "`nApplying previously selected drivers to remove..." -ForegroundColor Yellow
                foreach ($driver in $Global:RemovedDrivers) {
                    Write-Host "Removing driver: $driver" -ForegroundColor Gray
                    try {
                        dism /Image:$Global:MountPath /Remove-Driver /Driver:$driver
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host "✓ Successfully removed driver: $driver" -ForegroundColor Green
                        } else {
                            Write-Host "✗ Failed to remove driver: $driver" -ForegroundColor Red
                        }
                    } catch {
                        Write-Host "✗ Error removing driver $driver : $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
        } else {
            do {
                Clear-Host
                Write-Host "Driver Removal - Index: ${IndexNumber}" -ForegroundColor Cyan
                Write-Host "====================================" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "1. View all installed drivers" -ForegroundColor Yellow
                Write-Host "2. Remove specific drivers" -ForegroundColor Yellow
                Write-Host "3. Commit changes and continue" -ForegroundColor Green
                Write-Host "4. Discard changes and exit" -ForegroundColor Red
                
                $managementChoice = Read-Host "`nSelect option (1-4)"
                
                switch ($managementChoice) {
                    "1" { 
                        Show-InstalledDrivers -ImagePath $Global:MountPath
                        Read-Host "`nPress Enter to continue"
                    }
                    "2" {
                        Remove-SpecificDrivers -ImagePath $Global:MountPath -IsMultiIndex $IsMultiIndex
                    }
                    "3" {
                        Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
                        $Global:CurrentOperation = "Unmounting"
                        dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
                        
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host "SUCCESS: Changes committed for index ${IndexNumber}" -ForegroundColor Green
                        } else {
                            throw "Failed to commit changes for index ${IndexNumber}"
                        }
                        return
                    }
                    "4" {
                        Write-Host "Discarding changes and unmounting..." -ForegroundColor Yellow
                        dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                        Write-Host "Changes discarded." -ForegroundColor Yellow
                        return
                    }
                    default {
                        Write-Host "Invalid selection." -ForegroundColor Red
                        Start-Sleep -Seconds 1
                    }
                }
            } while ($true)
        }
        
        if ($IsMultiIndex -and -not $IsFirstIndex) {
            Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
            $Global:CurrentOperation = "Unmounting"
            dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "SUCCESS: Changes committed for index ${IndexNumber}" -ForegroundColor Green
            } else {
                throw "Failed to commit changes for index ${IndexNumber}"
            }
        }
        
    } catch {
        Write-Host "ERROR processing index ${IndexNumber} : $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Attempting to discard changes and unmount..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    }
}

function Show-InstalledDrivers {
    param([string]$ImagePath)
    
    Write-Host "Getting installed drivers (this may take a moment)..." -ForegroundColor Yellow
    try {
        Write-Host "`nInstalled Third-Party Drivers:" -ForegroundColor Cyan
        Write-Host "===============================" -ForegroundColor Cyan
        dism /Image:$ImagePath /Get-Drivers /All /Format:Table
    } catch {
        Write-Host "Error getting drivers: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Remove-SpecificDrivers {
    param([string]$ImagePath, [bool]$IsMultiIndex)
    
    Write-Host "`nRemoving Specific Drivers" -ForegroundColor Cyan
    Write-Host "Getting installed third-party drivers for reference..." -ForegroundColor Yellow
    try {
        Write-Host "`nInstalled Third-Party Drivers:" -ForegroundColor Cyan
        Write-Host "===============================" -ForegroundColor Cyan
        dism /Image:$ImagePath /Get-Drivers /All /Format:Table
    } catch {
        Write-Host "Error getting drivers: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nEnter driver OEM names separated by commas (or press Enter to skip):"
    Write-Host "Example: oem1.inf,oem2.inf,oem5.inf"
    Write-Host "Note: Use the 'Published Name' from the driver list above"
    
    $driverInput = Read-Host "Driver OEM names to remove"
    if (-not $driverInput) {
        return
    }
    
    $drivers = $driverInput -split ',' | ForEach-Object { $_.Trim() }
    
    if ($IsMultiIndex) {
        $Global:RemovedDrivers += $drivers
        $applyToAll = Read-Host "`nApply these driver removals to all indexes? (Y/n)"
        if ($applyToAll.ToLower() -eq 'n') {
            $Global:RemovedDrivers = @()
        }
    }
    
    Write-Host "`nRemoving drivers..." -ForegroundColor Yellow
    foreach ($driver in $drivers) {
        Write-Host "Removing driver: $driver" -ForegroundColor Gray
        try {
            dism /Image:$ImagePath /Remove-Driver /Driver:$driver
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully removed driver: $driver" -ForegroundColor Green
            } else {
                Write-Host "✗ Failed to remove driver: $driver" -ForegroundColor Red
            }
        } catch {
            Write-Host "✗ Error removing driver $driver : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Read-Host "`nPress Enter to continue"
}

function Export-DriversFromSystem {
    Write-Host "`nExport Drivers from Current System" -ForegroundColor Cyan
    Write-Host "====================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This will export all third-party drivers from the current Windows installation." -ForegroundColor Gray
    Write-Host ""
    $exportPath = Get-DriverExportPath -ExportType "Current System"
    if (-not $exportPath) {
        return
    }
    
    Write-Host "`nExporting drivers from current system..." -ForegroundColor Yellow
    Write-Host "Export Path: $exportPath" -ForegroundColor Gray
    Write-Host "This may take several minutes depending on the number of drivers installed." -ForegroundColor Cyan
    
    try {
        if (-not (Test-Path $exportPath)) {
            New-Item -ItemType Directory -Path $exportPath -Force | Out-Null
            Write-Host "Created export directory: $exportPath" -ForegroundColor Green
        }
        
        dism /Online /Export-Driver /Destination:$exportPath
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "`nSUCCESS: Drivers exported to $exportPath" -ForegroundColor Green
            $driverFiles = Get-ChildItem -Path $exportPath -Filter "*.inf" -Recurse -ErrorAction SilentlyContinue
            if ($driverFiles) {
                Write-Host "Total driver packages exported: $($driverFiles.Count)" -ForegroundColor Cyan
            }
        } else {
            throw "Driver export failed with exit code $LASTEXITCODE"
        }
        
    } catch {
        Write-Host "ERROR: Failed to export drivers - $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Read-Host "Press Enter to continue"
}

function Export-DriversFromMountedImage {
    Write-Host "`nExport Drivers from Mounted Image" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This will export all third-party drivers from a mounted Windows image." -ForegroundColor Gray
    Write-Host ""
    Write-Host "Available indexes in WIM/ESD:" -ForegroundColor Yellow
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    } catch {
        Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexNumber = Read-Host "`nEnter the index number to export drivers from"
    $exportPath = Get-DriverExportPath -ExportType "Mounted Image (Index $indexNumber)"
    if (-not $exportPath) {
        return
    }
    
    $Global:CurrentOperation = "Exporting Drivers"
    $Global:CleanupRequired = $true
    
    try {
        if (-not (Test-Path $exportPath)) {
            New-Item -ItemType Directory -Path $exportPath -Force | Out-Null
            Write-Host "Created export directory: $exportPath" -ForegroundColor Green
        }
        
        Write-Host "`nMounting WIM/ESD index ${indexNumber}..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$indexNumber /MountDir:$Global:MountPath /CheckIntegrity /ReadOnly
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${indexNumber}"
        }
        
        Write-Host "Exporting drivers from mounted image..." -ForegroundColor Yellow
        Write-Host "Export Path: $exportPath" -ForegroundColor Gray
        Write-Host "This may take several minutes depending on the number of drivers in the image." -ForegroundColor Cyan
        
        $Global:CurrentOperation = "Exporting Drivers"
        dism /Image:$Global:MountPath /Export-Driver /Destination:$exportPath
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Drivers successfully exported from mounted image" -ForegroundColor Green
            
            $driverFiles = Get-ChildItem -Path $exportPath -Filter "*.inf" -Recurse -ErrorAction SilentlyContinue
            if ($driverFiles) {
                Write-Host "Total driver packages exported: $($driverFiles.Count)" -ForegroundColor Cyan
            }
        } else {
            Write-Host "✗ Driver export failed" -ForegroundColor Red
        }
        
        Write-Host "Unmounting image..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Unmounting"
        dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Image unmounted successfully" -ForegroundColor Green
        } else {
            Write-Host "✗ Warning: Image unmount may have had issues" -ForegroundColor Yellow
        }
        
        Write-Host "`nSUCCESS: Driver export completed!" -ForegroundColor Green
        Write-Host "Drivers exported to: $exportPath" -ForegroundColor Cyan
        
    } catch {
        Write-Host "ERROR: Failed to export drivers - $($_.Exception.Message)" -ForegroundColor Red
        
        Write-Host "Attempting to unmount image..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
    
    Read-Host "Press Enter to continue"
}

function Get-DriverExportPath {
    param([string]$ExportType)
    
    Write-Host "Driver Export Directory Configuration" -ForegroundColor Magenta
    Write-Host "Export Type: $ExportType" -ForegroundColor Gray
    Write-Host ""
    $defaultParentPath = $null
    if ($Global:DriverPath) {
        $defaultParentPath = $Global:DriverPath
        Write-Host "Using saved driver path as default." -ForegroundColor Cyan
    } else {
        $defaultParentPath = Get-Location
        Write-Host "No driver path is currently saved." -ForegroundColor Yellow
    }
    
    Write-Host "A timestamped subdirectory will be created in your chosen parent directory." -ForegroundColor Cyan
    Write-Host "Suggested parent directory: $defaultParentPath" -ForegroundColor Green
    
    do {
        $userPath = Read-Host "Enter parent directory path (Press Enter for suggested path)"
        
        if (-not $userPath) {
            $parentPath = $defaultParentPath
            $userEnteredCustomPath = $false
        } else {
            $userPath = $userPath.Trim('"', "'")
            $parentPath = [System.Environment]::ExpandEnvironmentVariables($userPath)
            $userEnteredCustomPath = $true
        }
        
        if (Test-Path $parentPath -PathType Container) {
            if ($userEnteredCustomPath -and $parentPath -ne $Global:DriverPath) {
                $saveAsDefault = Read-Host "Save this path as your default driver path for future operations? (Y/n)"
                if ($saveAsDefault.ToLower() -ne 'n') {
                    $Global:DriverPath = $parentPath
                    Write-Host "✓ Driver path saved: $parentPath" -ForegroundColor Green
                    Write-Host "This path will now be used as the default for driver operations." -ForegroundColor Cyan
                }
            }
            
            $exportPath = Join-Path $parentPath "Exported_Drivers_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            Write-Host "Drivers will be exported to: $exportPath" -ForegroundColor Green
            Write-Host "(Subdirectory will be created automatically)" -ForegroundColor Gray
            return $exportPath
        } else {
            if (Test-Path $parentPath) {
                Write-Host "Path exists but is not a directory: $parentPath" -ForegroundColor Red
            } else {
                Write-Host "Parent directory does not exist: $parentPath" -ForegroundColor Red
            }
            Write-Host "Please enter a valid existing directory path." -ForegroundColor Yellow
        }
    } while ($true)
}

function Manage-WimlibSetup {
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "    Wimlib Setup (Advanced Features)" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "ABOUT WIMLIB:" -ForegroundColor Magenta
        Write-Host "Wimlib provides advanced WIM/ESD manipulation capabilities:" -ForegroundColor Gray
        Write-Host "• Fast metadata updates (name, description, flags)" -ForegroundColor Gray
        Write-Host "• Direct file operations without mounting" -ForegroundColor Gray
        Write-Host "• Better compression algorithms" -ForegroundColor Gray
        Write-Host "• Works with both WIM and ESD files" -ForegroundColor Gray
        Write-Host ""
        
        Write-Host "CURRENT STATUS:" -ForegroundColor Magenta
        Write-Host "Wimlib Path: " -NoNewline
        if ($Global:WimlibPath) {
            $wimlibExe = Join-Path $Global:WimlibPath "wimlib-imagex.exe"
            if (Test-Path $wimlibExe) {
                Write-Host $Global:WimlibPath -ForegroundColor Green
                try {
                    $versionOutput = & $wimlibExe --version 2>$null | Select-Object -First 1
                    if ($versionOutput) {
                        Write-Host "Version: $versionOutput" -ForegroundColor Green
                    }
                } catch {
                    Write-Host "Version: Unable to detect" -ForegroundColor Yellow
                }
            } else {
                Write-Host "$Global:WimlibPath (Invalid - exe not found)" -ForegroundColor Red
                $Global:WimlibPath = ""
            }
        } else {
            Write-Host "Not configured" -ForegroundColor Red
        }
        
        $wimlibAvailable = $Global:WimlibPath -and (Test-Path (Join-Path $Global:WimlibPath "wimlib-imagex.exe"))
        Write-Host "Status: " -NoNewline
        if ($wimlibAvailable) {
            Write-Host "Ready for advanced operations" -ForegroundColor Green
        } else {
            Write-Host "Not available - limited to DISM operations" -ForegroundColor Red
        }
        Write-Host ""
        
        Write-Host "SETUP OPTIONS:" -ForegroundColor Cyan
        Write-Host "1. Download Wimlib (opens browser)" -ForegroundColor Yellow
        Write-Host "2. Set Wimlib Path (after download)" -ForegroundColor Yellow
        Write-Host "3. Test Wimlib Installation" -ForegroundColor $(if ($wimlibAvailable) { "Green" } else { "DarkGray" })
        Write-Host "4. Clear Wimlib Path" -ForegroundColor $(if ($wimlibAvailable) { "Yellow" } else { "DarkGray" })
        Write-Host "5. Return to Configuration Menu" -ForegroundColor Red
        Write-Host ""
        Write-Host "Select an option (1-5): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { Download-Wimlib }
            "2" { Set-WimlibPath }
            "3" { 
                if ($wimlibAvailable) {
                    Test-WimlibInstallation
                } else {
                    Write-Host "Wimlib is not configured. Please set up wimlib first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                }
            }
            "4" {
                if ($wimlibAvailable) {
                    Clear-WimlibPath
                } else {
                    Write-Host "Wimlib path is not set." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue"
                }
            }
            "5" { return }
            default { 
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

function Download-Wimlib {
    Write-Host "`nDownload Wimlib" -ForegroundColor Cyan
    Write-Host "===============" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Wimlib is a portable tool - no installation required!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Primary Download URL:" -ForegroundColor Yellow
    Write-Host "https://wimlib.net/downloads/wimlib-1.14.4-windows-x86_64-bin.zip" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Backup Download URL:" -ForegroundColor Yellow
    Write-Host "https://github.com/amec0e/wimlib-imagex/raw/refs/heads/main/wimlib-1.14.4-windows-x86_64-bin.zip" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Instructions:" -ForegroundColor Yellow
    Write-Host "1. Download the ZIP file from either URL above" -ForegroundColor Gray
    Write-Host "2. Extract it to a folder (e.g., C:\Tools\wimlib)" -ForegroundColor Gray
    Write-Host "3. Come back here and use 'Set Wimlib Path' option" -ForegroundColor Gray
    Write-Host "4. Point to the extracted folder containing wimlib-imagex.exe" -ForegroundColor Gray
    Write-Host ""
    
    Write-Host "Download Options:" -ForegroundColor Magenta
    Write-Host "1. Open primary download link (wimlib.net)" -ForegroundColor Green
    Write-Host "2. Open backup download link (GitHub)" -ForegroundColor Cyan
    Write-Host "3. Skip download" -ForegroundColor Yellow
    
    $choice = Read-Host "`nSelect option (1-3)"
    
    switch ($choice) {
        "1" {
            try {
                Start-Process "https://wimlib.net/downloads/wimlib-1.14.4-windows-x86_64-bin.zip"
                Write-Host "✓ Primary download link opened in browser" -ForegroundColor Green
            } catch {
                Write-Host "✗ Could not open browser. Please manually visit the primary URL above." -ForegroundColor Red
            }
        }
        "2" {
            try {
                Start-Process "https://github.com/amec0e/wimlib-imagex/raw/refs/heads/main/wimlib-1.14.4-windows-x86_64-bin.zip"
                Write-Host "✓ Backup download link opened in browser" -ForegroundColor Green
            } catch {
                Write-Host "✗ Could not open browser. Please manually visit the backup URL above." -ForegroundColor Red
            }
        }
        "3" {
            Write-Host "Download skipped." -ForegroundColor Yellow
        }
        default {
            Write-Host "Invalid selection. Download skipped." -ForegroundColor Red
        }
    }
    
    Write-Host ""
    Write-Host "After downloading and extracting, use option 2 to set the path." -ForegroundColor Cyan
    Read-Host "Press Enter to continue"
}

function Set-WimlibPath {
    Write-Host "`nSet Wimlib Path" -ForegroundColor Cyan
    Write-Host "===============" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Please provide the path to your extracted wimlib directory." -ForegroundColor Gray
    Write-Host "This should contain wimlib-imagex.exe and various .cmd helper files" -ForegroundColor Gray
    Write-Host ""
    
    if ($Global:WimlibPath) {
        Write-Host "Current path: $Global:WimlibPath" -ForegroundColor Yellow
        Write-Host ""
    }
    
    do {
        $path = Read-Host "Enter wimlib directory path (or press Enter to cancel)"
        
        if (-not $path) {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            Read-Host "Press Enter to continue"
            return
        }
        
        # Use the same validation as Set-ISOPath
        $validatedPath = Test-DirectoryPath -Path $path -PathType "Wimlib"
        if ($validatedPath) {
            # Check for wimlib-imagex.exe
            $wimlibExe = Join-Path $validatedPath "wimlib-imagex.exe"
            
            if (Test-Path $wimlibExe) {
                # Test if wimlib-imagex works
                Write-Host "Testing wimlib-imagex..." -ForegroundColor Yellow
                try {
                    $versionOutput = & $wimlibExe --version 2>$null
                    if ($LASTEXITCODE -eq 0 -and $versionOutput) {
                        $Global:WimlibPath = $validatedPath
                        Write-Host "✓ Wimlib path set successfully: ${validatedPath}" -ForegroundColor Green
                        Write-Host "✓ Version: $($versionOutput | Select-Object -First 1)" -ForegroundColor Green
                        
                        # Check for all .cmd files
                        try {
                            $cmdFiles = Get-ChildItem -Path $validatedPath -Filter "*.cmd" -ErrorAction SilentlyContinue
                            if ($cmdFiles.Count -gt 0) {
                                Write-Host "✓ Found $($cmdFiles.Count) batch helper files (.cmd):" -ForegroundColor Green
                                
                                # Show first few .cmd files as examples
                                $exampleCmds = $cmdFiles | Select-Object -First 5
                                foreach ($cmd in $exampleCmds) {
                                    Write-Host "  $($cmd.Name)" -ForegroundColor Gray
                                }
                                if ($cmdFiles.Count -gt 5) {
                                    Write-Host "  ... and $($cmdFiles.Count - 5) more" -ForegroundColor Gray
                                }
                            } else {
                                Write-Host "Note: No .cmd batch helper files found, but direct execution will work" -ForegroundColor Yellow
                            }
                        } catch {
                            Write-Host "Note: Could not scan for .cmd files, but direct execution will work" -ForegroundColor Yellow
                        }
                        
                        Write-Host ""
                        Write-Host "Advanced features are now available!" -ForegroundColor Cyan
                        Read-Host "Press Enter to continue"
                        return
                    } else {
                        Write-Host "Error: wimlib-imagex test failed (Exit code: $LASTEXITCODE)" -ForegroundColor Red
                    }
                } catch {
                    Write-Host "Error testing wimlib-imagex: $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "wimlib-imagex.exe not found in: ${validatedPath}" -ForegroundColor Red
                Write-Host "Expected: $wimlibExe" -ForegroundColor Gray
                
                # Show what files are in the directory
                try {
                    Write-Host "`nFiles found in directory:" -ForegroundColor Gray
                    $files = Get-ChildItem $validatedPath -Name | Select-Object -First 15
                    foreach ($file in $files) {
                        Write-Host "  $file" -ForegroundColor Gray
                    }
                    if ((Get-ChildItem $validatedPath).Count -gt 15) {
                        Write-Host "  ... and more" -ForegroundColor Gray
                    }
                } catch {
                    Write-Host "Could not list directory contents." -ForegroundColor Red
                }
            }
        } else {
            Write-Host "Please provide a valid existing directory path when ready." -ForegroundColor Red
        }
        
        $retry = Read-Host "`nTry again? (Y/n)"
        if ($retry.ToLower() -eq 'n') { 
            return 
        }
    } while ($true)
}

function Test-WimlibInstallation {
    Write-Host "`nTesting Wimlib Installation" -ForegroundColor Cyan
    Write-Host "============================" -ForegroundColor Cyan
    
    $wimlibExe = Join-Path $Global:WimlibPath "wimlib-imagex.exe"
    
    Write-Host "Testing wimlib-imagex.exe..." -ForegroundColor Yellow
    try {
        $versionOutput = & $wimlibExe --version
        Write-Host "✓ Version test passed" -ForegroundColor Green
        Write-Host "Version: $($versionOutput | Select-Object -First 1)" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "Available commands:" -ForegroundColor Yellow
        $commands = @("info", "apply", "capture", "export", "delete", "update", "optimize")
        foreach ($cmd in $commands) {
            try {
                & $wimlibExe $cmd --help 2>$null | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "✓ $cmd - Available" -ForegroundColor Green
                } else {
                    Write-Host "✗ $cmd - Not working" -ForegroundColor Red
                }
            } catch {
                Write-Host "✗ $cmd - Error testing" -ForegroundColor Red
            }
        }
        
        Write-Host ""
        Write-Host "✓ Wimlib installation test completed successfully!" -ForegroundColor Green
        
    } catch {
        Write-Host "✗ Wimlib test failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please check your wimlib installation." -ForegroundColor Yellow
    }
    
    Read-Host "Press Enter to continue"
}

function Clear-WimlibPath {
    Write-Host "`nClear Wimlib Path" -ForegroundColor Cyan
    Write-Host "=================" -ForegroundColor Cyan
    Write-Host "Current path: $Global:WimlibPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "This will disable advanced wimlib features." -ForegroundColor Gray
    Write-Host "You can set it up again later if needed." -ForegroundColor Gray
    
    $confirm = Read-Host "`nAre you sure you want to clear the wimlib path? (y/N)"
    if ($confirm.ToLower() -eq 'y') {
        $Global:WimlibPath = ""
        Write-Host "✓ Wimlib path cleared" -ForegroundColor Green
    } else {
        Write-Host "Operation cancelled" -ForegroundColor Yellow
    }
    
    Read-Host "Press Enter to continue"
}

function Show-AvailableKeys {
    param([string]$KeyType)
    
    if ($KeyType -eq "RTM") {
        Write-Host "`nAvailable RTM Keys:" -ForegroundColor Cyan
        Write-Host "===================================================" -ForegroundColor Cyan
        Write-Host ""
        
        $sortedKeys = $Global:RTMKeys.GetEnumerator() | Sort-Object Name
        foreach ($key in $sortedKeys) {
            Write-Host "$($key.Name): " -NoNewline -ForegroundColor Yellow
            Write-Host $key.Value -ForegroundColor Green
        }
        
        Write-Host ""
        Write-Host "Note: These are Microsoft's RTM Keys." -ForegroundColor Gray
        Write-Host "They allow installation without product key prompts for volume licensing." -ForegroundColor Gray
        
    } elseif ($KeyType -eq "KMS") {
        Write-Host "`nAvailable KMS Client Product Keys (Official Microsoft List):" -ForegroundColor Cyan
        Write-Host "=============================================================" -ForegroundColor Cyan
        Write-Host ""
        
        $sortedKeys = $Global:KMSClientKeys.GetEnumerator() | Sort-Object Name
        foreach ($key in $sortedKeys) {
            Write-Host "$($key.Name): " -NoNewline -ForegroundColor Yellow
            Write-Host $key.Value -ForegroundColor Green
        }
        
        Write-Host ""
        Write-Host "Note: Official Microsoft KMS Client Setup Keys." -ForegroundColor Gray
        Write-Host "These configure the system for KMS activation in enterprise environments." -ForegroundColor Gray
        Write-Host "Requires a KMS server for activation." -ForegroundColor Gray
    }
    
    Read-Host "Press Enter to continue"
}

function Set-ProductKeyForIndex {
    Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    } catch {
        Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexNumber = Read-Host "`nEnter the index number to set product key for"
    Set-ProductKeyForSpecificIndex -IndexNumber $indexNumber
}

function Get-ProductKeyForBaseIndex {
    param([string]$IndexNumber)
    
    $Global:CurrentOperation = "Getting Product Key for Base Index"
    $Global:CleanupRequired = $true
    
    try {
        Write-Host "`nMounting WIM/ESD index ${IndexNumber} for product key selection..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$IndexNumber /MountDir:$Global:MountPath /CheckIntegrity /ReadOnly
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${IndexNumber}"
        }
        
        # Detect edition information
        $editionInfo = Get-MountedImageEditionInfo -IndexNumber $IndexNumber
        
        Write-Host "`nDetected Edition Information for Base Index:" -ForegroundColor Cyan
        Write-Host "=============================================" -ForegroundColor Cyan
        Write-Host "Index: ${IndexNumber}" -ForegroundColor Gray
        Write-Host "Name: $($editionInfo.Name)" -ForegroundColor Yellow
        Write-Host "Description: $($editionInfo.Description)" -ForegroundColor Gray
        Write-Host "Edition: $($editionInfo.Edition)" -ForegroundColor Yellow
        Write-Host "Flags: $($editionInfo.Flags)" -ForegroundColor Gray
        Write-Host ""
        
        # Get recommended keys
        $recommendations = Get-RecommendedProductKeys -EditionInfo $editionInfo
        
        $selectedKey = $null
        if ($recommendations.Count -gt 0) {
            Write-Host "Recommended Product Keys:" -ForegroundColor Green
            Write-Host "=========================" -ForegroundColor Green
            $i = 1
            foreach ($rec in $recommendations) {
                Write-Host "$i. $($rec.Type): $($rec.EditionName)" -ForegroundColor Cyan
                Write-Host "   Key: $($rec.Key)" -ForegroundColor Green
                Write-Host "   Reason: $($rec.Reason)" -ForegroundColor Gray
                Write-Host ""
                $i++
            }
            
            Write-Host "Options:" -ForegroundColor Yellow
            Write-Host "1-$($recommendations.Count). Use recommended key" -ForegroundColor Green
            Write-Host "$($recommendations.Count + 1). Choose from all available keys" -ForegroundColor Yellow
            Write-Host "$($recommendations.Count + 2). Enter custom key" -ForegroundColor Yellow
            Write-Host "$($recommendations.Count + 3). Cancel operation" -ForegroundColor Red
            
            do {
                $choice = Read-Host "Select option"
                $choiceNum = 0
                if ([int]::TryParse($choice, [ref]$choiceNum)) {
                    if ($choiceNum -ge 1 -and $choiceNum -le $recommendations.Count) {
                        $selectedKey = $recommendations[$choiceNum - 1].Key
                        Write-Host "Selected: $($recommendations[$choiceNum - 1].EditionName) ($($recommendations[$choiceNum - 1].Type))" -ForegroundColor Green
                        break
                    } elseif ($choiceNum -eq ($recommendations.Count + 1)) {
                        $selectedKey = Get-ProductKeyChoice
                        break
                    } elseif ($choiceNum -eq ($recommendations.Count + 2)) {
                        $selectedKey = Get-CustomProductKey
                        break
                    } elseif ($choiceNum -eq ($recommendations.Count + 3)) {
                        $selectedKey = $null
                        break
                    }
                }
                Write-Host "Invalid selection. Please try again." -ForegroundColor Red
            } while ($true)
        } else {
            Write-Host "No specific recommendations found for this edition." -ForegroundColor Yellow
            $selectedKey = Get-ProductKeyChoice
        }
        
        Write-Host "`nUnmounting base index..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Unmounting"
        dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Base index unmounted successfully" -ForegroundColor Green
        } else {
            Write-Host "✗ Warning: Base index unmount may have had issues" -ForegroundColor Yellow
        }
        
        return $selectedKey
        
    } catch {
        Write-Host "ERROR: Failed to get product key for base index ${IndexNumber} - $($_.Exception.Message)" -ForegroundColor Red
        
        Write-Host "Attempting to unmount base index..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
        return $null
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
}

function Set-ProductKeyForAllIndexes {
    Write-Host "`nThis will automatically set product keys for all indexes based on exact name matching." -ForegroundColor Yellow
    Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    } catch {
        Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexes = Get-WimIndexes
    if ($indexes.Count -eq 0) {
        Write-Host "No indexes found in the image." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nAutomatic Product Key Assignment:" -ForegroundColor Cyan
    Write-Host "=================================" -ForegroundColor Cyan
    Write-Host "• Each index will be analyzed for exact name matches only" -ForegroundColor Gray
    Write-Host "• RTM keys will be preferred over KMS Client keys" -ForegroundColor Gray
    Write-Host "• Indexes without exact matches will be skipped" -ForegroundColor Gray
    Write-Host ""
    
    $autoSet = Read-Host "Proceed with automatic product key assignment? (Y/n)"
    if ($autoSet.ToLower() -eq 'n') {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        return
    }
    
    $successCount = 0
    $failCount = 0
    $skipCount = 0
    
    $Global:CurrentOperation = "Setting Product Keys (All Indexes)"
    $Global:CleanupRequired = $true
    
    # Process each index individually with full DISM output visibility
    foreach ($index in $indexes) {
        Write-Host "`n--- Processing Index ${index} ---" -ForegroundColor Magenta
        
        try {
            Write-Host "`nMounting WIM/ESD index ${index}..." -ForegroundColor Yellow
            $Global:CurrentOperation = "Mounting Index ${index}"
            dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$index /MountDir:$Global:MountPath /CheckIntegrity
            
            if ($LASTEXITCODE -ne 0) {
                Write-Host "✗ Failed to mount index ${index}" -ForegroundColor Red
                $failCount++
                continue
            }
            
            # Detect edition information
            $editionInfo = Get-MountedImageEditionInfo -IndexNumber $index
            
            Write-Host "Detected Edition: $($editionInfo.Name)" -ForegroundColor Cyan
            Write-Host "Current Edition: $($editionInfo.Edition)" -ForegroundColor Gray
            
            # Get exact matches using the same logic as single index
            $exactMatches = Get-RecommendedProductKeys -EditionInfo $editionInfo
            
            # Filter out any invalid matches
            $validMatches = @()
            foreach ($match in $exactMatches) {
                if ($match -and $match.Key -and $match.EditionName -and $match.Key.Trim() -ne "" -and $match.EditionName.Trim() -ne "") {
                    $validMatches += $match
                }
            }
            
            if ($validMatches.Count -eq 0) {
                Write-Host "✗ No exact matches found for this edition" -ForegroundColor Red
                Write-Host "Unmounting without changes..." -ForegroundColor Yellow
                dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                $skipCount++
                continue
            }
            
            # Choose the best match (prefer RTM over KMS Client)
            $selectedMatch = $null
            foreach ($match in $validMatches) {
                if ($match.Type -eq "RTM") {
                    $selectedMatch = $match
                    break
                }
            }
            if (-not $selectedMatch) {
                $selectedMatch = $validMatches[0]  # Use first available if no RTM found
            }
            
            Write-Host "Selected Key: $($selectedMatch.Key) ($($selectedMatch.Type): $($selectedMatch.EditionName))" -ForegroundColor Green
            
            Write-Host "Setting product key..." -ForegroundColor Yellow
            $Global:CurrentOperation = "Setting Product Key for Index ${index}"
            dism /Image:$Global:MountPath /Set-ProductKey:$($selectedMatch.Key)
            
            $keySetSuccess = $false
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Product key set successfully" -ForegroundColor Green
                $keySetSuccess = $true
            } else {
                Write-Host "✗ Failed to set product key" -ForegroundColor Red
            }
            
            Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
            $Global:CurrentOperation = "Unmounting Index ${index}"
            dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Changes committed successfully for index ${index}" -ForegroundColor Green
                if ($keySetSuccess) {
                    $successCount++
                } else {
                    $failCount++
                }
            } else {
                Write-Host "✗ Failed to commit changes for index ${index}" -ForegroundColor Red
                $failCount++
            }
            
        } catch {
            Write-Host "✗ ERROR processing index ${index} : $($_.Exception.Message)" -ForegroundColor Red
            
            Write-Host "Attempting to discard changes and unmount..." -ForegroundColor Yellow
            try {
                dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
            } catch {
                Write-Host "Manual cleanup may be required." -ForegroundColor Red
            }
            $failCount++
        }
    }
    
    $Global:CurrentOperation = "None"
    $Global:CleanupRequired = $false
    
    Write-Host "`nSummary:" -ForegroundColor Cyan
    Write-Host "Successfully processed: ${successCount} indexes" -ForegroundColor Green
    if ($skipCount -gt 0) {
        Write-Host "Skipped (no exact match): $skipCount indexes" -ForegroundColor Yellow
    }
    if ($failCount -gt 0) {
        Write-Host "Failed to process: $failCount indexes" -ForegroundColor Red
    }
    
    Read-Host "Press Enter to continue"
}

function Select-KeyFromCollection {
    param(
        [hashtable]$Collection,
        [string]$CollectionName
    )
    
    Write-Host "`nAvailable $CollectionName Keys:" -ForegroundColor Cyan
    Write-Host "$(('=' * ($CollectionName.Length + 6)))" -ForegroundColor Cyan
    Write-Host ""
    
    $sortedKeys = $Collection.GetEnumerator() | Sort-Object Name
    $keyList = @()
    $i = 1
    
    foreach ($key in $sortedKeys) {
        Write-Host "$i. $($key.Name)" -ForegroundColor Yellow
        Write-Host "   Key: $($key.Value)" -ForegroundColor Green
        Write-Host ""
        $keyList += @{
            Number = $i
            Name = $key.Name
            Key = $key.Value
        }
        $i++
    }
    
    Write-Host "Options:" -ForegroundColor Magenta
    Write-Host "1-$($keyList.Count). Select a key from above" -ForegroundColor Green
    Write-Host "$($keyList.Count + 1). Enter custom key" -ForegroundColor Yellow
    Write-Host "$($keyList.Count + 2). Cancel" -ForegroundColor Red
    
    do {
        $choice = Read-Host "`nSelect option"
        $choiceNum = 0
        
        if ([int]::TryParse($choice, [ref]$choiceNum)) {
            if ($choiceNum -ge 1 -and $choiceNum -le $keyList.Count) {
                $selectedKey = $keyList[$choiceNum - 1]
                Write-Host "Selected: $($selectedKey.Name)" -ForegroundColor Green
                Write-Host "Key: $($selectedKey.Key)" -ForegroundColor Cyan
                
                $confirm = Read-Host "Use this key? (Y/n)"
                if ($confirm.ToLower() -ne 'n') {
                    return $selectedKey.Key
                }
            } elseif ($choiceNum -eq ($keyList.Count + 1)) {
                return Get-CustomProductKey
            } elseif ($choiceNum -eq ($keyList.Count + 2)) {
                Write-Host "Selection cancelled." -ForegroundColor Yellow
                return $null
            }
        }
        
        Write-Host "Invalid selection. Please try again." -ForegroundColor Red
    } while ($true)
}

function Get-ProductKeyChoice {
    Write-Host "`nProduct Key Selection:" -ForegroundColor Cyan
    Write-Host "1. Select from RTM keys" -ForegroundColor Yellow
    Write-Host "2. Select from KMS keys" -ForegroundColor Yellow
    Write-Host "3. Enter custom key" -ForegroundColor Yellow
    Write-Host "4. Cancel" -ForegroundColor Red
    
    $choice = Read-Host "`nSelect option (1-4)"
    
    switch ($choice) {
        "1" {
            return Select-KeyFromCollection -Collection $Global:RTMKeys -CollectionName "RTM"
        }
        "2" {
            return Select-KeyFromCollection -Collection $Global:KMSClientKeys -CollectionName "KMS Client"
        }
        "3" {
            return Get-CustomProductKey
        }
        "4" {
            Write-Host "Selection cancelled." -ForegroundColor Yellow
            return $null
        }
        default {
            Write-Host "Invalid selection." -ForegroundColor Red
            return $null
        }
    }
}

function Set-ProductKeyForSpecificIndex {
    param(
        [string]$IndexNumber,
        [string]$ProductKey = $null,
        [switch]$SuppressPrompt
    )
    
    $Global:CurrentOperation = "Setting Product Key"
    $Global:CleanupRequired = $true
    
    try {
        Write-Host "`nMounting WIM/ESD index ${IndexNumber}..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Mounting"
        dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$IndexNumber /MountDir:$Global:MountPath /CheckIntegrity
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM/ESD index ${IndexNumber}"
        }
        
        # If no product key provided, detect edition and get recommendations (for single index mode)
        if (-not $ProductKey) {
            # Detect edition information
            $editionInfo = Get-MountedImageEditionInfo -IndexNumber $IndexNumber
            
            Write-Host "`nDetected Edition Information:" -ForegroundColor Cyan
            Write-Host "=============================" -ForegroundColor Cyan
            Write-Host "Index: ${IndexNumber}" -ForegroundColor Gray
            Write-Host "Name: $($editionInfo.Name)" -ForegroundColor Yellow
            Write-Host "Description: $($editionInfo.Description)" -ForegroundColor Gray
            Write-Host "Edition: $($editionInfo.Edition)" -ForegroundColor Yellow
            Write-Host "Flags: $($editionInfo.Flags)" -ForegroundColor Gray
            Write-Host ""
            
            # Try to find exact matches
            $exactMatches = Get-RecommendedProductKeys -EditionInfo $editionInfo
            
            if ($exactMatches.Count -gt 0) {
                # Filter out any invalid matches and ensure we have a clean array
                $validMatches = @()
                foreach ($match in $exactMatches) {
                    if ($match -and $match.Key -and $match.EditionName -and $match.Key.Trim() -ne "" -and $match.EditionName.Trim() -ne "") {
                        $validMatches += $match
                    }
                }
                
                Write-Host "Found Exact Matches:" -ForegroundColor Green
                Write-Host "====================" -ForegroundColor Green
                $i = 1
                foreach ($match in $validMatches) {
                    Write-Host "$i. $($match.Type): $($match.EditionName)" -ForegroundColor Cyan
                    Write-Host "   Key: $($match.Key)" -ForegroundColor Green
                    Write-Host ""
                    $i++
                }
                
                Write-Host "Options:" -ForegroundColor Yellow
                if ($validMatches.Count -eq 1) {
                    # Single match
                    Write-Host "1. Use the exact match" -ForegroundColor Green
                    Write-Host "2. Browse all available keys" -ForegroundColor Yellow
                    Write-Host "3. Enter custom key" -ForegroundColor Yellow
                    Write-Host "4. Skip setting product key" -ForegroundColor Red
                    
                    do {
                        $choice = Read-Host "Select option (1-4)"
                        switch ($choice) {
                            "1" {
                                $ProductKey = $validMatches[0].Key
                                Write-Host "Selected: $($validMatches[0].EditionName) ($($validMatches[0].Type))" -ForegroundColor Green
                                break
                            }
                            "2" {
                                $ProductKey = Get-ProductKeyChoice
                                if (-not $ProductKey) { 
                                    Write-Host "Operation cancelled." -ForegroundColor Yellow
                                    dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                                    return
                                }
                                break
                            }
                            "3" {
                                $ProductKey = Get-CustomProductKey
                                if (-not $ProductKey) {
                                    Write-Host "Operation cancelled." -ForegroundColor Yellow
                                    dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                                    return
                                }
                                break
                            }
                            "4" {
                                Write-Host "Skipping product key setting." -ForegroundColor Yellow
                                dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                                return
                            }
                            default {
                                Write-Host "Invalid selection. Please choose 1-4." -ForegroundColor Red
                                continue
                            }
                        }
                        break
                    } while ($true)
                } else {
                    # Multiple matches
                    Write-Host "1-$($validMatches.Count). Use exact match" -ForegroundColor Green
                    Write-Host "$($validMatches.Count + 1). Browse all available keys" -ForegroundColor Yellow
                    Write-Host "$($validMatches.Count + 2). Enter custom key" -ForegroundColor Yellow
                    Write-Host "$($validMatches.Count + 3). Skip setting product key" -ForegroundColor Red
                    
                    do {
                        $choice = Read-Host "Select option"
                        $choiceNum = 0
                        if ([int]::TryParse($choice, [ref]$choiceNum)) {
                            if ($choiceNum -ge 1 -and $choiceNum -le $validMatches.Count) {
                                $ProductKey = $validMatches[$choiceNum - 1].Key
                                Write-Host "Selected: $($validMatches[$choiceNum - 1].EditionName) ($($validMatches[$choiceNum - 1].Type))" -ForegroundColor Green
                                break
                            } elseif ($choiceNum -eq ($validMatches.Count + 1)) {
                                $ProductKey = Get-ProductKeyChoice
                                if (-not $ProductKey) { 
                                    Write-Host "Operation cancelled." -ForegroundColor Yellow
                                    dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                                    return
                                }
                                break
                            } elseif ($choiceNum -eq ($validMatches.Count + 2)) {
                                $ProductKey = Get-CustomProductKey
                                if (-not $ProductKey) {
                                    Write-Host "Operation cancelled." -ForegroundColor Yellow
                                    dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                                    return
                                }
                                break
                            } elseif ($choiceNum -eq ($validMatches.Count + 3)) {
                                Write-Host "Skipping product key setting." -ForegroundColor Yellow
                                dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                                return
                            }
                        }
                        Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                    } while ($true)
                }
            } else {
                Write-Host "No Exact Matches Found" -ForegroundColor Red
                Write-Host "======================" -ForegroundColor Red
                Write-Host "No exact name matches were found for '$($editionInfo.Name)'" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "Options:" -ForegroundColor Yellow
                Write-Host "1. Browse all available keys and select manually" -ForegroundColor Green
                Write-Host "2. Enter custom product key" -ForegroundColor Yellow
                Write-Host "3. Skip setting product key" -ForegroundColor Red
                
                do {
                    $choice = Read-Host "Select option (1-3)"
                    switch ($choice) {
                        "1" {
                            Write-Host ""
                            Write-Host "TIP: Look for a key that matches your edition, then copy and paste it as a custom key." -ForegroundColor Cyan
                            $ProductKey = Get-ProductKeyChoice
                            if (-not $ProductKey) { 
                                Write-Host "Operation cancelled." -ForegroundColor Yellow
                                dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                                return
                            }
                            break
                        }
                        "2" {
                            Write-Host ""
                            Write-Host "TIP: You can copy a key from the available collections and paste it here." -ForegroundColor Cyan
                            $ProductKey = Get-CustomProductKey
                            if (-not $ProductKey) {
                                Write-Host "Operation cancelled." -ForegroundColor Yellow
                                dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                                return
                            }
                            break
                        }
                        "3" {
                            Write-Host "Skipping product key setting." -ForegroundColor Yellow
                            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                            return
                        }
                        default {
                            Write-Host "Invalid selection. Please choose 1-3." -ForegroundColor Red
                            continue
                        }
                    }
                    break
                } while ($true)
            }
        }
        
        Write-Host "`nSetting product key..." -ForegroundColor Yellow
        Write-Host "Key: $ProductKey" -ForegroundColor Cyan
        
        $Global:CurrentOperation = "Setting Product Key"
        dism /Image:$Global:MountPath /Set-ProductKey:$ProductKey
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Product key set successfully" -ForegroundColor Green
        } else {
            Write-Host "✗ Failed to set product key" -ForegroundColor Red
        }
        
        Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
        $Global:CurrentOperation = "Unmounting"
        dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Changes committed successfully for index ${IndexNumber}" -ForegroundColor Green
        } else {
            throw "Failed to commit changes for index ${IndexNumber}"
        }
        
    } catch {
        Write-Host "ERROR: Failed to set product key for index ${IndexNumber} - $($_.Exception.Message)" -ForegroundColor Red
        
        Write-Host "Attempting to discard changes and unmount..." -ForegroundColor Yellow
        try {
            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
        } catch {
            Write-Host "Manual cleanup may be required." -ForegroundColor Red
        }
    } finally {
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
        
        if (-not $SuppressPrompt) {
            Read-Host "Press Enter to continue"
        }
    }
}

function Manage-ProductKeys {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "     Product Key Management" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "ABOUT PRODUCT KEYS:" -ForegroundColor Magenta
        Write-Host "• RTM Keys: RTM Keys for install activation" -ForegroundColor Gray
        Write-Host "• KMS Client Keys: For KMS server activation in enterprise environments" -ForegroundColor Gray
        Write-Host "• Both types allow installation without prompting for a product key" -ForegroundColor Gray
        Write-Host "• These are generic keys - users can change to their own keys later" -ForegroundColor Gray
        Write-Host "• Keys do NOT automatically activate Windows" -ForegroundColor Gray
        Write-Host ""
        
        Write-Host "CURRENT CONFIGURATION:" -ForegroundColor Magenta
        Write-Host "Install Image Path: " -NoNewline
        Write-Host $Global:InstallWimPath -ForegroundColor Green
        Write-Host "Mount Path: " -NoNewline
        Write-Host "$Global:MountPath (Ready)" -ForegroundColor Green
        
        try {
            $indexes = Get-WimIndexes
            Write-Host "Available Indexes: " -NoNewline
            if ($indexes.Count -gt 0) {
                Write-Host "$($indexes.Count) found" -ForegroundColor Green
            } else {
                Write-Host "None found" -ForegroundColor Red
            }
        } catch {
            Write-Host "Available Indexes: " -NoNewline
            Write-Host "Unable to read" -ForegroundColor Red
        }
        Write-Host ""
        
        Write-Host "PRODUCT KEY OPTIONS:" -ForegroundColor Cyan
        Write-Host "1. Set key for specific index" -ForegroundColor Yellow
        Write-Host "2. Set key for all indexes" -ForegroundColor Yellow
        Write-Host "3. View available RTM keys" -ForegroundColor Cyan
        Write-Host "4. View available KMS Client keys (Official)" -ForegroundColor Cyan
        Write-Host "5. Return to Edition Management" -ForegroundColor Red
        Write-Host ""
        Write-Host "Select an option (1-5): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { Set-ProductKeyForIndex }
            "2" { Set-ProductKeyForAllIndexes }
            "3" { Show-AvailableKeys -KeyType "RTM" }
            "4" { Show-AvailableKeys -KeyType "KMS" }
            "5" { return }
            default {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

function Get-MountedImageEditionInfo {
    param([string]$IndexNumber)
    
    $editionInfo = @{
        Name = "Unknown"
        Description = "Unknown"
        Edition = "Unknown"
        Flags = "Unknown"
        DisplayName = "Unknown"
        DisplayDescription = "Unknown"
    }
    
    try {
        # Get current edition from DISM first (most reliable)
        Write-Host "Getting current edition information with DISM..." -ForegroundColor Cyan
        $dismEdition = dism /Image:$Global:MountPath /Get-CurrentEdition
        
        foreach ($line in $dismEdition) {
            if ($line -match "Current Edition\s*:\s*(.+)") {
                $editionInfo.Edition = $matches[1].Trim()
                break
            }
        }
        
        # If DISM edition failed, try to get it from image info
        if ($editionInfo.Edition -eq "Unknown") {
            Write-Host "DISM Get-CurrentEdition failed, using image info..." -ForegroundColor Yellow
        }
        
        # Get image info from WIM file for basic details
        $dismInfo = dism /Get-WimInfo /WimFile:$Global:InstallWimPath /Index:$IndexNumber
        
        foreach ($line in $dismInfo) {
            if ($line -match "Name\s*:\s*(.+)") {
                $editionInfo.Name = $matches[1].Trim()
            } elseif ($line -match "Description\s*:\s*(.+)") {
                $editionInfo.Description = $matches[1].Trim()
            }
        }
        
        # Try with wimlib if available for NAME, DESCRIPTION, DISPLAYNAME, DISPLAYDESCRIPTION, and FLAGS
        $wimlibPath = Get-WimlibImagexPath
        if ($wimlibPath) {
            Write-Host "Using wimlib to get detailed metadata..." -ForegroundColor Cyan
            try {
                $wimlibInfo = & $wimlibPath info $Global:InstallWimPath $IndexNumber
                
                foreach ($line in $wimlibInfo) {
                    if ($line -match "^Name\s*:\s*(.+)") {
                        $wimlibName = $matches[1].Trim()
                        if ($wimlibName -and $wimlibName -ne "") {
                            $editionInfo.Name = $wimlibName
                        }
                    } elseif ($line -match "^Description\s*:\s*(.+)") {
                        $wimlibDesc = $matches[1].Trim()
                        if ($wimlibDesc -and $wimlibDesc -ne "") {
                            $editionInfo.Description = $wimlibDesc
                        }
                    } elseif ($line -match "^DISPLAYNAME\s*:\s*(.+)") {
                        $wimlibDisplayName = $matches[1].Trim()
                        if ($wimlibDisplayName -and $wimlibDisplayName -ne "") {
                            $editionInfo.DisplayName = $wimlibDisplayName
                        }
                    } elseif ($line -match "^DISPLAYDESCRIPTION\s*:\s*(.+)") {
                        $wimlibDisplayDesc = $matches[1].Trim()
                        if ($wimlibDisplayDesc -and $wimlibDisplayDesc -ne "") {
                            $editionInfo.DisplayDescription = $wimlibDisplayDesc
                        }
                    } elseif ($line -match "^FLAGS\s*:\s*(.+)") {
                        $wimlibFlags = $matches[1].Trim()
                        if ($wimlibFlags -and $wimlibFlags -ne "") {
                            $editionInfo.Flags = $wimlibFlags
                        }
                    }
                }
            } catch {
                Write-Host "Wimlib detection failed, using DISM data only..." -ForegroundColor Yellow
            }
        }
        
        # If Edition is still unknown, try to derive from Flags
        if ($editionInfo.Edition -eq "Unknown" -and $editionInfo.Flags -ne "Unknown") {
            $editionInfo.Edition = $editionInfo.Flags
        }
        
        # If DisplayName is still unknown, use Name
        if ($editionInfo.DisplayName -eq "Unknown" -and $editionInfo.Name -ne "Unknown") {
            $editionInfo.DisplayName = $editionInfo.Name
        }
        
        # If DisplayDescription is still unknown, use Description
        if ($editionInfo.DisplayDescription -eq "Unknown" -and $editionInfo.Description -ne "Unknown") {
            $editionInfo.DisplayDescription = $editionInfo.Description
        }
        
    } catch {
        Write-Host "Warning: Could not fully detect edition information: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    return $editionInfo
}

function Get-RecommendedProductKeys {
    param($EditionInfo)
    
    # Try exact matches only for the main identifiers
    $searchTerms = @(
        $EditionInfo.Name,
        $EditionInfo.DisplayName
    ) | Where-Object { $_ -and $_ -ne "Unknown" -and $_ -notmatch "Microsoft.*Operating System" } | Sort-Object -Unique
    
    Write-Host "Searching for exact matches with: $($searchTerms -join ', ')" -ForegroundColor Gray
    
    $exactMatches = @()
    
    # Search for exact matches in RTM keys first
    foreach ($searchTerm in $searchTerms) {
        $rtmMatch = Find-ExactMatchOnly -SearchTerm $searchTerm -KeyCollection $Global:RTMKeys -KeyType "RTM"
        if ($rtmMatch) {
            $exactMatches += $rtmMatch
            Write-Host "✓ Found exact RTM match: $($rtmMatch.EditionName)" -ForegroundColor Green
        }
    }
    
    # Search for exact matches in KMS Client keys
    foreach ($searchTerm in $searchTerms) {
        $kmsMatch = Find-ExactMatchOnly -SearchTerm $searchTerm -KeyCollection $Global:KMSClientKeys -KeyType "KMS Client"
        if ($kmsMatch) {
            $exactMatches += $kmsMatch
            Write-Host "✓ Found exact KMS match: $($kmsMatch.EditionName)" -ForegroundColor Green
        }
    }
    
    # Remove duplicates and filter out invalid entries
    $validMatches = @()
    $seenKeys = @()
    
    foreach ($match in $exactMatches) {
        # Skip if match is null or has empty key/name
        if (-not $match -or -not $match.Key -or -not $match.EditionName -or $match.Key.Trim() -eq "" -or $match.EditionName.Trim() -eq "") {
            continue
        }
        
        # Skip if we've already seen this key
        if ($seenKeys -contains $match.Key) {
            continue
        }
        
        $validMatches += $match
        $seenKeys += $match.Key
    }
    
    if ($validMatches.Count -eq 0) {
        Write-Host "✗ No exact matches found for any edition name" -ForegroundColor Red
    }
    
    return $validMatches
}

function Find-ExactMatchOnly {
    param($SearchTerm, $KeyCollection, $KeyType)
    
    foreach ($key in $KeyCollection.GetEnumerator()) {
        if ($key.Name -eq $SearchTerm) {
            return @{
                Type = $KeyType
                EditionName = $key.Name
                Key = $key.Value
                Reason = "Exact name match"
            }
        }
    }
    return $null
}

function Get-RecommendedProductKeys {
    param($EditionInfo)
    
    # Try exact matches only for the main identifiers
    $searchTerms = @(
        $EditionInfo.Name,
        $EditionInfo.DisplayName
    ) | Where-Object { $_ -and $_ -ne "Unknown" -and $_ -notmatch "Microsoft.*Operating System" } | Sort-Object -Unique
    
    Write-Host "Searching for exact matches with: $($searchTerms -join ', ')" -ForegroundColor Gray
    
    $exactMatches = @()
    
    # Search for exact matches in RTM keys first
    foreach ($searchTerm in $searchTerms) {
        $rtmMatch = Find-ExactMatchOnly -SearchTerm $searchTerm -KeyCollection $Global:RTMKeys -KeyType "RTM"
        if ($rtmMatch) {
            $exactMatches += $rtmMatch
            Write-Host "✓ Found exact RTM match: $($rtmMatch.EditionName)" -ForegroundColor Green
        }
    }
    
    # Search for exact matches in KMS Client keys
    foreach ($searchTerm in $searchTerms) {
        $kmsMatch = Find-ExactMatchOnly -SearchTerm $searchTerm -KeyCollection $Global:KMSClientKeys -KeyType "KMS Client"
        if ($kmsMatch) {
            $exactMatches += $kmsMatch
            Write-Host "✓ Found exact KMS match: $($kmsMatch.EditionName)" -ForegroundColor Green
        }
    }
    
    # Remove duplicates (same key from different search terms)
    $uniqueMatches = @()
    $seenKeys = @()
    
    foreach ($match in $exactMatches) {
        if ($seenKeys -notcontains $match.Key) {
            $uniqueMatches += $match
            $seenKeys += $match.Key
        }
    }
    
    if ($uniqueMatches.Count -eq 0) {
        Write-Host "✗ No exact matches found for any edition name" -ForegroundColor Red
    }
    
    return $uniqueMatches
}

function Get-CustomProductKey {
    Write-Host "`nEnter Custom Product Key:" -ForegroundColor Cyan
    Write-Host "=========================" -ForegroundColor Cyan
    Write-Host "Format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" -ForegroundColor Gray
    Write-Host "You can type in lowercase - it will be converted to uppercase automatically." -ForegroundColor Gray
    Write-Host ""
    
    do {
        $customKey = Read-Host "Enter product key (or press Enter to cancel)"
        if (-not $customKey) {
            Write-Host "Custom key entry cancelled." -ForegroundColor Yellow
            return $null
        }
        
        # Remove any spaces and convert to uppercase
        $cleanKey = $customKey.Replace(" ", "").ToUpper()
        
        # Add hyphens if user entered without them (25 characters total)
        if ($cleanKey.Length -eq 25 -and $cleanKey -notmatch '-') {
            $cleanKey = $cleanKey.Substring(0,5) + "-" + $cleanKey.Substring(5,5) + "-" + $cleanKey.Substring(10,5) + "-" + $cleanKey.Substring(15,5) + "-" + $cleanKey.Substring(20,5)
        }
        
        if ($cleanKey -match '^[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$') {
            Write-Host "Formatted key: $cleanKey" -ForegroundColor Green
            $confirm = Read-Host "Use this key? (Y/n)"
            if ($confirm.ToLower() -ne 'n') {
                return $cleanKey
            }
        } else {
            Write-Host "Invalid key format. Please use: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" -ForegroundColor Red
            Write-Host "You can enter with or without hyphens, and in upper or lowercase." -ForegroundColor Gray
        }
    } while ($true)
}

function Show-ProvisionedAppxPackages {
    param([string]$ImagePath)
    
    Write-Host "Getting provisioned application packages (this may take a moment)..." -ForegroundColor Yellow
    try {
        Write-Host "`nProvisioned Application Packages (.appx/.appxbundle):" -ForegroundColor Cyan
        Write-Host "====================================================" -ForegroundColor Cyan
        dism /Image:$ImagePath /Get-ProvisionedAppxPackages
    } catch {
        Write-Host "Error getting application packages: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Remove-AppxPackages {
    param([string]$ImagePath, [bool]$IsMultiIndex)
    
    Write-Host "`nRemoving Application Packages" -ForegroundColor Cyan
    Write-Host "Getting provisioned application packages for reference..." -ForegroundColor Yellow
    try {
        Write-Host "`nProvisioned Application Packages:" -ForegroundColor Cyan
        Write-Host "==================================" -ForegroundColor Cyan
        dism /Image:$ImagePath /Get-ProvisionedAppxPackages
    } catch {
        Write-Host "Error getting application packages: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nEnter application package names separated by commas (or press Enter to skip):"
    Write-Host "Example: Microsoft.WindowsCalculator_8wekyb3d8bbwe,Microsoft.WindowsCamera_8wekyb3d8bbwe"
    Write-Host "Note: Use the full 'PackageName' from the list above"
    
    $packageInput = Read-Host "Application package names to remove"
    if (-not $packageInput) {
        return
    }
    
    $packages = $packageInput -split ',' | ForEach-Object { $_.Trim() }
    
    if ($IsMultiIndex) {
        $Global:RemovedAppxPackages += $packages
        $applyToAll = Read-Host "`nApply these application package removals to all indexes? (Y/n)"
        if ($applyToAll.ToLower() -eq 'n') {
            $Global:RemovedAppxPackages = @()
        }
    }
    
    Write-Host "`nRemoving application packages..." -ForegroundColor Yellow
    foreach ($package in $packages) {
        Write-Host "Removing application package: $package" -ForegroundColor Gray
        try {
            dism /Image:$ImagePath /Remove-ProvisionedAppxPackage /PackageName:$package
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Successfully removed application package: $package" -ForegroundColor Green
            } else {
                Write-Host "✗ Failed to remove application package: $package" -ForegroundColor Red
            }
        } catch {
            Write-Host "✗ Error removing application package $package : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Read-Host "`nPress Enter to continue"
}

function Add-AppxPackages {
    param([string]$ImagePath, [bool]$IsMultiIndex)
    
    Write-Host "`nAdding Application Packages" -ForegroundColor Cyan
    $appxPackagePath = $Global:AppxPackagePath
    if (-not $appxPackagePath) {
        do {
            $path = Read-Host "Enter the path to your APPX/APPXBUNDLE packages folder"
            $validatedPath = Test-DirectoryPath -Path $path -PathType "APPX Package"
            if ($validatedPath) {
                $appxPackagePath = $validatedPath
                $save = Read-Host "Save this APPX package path for future use? (Y/n)"
                if ($save.ToLower() -ne 'n') {
                    $Global:AppxPackagePath = $appxPackagePath
                }
                break
            } else {
                Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
            }
        } while ($true)
    } else {
        Write-Host "Using saved APPX package path: ${appxPackagePath}" -ForegroundColor Green
        $useExisting = Read-Host "Use this path? (Y/n)"
        if ($useExisting.ToLower() -eq 'n') {
            do {
                $path = Read-Host "Enter the path to your APPX/APPXBUNDLE packages folder"
                $validatedPath = Test-DirectoryPath -Path $path -PathType "APPX Package"
                if ($validatedPath) {
                    $appxPackagePath = $validatedPath
                    break
                } else {
                    Write-Host "Please enter a valid existing directory path." -ForegroundColor Red
                }
            } while ($true)
        }
    }
    
    Write-Host "`nScanning for APPX/APPXBUNDLE packages in: ${appxPackagePath}" -ForegroundColor Yellow
    $appxFiles = Get-ChildItem -Path $appxPackagePath -Filter "*.appx" -Recurse -ErrorAction SilentlyContinue
    $appxbundleFiles = Get-ChildItem -Path $appxPackagePath -Filter "*.appxbundle" -Recurse -ErrorAction SilentlyContinue
    
    $packages = @()
    if ($appxFiles) {
        $packages += $appxFiles.FullName
    }
    if ($appxbundleFiles) {
        $packages += $appxbundleFiles.FullName
    }
    
    $packages = $packages | Where-Object { $_ -and $_.Trim() }
    
    if ($packages.Count -eq 0) {
        Write-Host "No .appx or .appxbundle files found in: ${appxPackagePath}" -ForegroundColor Red
        Read-Host "`nPress Enter to continue"
        return
    }
    
    Write-Host "Found $($packages.Count) application package file(s):" -ForegroundColor Green
    foreach ($pkg in $packages) {
        $fileType = if ($pkg -like "*.appxbundle") { "APPXBUNDLE" } else { "APPX" }
        Write-Host "  - $(Split-Path $pkg -Leaf) [$fileType]" -ForegroundColor Gray
    }
    
    # Group packages by type for processing
    $appxPackages = $packages | Where-Object { $_ -like "*.appx" }
    $appxbundlePackages = $packages | Where-Object { $_ -like "*.appxbundle" }
    
    $proceed = Read-Host "`nProceed with adding these application packages? (Y/n)"
    if ($proceed.ToLower() -eq 'n') {
        return
    }
    
    if ($IsMultiIndex) {
        $Global:AddedAppxPackages += $packages
        $applyToAll = Read-Host "`nApply these application packages to all indexes? (Y/n)"
        if ($applyToAll.ToLower() -eq 'n') {
            $Global:AddedAppxPackages = @()
        }
    }
    
    Write-Host "`nAdding application packages..." -ForegroundColor Yellow
    
    # Process APPX files (may need dependencies)
    foreach ($package in $appxPackages) {
        Write-Host "`nProcessing APPX: $(Split-Path $package -Leaf)" -ForegroundColor Cyan
        
        # Check for dependencies and license
        $packageDir = Split-Path $package -Parent
        $packageName = [System.IO.Path]::GetFileNameWithoutExtension($package)
        
        # Look for dependencies
        $dependencies = Get-ChildItem -Path $packageDir -Filter "*dependency*.appx" -ErrorAction SilentlyContinue
        $frameworkDeps = Get-ChildItem -Path $packageDir -Filter "*framework*.appx" -ErrorAction SilentlyContinue
        $allDeps = @()
        if ($dependencies) { $allDeps += $dependencies.FullName }
        if ($frameworkDeps) { $allDeps += $frameworkDeps.FullName }
        
        # Look for license file
        $licenseFile = Get-ChildItem -Path $packageDir -Filter "*license*.xml" -ErrorAction SilentlyContinue | Select-Object -First 1
        
        try {
            if ($allDeps.Count -gt 0 -or $licenseFile) {
                Write-Host "  Using folder-based installation..." -ForegroundColor Gray
                
                # Use folder path method for complex installations
                $dismArgs = @("/Image:$ImagePath", "/Add-ProvisionedAppxPackage", "/FolderPath:$packageDir")
                
                if ($allDeps.Count -gt 0) {
                    Write-Host "  Found dependencies: $($allDeps.Count)" -ForegroundColor Yellow
                    foreach ($dep in $allDeps) {
                        $dismArgs += "/DependencyPackagePath:$dep"
                        Write-Host "    - $(Split-Path $dep -Leaf)" -ForegroundColor Gray
                    }
                }
                
                if ($licenseFile) {
                    Write-Host "  Found license: $(Split-Path $licenseFile.FullName -Leaf)" -ForegroundColor Yellow
                    $dismArgs += "/LicensePath:$($licenseFile.FullName)"
                }
                
                dism @dismArgs
            } else {
                Write-Host "  Using direct package installation..." -ForegroundColor Gray
                dism /Image:$ImagePath /Add-ProvisionedAppxPackage /PackagePath:$package /SkipLicense
            }
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  ✓ Successfully added: $(Split-Path $package -Leaf)" -ForegroundColor Green
            } else {
                Write-Host "  ✗ Failed to add: $(Split-Path $package -Leaf)" -ForegroundColor Red
            }
        } catch {
            Write-Host "  ✗ Error adding $(Split-Path $package -Leaf) : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # Process APPXBUNDLE files
    foreach ($package in $appxbundlePackages) {
        Write-Host "`nProcessing APPXBUNDLE: $(Split-Path $package -Leaf)" -ForegroundColor Cyan
        try {
            dism /Image:$ImagePath /Add-ProvisionedAppxPackage /PackagePath:$package /SkipLicense
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  ✓ Successfully added: $(Split-Path $package -Leaf)" -ForegroundColor Green
            } else {
                Write-Host "  ✗ Failed to add: $(Split-Path $package -Leaf)" -ForegroundColor Red
            }
        } catch {
            Write-Host "  ✗ Error adding $(Split-Path $package -Leaf) : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Read-Host "`nPress Enter to continue"
}

function Create-ISO {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("UEFI", "BIOS", "Hybrid")]
        [string]$BootMode
    )
    
    $outputPath = Get-ISOOutputPath
    if (-not $outputPath) { return }
    
    try {
        Write-Host "`nCreating $BootMode ISO..." -ForegroundColor Yellow
        $oscdimgPath = Get-OSCDImgPath
        $bootParams = Get-BootParameters -BootMode $BootMode
        if (-not $bootParams) {
            Write-Host "Failed to determine boot parameters for $BootMode mode." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        $arguments = @(
            "-m"
            "-o"
            "-u2"
            "-udfver102"
            $bootParams
            "$Global:ExtractedISOPath"
            "$outputPath"
        )
        
        Write-Host "Boot Mode: $BootMode" -ForegroundColor Cyan
        Write-Host "Boot Parameters: $bootParams" -ForegroundColor Gray
        Write-Host "Running: $oscdimgPath $($arguments -join ' ')" -ForegroundColor Gray
        Write-Host "`nISO Creation Output:" -ForegroundColor Cyan
        Write-Host "===================" -ForegroundColor Cyan
        
        & $oscdimgPath @arguments
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "`nSUCCESS: Created $BootMode ISO at $outputPath" -ForegroundColor Green
            try {
                $fileInfo = Get-Item $outputPath
                $sizeGB = [math]::Round($fileInfo.Length / 1GB, 2)
                Write-Host "ISO Size: $sizeGB GB" -ForegroundColor Cyan
            } catch {}
            
        } else {
            Write-Host "`nFAILED: ISO creation failed with exit code $LASTEXITCODE" -ForegroundColor Red
        }
        
    } catch {
        Write-Host "Error creating ISO: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Read-Host "Press Enter to continue"
}

function Update-ExportedIndexMetadata {
    param(
        [string]$IndexName,
        [string]$NewName,
        [string]$NewDescription,
        [string]$NewFlags = $null
    )
    
    $wimlibPath = Get-WimlibImagexPath
    if (-not $wimlibPath) {
        Write-Host "Wimlib not available - skipping advanced metadata update" -ForegroundColor Yellow
        return $false
    }
    
    try {
        # Find the index number by searching for the exported name
        $wimInfo = & $wimlibPath info $Global:InstallWimPath
        $indexNumber = $null
        
        # Parse wimlib output to find the index with matching name
        $currentIndex = $null
        foreach ($line in $wimInfo) {
            if ($line -match "Index\s*:\s*(\d+)") {
                $currentIndex = $matches[1]
            } elseif ($line -match "Name\s*:\s*(.+)" -and $currentIndex) {
                $foundName = $matches[1].Trim()
                if ($foundName -eq $IndexName) {
                    $indexNumber = $currentIndex
                    break
                }
            }
        }
        
        if ($indexNumber) {
            Write-Host "  Updating metadata for index ${indexNumber}..." -ForegroundColor Cyan
            
            # Build command for updating name, description, and display properties
            $args = @("info", $Global:InstallWimPath, $indexNumber, $NewName, $NewDescription)
            
            # Add standard display properties
            $args += "--image-property"
            $args += "DISPLAYNAME=$NewName"
            $args += "--image-property"
            $args += "DISPLAYDESCRIPTION=$NewDescription"
            
            # Add FLAGS and EDITIONID if provided
            if ($NewFlags -and $NewFlags.Trim() -ne "" -and $NewFlags -ne "Unknown") {
                $args += "--image-property"
                $args += "FLAGS=$NewFlags"
                $args += "--image-property"
                $args += "EDITIONID=$newFlags"
            }
            
            & $wimlibPath @args
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  ✓ Metadata updated: '$NewName'" -ForegroundColor Green
                if ($NewFlags -and $NewFlags -ne "Unknown") {
                    Write-Host "  ✓ FLAGS and EDITIONID set to: '$NewFlags'" -ForegroundColor Green
                }
                return $true
            } else {
                Write-Host "  ✗ Failed to update metadata for '$IndexName'" -ForegroundColor Red
                return $false
            }
        } else {
            Write-Host "  ✗ Could not find exported index '$IndexName'" -ForegroundColor Red
            return $false
        }
        
    } catch {
        Write-Host "  ✗ Error updating metadata: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Mount-IndexOnly {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nMount Index" -ForegroundColor Cyan
    Write-Host "===========" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This will mount an index for manual operations." -ForegroundColor Gray
    Write-Host "Use the Cleanup/Unmount option when finished." -ForegroundColor Yellow
    Write-Host ""
    
    Write-Host "Available indexes in WIM/ESD:" -ForegroundColor Yellow
    try {
        dism /Get-WimInfo /WimFile:$Global:InstallWimPath
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    } catch {
        Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $availableIndexes = Get-WimIndexes
    if ($availableIndexes.Count -eq 0) {
        Write-Host "No indexes found in the image." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $indexNumber = Read-Host "`nEnter the index number to mount"
    
    # Validate the index number
    if ($availableIndexes -notcontains $indexNumber) {
        Write-Host "Invalid index number: ${indexNumber}" -ForegroundColor Red
        Write-Host "Available indexes: $($availableIndexes -join ', ')" -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    Write-Host "`nMount Options:" -ForegroundColor Cyan
    Write-Host "1. Mount Read-Write (for making changes)" -ForegroundColor Yellow
    Write-Host "2. Mount Read-Only (for viewing only)" -ForegroundColor Green
    Write-Host "3. Cancel" -ForegroundColor Red
    
    $mountChoice = Read-Host "`nSelect mount option (1-3)"
    
    $readOnly = $false
    switch ($mountChoice) {
        "1" { 
            $readOnly = $false
            Write-Host "Mounting as Read-Write..." -ForegroundColor Yellow
        }
        "2" { 
            $readOnly = $true
            Write-Host "Mounting as Read-Only..." -ForegroundColor Green
        }
        "3" { 
            Write-Host "Mount operation cancelled." -ForegroundColor Yellow
            Read-Host "Press Enter to continue"
            return
        }
        default {
            Write-Host "Invalid selection. Defaulting to Read-Only mount." -ForegroundColor Yellow
            $readOnly = $true
        }
    }
    
    $Global:CurrentOperation = "Mounting Index ${indexNumber}"
    $Global:CleanupRequired = $true
    
    try {
        Write-Host "`nMounting WIM/ESD index ${indexNumber}..." -ForegroundColor Yellow
        Write-Host "Source: $Global:InstallWimPath" -ForegroundColor Gray
        Write-Host "Mount Point: $Global:MountPath" -ForegroundColor Gray
        Write-Host "Mode: $(if ($readOnly) { "Read-Only" } else { "Read-Write" })" -ForegroundColor Gray
        Write-Host ""
        
        if ($readOnly) {
            dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$indexNumber /MountDir:$Global:MountPath /CheckIntegrity /ReadOnly
        } else {
            dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$indexNumber /MountDir:$Global:MountPath /CheckIntegrity
        }
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "`n✓ SUCCESS: Index ${indexNumber} mounted successfully!" -ForegroundColor Green
            Write-Host ""
            Write-Host "Mount Details:" -ForegroundColor Cyan
            Write-Host "==============" -ForegroundColor Cyan
            Write-Host "Index: ${indexNumber}" -ForegroundColor Yellow
            Write-Host "Mount Path: $Global:MountPath" -ForegroundColor Yellow
            Write-Host "Access Mode: $(if ($readOnly) { "Read-Only" } else { "Read-Write" })" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "You can now:" -ForegroundColor Magenta
            if ($readOnly) {
                Write-Host "• Browse the mounted image files" -ForegroundColor Gray
                Write-Host "• View image contents and configuration" -ForegroundColor Gray
                Write-Host "• Use other tools to inspect the image" -ForegroundColor Gray
            } else {
                Write-Host "• Use DISM commands directly on the mounted image" -ForegroundColor Gray
                Write-Host "• Make manual modifications to the image" -ForegroundColor Gray
                Write-Host "• Use other tools to modify the image" -ForegroundColor Gray
            }
            Write-Host ""
            Write-Host "IMPORTANT:" -ForegroundColor Red
            Write-Host "• Use option 10 (Cleanup/Unmount) when finished" -ForegroundColor Yellow
            Write-Host "• Choose 'Commit' to save changes or 'Discard' to abandon them" -ForegroundColor Yellow
            if (-not $readOnly) {
                Write-Host "• The script will remain in 'mounted' state until you unmount" -ForegroundColor Yellow
            }
            
            # Update global state to reflect that we're no longer in a "clean" state
            $Global:CleanupRequired = $true
            
        } else {
            throw "Failed to mount WIM/ESD index ${indexNumber} (Exit code: $LASTEXITCODE)"
        }
        
    } catch {
        Write-Host "`nERROR: Failed to mount index ${indexNumber} - $($_.Exception.Message)" -ForegroundColor Red
        $Global:CurrentOperation = "None"
        $Global:CleanupRequired = $false
    }
    
    Read-Host "Press Enter to continue"
}

function Get-OSCDImgPath {
   $paths = @(
       "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
       "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
       "${env:ProgramFiles(x86)}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
       "${env:ProgramFiles}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"
   )
   
   foreach ($path in $paths) {
       if (Test-Path $path) {
           return $path
       }
   }
   
   try {
       $oscdimg = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
       if ($oscdimg) {
           return $oscdimg.Source
       }
   } catch {}
   
   return $null
}

function Manage-Registry {
    if (-not (Test-RequiredPaths)) {
        Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    if (-not (Test-MountPathEmpty)) {
        Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "       Registry Management" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "CURRENT CONFIGURATION:" -ForegroundColor Magenta
        Write-Host "Install Image Path: " -NoNewline
        Write-Host $Global:InstallWimPath -ForegroundColor Green
        Write-Host "Mount Path: " -NoNewline
        Write-Host "$Global:MountPath (Ready)" -ForegroundColor Green
        if ($Global:BackupPath) {
            Write-Host "Backup Path: " -NoNewline
            Write-Host $Global:BackupPath -ForegroundColor Green
        }
        Write-Host ""
        
        Write-Host "REGISTRY OPERATIONS:" -ForegroundColor Cyan
        Write-Host "1. Apply Common Tweaks" -ForegroundColor Green
        Write-Host "2. Return to Main Menu" -ForegroundColor Red
        Write-Host ""
        Write-Host "Select an option (1-2): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { Apply-CommonTweaks }
            "2" { return }
            default {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

function Apply-CommonTweaks {
    # Track which tweaks are selected
    $tweakStates = @{
        "DisableSearchIndexing" = $false
        "DisableTelemetry" = $false
        "ClassicContextMenu" = $false
        "DisableAutoRun" = $false
        "EnableLongPaths" = $false
        "BypassTPM" = $false
        "BypassRAM" = $false
        "BypassSecureBoot" = $false
        "BypassCPU" = $false
        "DisableMSAccount" = $false
        "DisableWeakSSLTLS" = $false
        "EnableSMBSigning" = $false
        "DisableIE11" = $false
        "EnableCertPaddingCheck" = $false
        "RestrictNetworkLocation" = $false
        "PreventDeviceEncryption" = $false
        "EnableADRecoveryBackup" = $false
    }
    
    # Selection interface
    do {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "      Apply Common Tweaks" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "Select tweaks to apply:" -ForegroundColor Yellow
        Write-Host ""
        
        # Display all tweaks with current state
        Write-Host "1. $(if ($tweakStates.DisableSearchIndexing) { '✓' } else { '✗' }) Disable Windows Search Indexing" -ForegroundColor $(if ($tweakStates.DisableSearchIndexing) { "Green" } else { "Red" })
        Write-Host "2. $(if ($tweakStates.DisableTelemetry) { '✓' } else { '✗' }) Disable Telemetry/Data Collection" -ForegroundColor $(if ($tweakStates.DisableTelemetry) { "Green" } else { "Red" })
        Write-Host "3. $(if ($tweakStates.ClassicContextMenu) { '✓' } else { '✗' }) Classic Context Menu (Windows 11)" -ForegroundColor $(if ($tweakStates.ClassicContextMenu) { "Green" } else { "Red" })
        Write-Host "4. $(if ($tweakStates.DisableAutoRun) { '✓' } else { '✗' }) Disable AutoRun/AutoPlay" -ForegroundColor $(if ($tweakStates.DisableAutoRun) { "Green" } else { "Red" })
        Write-Host "5. $(if ($tweakStates.EnableLongPaths) { '✓' } else { '✗' }) Enable Win32 Long Paths" -ForegroundColor $(if ($tweakStates.EnableLongPaths) { "Green" } else { "Red" })
        Write-Host "6. $(if ($tweakStates.BypassTPM) { '✓' } else { '✗' }) Bypass TPM Requirement" -ForegroundColor $(if ($tweakStates.BypassTPM) { "Green" } else { "Red" })
        Write-Host "7. $(if ($tweakStates.BypassRAM) { '✓' } else { '✗' }) Bypass RAM Requirement" -ForegroundColor $(if ($tweakStates.BypassRAM) { "Green" } else { "Red" })
        Write-Host "8. $(if ($tweakStates.BypassSecureBoot) { '✓' } else { '✗' }) Bypass Secure Boot Requirement" -ForegroundColor $(if ($tweakStates.BypassSecureBoot) { "Green" } else { "Red" })
        Write-Host "9. $(if ($tweakStates.BypassCPU) { '✓' } else { '✗' }) Bypass CPU Requirement" -ForegroundColor $(if ($tweakStates.BypassCPU) { "Green" } else { "Red" })
        Write-Host "10. $(if ($tweakStates.DisableMSAccount) { '✓' } else { '✗' }) Disable Microsoft Account Requirement" -ForegroundColor $(if ($tweakStates.DisableMSAccount) { "Green" } else { "Red" })
        Write-Host "11. $(if ($tweakStates.DisableWeakSSLTLS) { '✓' } else { '✗' }) Disable Weak SSL/TLS Protocols" -ForegroundColor $(if ($tweakStates.DisableWeakSSLTLS) { "Green" } else { "Red" })
        Write-Host "12. $(if ($tweakStates.EnableSMBSigning) { '✓' } else { '✗' }) Enable SMB Security Signing" -ForegroundColor $(if ($tweakStates.EnableSMBSigning) { "Green" } else { "Red" })
        Write-Host "13. $(if ($tweakStates.DisableIE11) { '✓' } else { '✗' }) Disable Internet Explorer 11" -ForegroundColor $(if ($tweakStates.DisableIE11) { "Green" } else { "Red" })
        Write-Host "14. $(if ($tweakStates.EnableCertPaddingCheck) { '✓' } else { '✗' }) Enable Certificate Padding Check" -ForegroundColor $(if ($tweakStates.EnableCertPaddingCheck) { "Green" } else { "Red" })
        Write-Host "15. $(if ($tweakStates.RestrictNetworkLocation) { '✓' } else { '✗' }) Restrict Network Location Changes" -ForegroundColor $(if ($tweakStates.RestrictNetworkLocation) { "Green" } else { "Red" })
        Write-Host "16. $(if ($tweakStates.PreventDeviceEncryption) { '✓' } else { '✗' }) Enable Prevent Device Encryption (Disable BitLocker)" -ForegroundColor $(if ($tweakStates.PreventDeviceEncryption) { "Green" } else { "Red" })
        Write-Host "17. $(if ($tweakStates.EnableADRecoveryBackup) { '✓' } else { '✗' }) Enable AD Recovery Key Backup" -ForegroundColor $(if ($tweakStates.EnableADRecoveryBackup) { "Green" } else { "Red" })
        Write-Host ""
        
        $selectedCount = 0
        foreach ($value in $tweakStates.Values) {
            if ($value -eq $true) {
                $selectedCount++
            }
        }
        
        $toggleAllStatus = if ($selectedCount -eq 0) { "All Disabled" } elseif ($selectedCount -eq $tweakStates.Count) { "All Enabled" } else { "Mixed ($selectedCount selected)" }
        
        Write-Host "18. Toggle All (Currently: $toggleAllStatus)" -ForegroundColor Magenta
        Write-Host "19. Apply Selected Tweaks ($selectedCount selected)" -ForegroundColor $(if ($selectedCount -gt 0) { "Green" } else { "DarkGray" })
        Write-Host "20. Return to Registry Management" -ForegroundColor Red
        Write-Host ""
        Write-Host "Select option (1-20): " -NoNewline -ForegroundColor White
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { $tweakStates.DisableSearchIndexing = -not $tweakStates.DisableSearchIndexing }
            "2" { $tweakStates.DisableTelemetry = -not $tweakStates.DisableTelemetry }
            "3" { $tweakStates.ClassicContextMenu = -not $tweakStates.ClassicContextMenu }
            "4" { $tweakStates.DisableAutoRun = -not $tweakStates.DisableAutoRun }
            "5" { $tweakStates.EnableLongPaths = -not $tweakStates.EnableLongPaths }
            "6" { $tweakStates.BypassTPM = -not $tweakStates.BypassTPM }
            "7" { $tweakStates.BypassRAM = -not $tweakStates.BypassRAM }
            "8" { $tweakStates.BypassSecureBoot = -not $tweakStates.BypassSecureBoot }
            "9" { $tweakStates.BypassCPU = -not $tweakStates.BypassCPU }
            "10" { $tweakStates.DisableMSAccount = -not $tweakStates.DisableMSAccount }
            "11" { $tweakStates.DisableWeakSSLTLS = -not $tweakStates.DisableWeakSSLTLS }
            "12" { $tweakStates.EnableSMBSigning = -not $tweakStates.EnableSMBSigning }
            "13" { $tweakStates.DisableIE11 = -not $tweakStates.DisableIE11 }
            "14" { $tweakStates.EnableCertPaddingCheck = -not $tweakStates.EnableCertPaddingCheck }
            "15" { $tweakStates.RestrictNetworkLocation = -not $tweakStates.RestrictNetworkLocation }
            "16" { $tweakStates.PreventDeviceEncryption = -not $tweakStates.PreventDeviceEncryption }
            "17" { $tweakStates.EnableADRecoveryBackup = -not $tweakStates.EnableADRecoveryBackup }
            "18" { 
                # Explicit toggle all logic
                if ($selectedCount -eq $tweakStates.Count) {
                    # All are enabled, so disable all
                    $tweakStates.DisableSearchIndexing = $false
                    $tweakStates.DisableTelemetry = $false
                    $tweakStates.ClassicContextMenu = $false
                    $tweakStates.DisableAutoRun = $false
                    $tweakStates.EnableLongPaths = $false
                    $tweakStates.BypassTPM = $false
                    $tweakStates.BypassRAM = $false
                    $tweakStates.BypassSecureBoot = $false
                    $tweakStates.BypassCPU = $false
                    $tweakStates.DisableMSAccount = $false
                    $tweakStates.DisableWeakSSLTLS = $false
                    $tweakStates.EnableSMBSigning = $false
                    $tweakStates.DisableIE11 = $false
                    $tweakStates.EnableCertPaddingCheck = $false
                    $tweakStates.RestrictNetworkLocation = $false
                    $tweakStates.PreventDeviceEncryption = $false
                    $tweakStates.EnableADRecoveryBackup = $false
                } else {
                    # Not all are enabled, so enable all
                    $tweakStates.DisableSearchIndexing = $true
                    $tweakStates.DisableTelemetry = $true
                    $tweakStates.ClassicContextMenu = $true
                    $tweakStates.DisableAutoRun = $true
                    $tweakStates.EnableLongPaths = $true
                    $tweakStates.BypassTPM = $true
                    $tweakStates.BypassRAM = $true
                    $tweakStates.BypassSecureBoot = $true
                    $tweakStates.BypassCPU = $true
                    $tweakStates.DisableMSAccount = $true
                    $tweakStates.DisableWeakSSLTLS = $true
                    $tweakStates.EnableSMBSigning = $true
                    $tweakStates.DisableIE11 = $true
                    $tweakStates.EnableCertPaddingCheck = $true
                    $tweakStates.RestrictNetworkLocation = $true
                    $tweakStates.PreventDeviceEncryption = $true
                    $tweakStates.EnableADRecoveryBackup = $true
                }
            }
            "19" { 
                if ($selectedCount -gt 0) {
                    # Application phase - apply the selected tweaks
                    Write-Host "`nRegistry Tweaks Options:" -ForegroundColor Cyan
                    Write-Host "1. Apply to current system (live)" -ForegroundColor Yellow
                    Write-Host "2. Apply to specific index" -ForegroundColor Yellow
                    Write-Host "3. Apply to all indexes" -ForegroundColor Yellow
                    Write-Host "4. Cancel" -ForegroundColor Red
                    Write-Host ""
                    Write-Host "Select option (1-4): " -NoNewline -ForegroundColor White
                    $applyChoice = Read-Host
                    
                    $Global:CurrentOperation = "Applying Registry Tweaks"
                    $Global:CleanupRequired = $true
                    
                    try {
                        if ($applyChoice -eq "1") {
                            Write-Host "`nApplying tweaks to current live system..." -ForegroundColor Cyan
                            Write-Host "⚠ WARNING: This will modify your current Windows installation!" -ForegroundColor Red
                            $confirm = Read-Host "Are you sure you want to proceed? (Y/n)"
                            if ($confirm.ToLower() -ne 'n') {
                                $result = Apply-TweaksToCurrentSystem -TweakStates $tweakStates
                                Write-Host "`nTweaks application completed: $($result.Success) successful, $($result.Failed) failed" -ForegroundColor Cyan
                                if ($result.RequiresReboot) {
                                    Write-Host "⚠ Some changes require a system reboot to take effect." -ForegroundColor Yellow
                                }
                            } else {
                                Write-Host "Operation cancelled." -ForegroundColor Yellow
                            }
                        }
                        elseif ($applyChoice -eq "2") {
                            Write-Host "`nAvailable indexes in WIM/ESD:" -ForegroundColor Cyan
                            try {
                                dism /Get-WimInfo /WimFile:$Global:InstallWimPath
                                if ($LASTEXITCODE -ne 0) {
                                    Write-Host "Error getting WIM/ESD information." -ForegroundColor Red
                                    Read-Host "Press Enter to continue"
                                    continue
                                }
                            } catch {
                                Write-Host "Error getting WIM/ESD information: $($_.Exception.Message)" -ForegroundColor Red
                                Read-Host "Press Enter to continue"
                                continue
                            }
                            
                            $indexNumber = Read-Host "`nEnter the index number"
                            
                            # Mount, apply tweaks, and unmount for single index
                            Write-Host "`nMounting WIM/ESD index ${indexNumber}..." -ForegroundColor Yellow
                            dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$indexNumber /MountDir:$Global:MountPath /CheckIntegrity
                            
                            if ($LASTEXITCODE -ne 0) {
                                throw "Failed to mount WIM/ESD index ${indexNumber}"
                            }
                            
                            # Apply tweaks to mounted index
                            $results = Apply-TweaksToMountedImage -IndexNumber $indexNumber -TweakStates $tweakStates
                            
                            Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
                            dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
                            
                            if ($LASTEXITCODE -eq 0) {
                                Write-Host "✓ Changes committed successfully for index ${indexNumber}" -ForegroundColor Green
                                Write-Host "Applied $($results.Success) tweaks successfully, $($results.Failed) failed" -ForegroundColor Cyan
                            } else {
                                throw "Failed to commit changes for index ${indexNumber}"
                            }
                            
                        } elseif ($applyChoice -eq "3") {
                            $indexes = Get-WimIndexes
                            if ($indexes.Count -eq 0) {
                                Write-Host "No indexes found in the image." -ForegroundColor Red
                                Read-Host "Press Enter to continue"
                                continue
                            }
                            
                            Write-Host "`nThis will apply tweaks to ALL indexes ($($indexes.Count) found)." -ForegroundColor Yellow
                            $proceed = Read-Host "Proceed with applying tweaks to all indexes? (Y/n)"
                            if ($proceed.ToLower() -eq 'n') {
                                continue
                            }
                            
                            $totalSuccess = 0
                            $totalFailed = 0
                            
                            foreach ($index in $indexes) {
                                Write-Host "`n--- Processing Index ${index} ---" -ForegroundColor Magenta
                                
                                try {
                                    Write-Host "Mounting WIM/ESD index ${index}..." -ForegroundColor Yellow
                                    dism /Mount-Wim /WimFile:$Global:InstallWimPath /Index:$index /MountDir:$Global:MountPath /CheckIntegrity
                                    
                                    if ($LASTEXITCODE -ne 0) {
                                        throw "Failed to mount WIM/ESD index ${index}"
                                    }
                                    
                                    # Apply tweaks to mounted index
                                    $results = Apply-TweaksToMountedImage -IndexNumber $index -TweakStates $tweakStates
                                    $totalSuccess += $results.Success
                                    $totalFailed += $results.Failed
                                    
                                    Write-Host "Committing changes and unmounting..." -ForegroundColor Yellow
                                    dism /Unmount-Wim /MountDir:$Global:MountPath /Commit /CheckIntegrity
                                    
                                    if ($LASTEXITCODE -eq 0) {
                                        Write-Host "✓ Changes committed successfully for index ${index}" -ForegroundColor Green
                                    } else {
                                        throw "Failed to commit changes for index ${index}"
                                    }
                                    
                                } catch {
                                    Write-Host "✗ ERROR processing index ${index} : $($_.Exception.Message)" -ForegroundColor Red
                                    Write-Host "Attempting to discard changes and unmount..." -ForegroundColor Yellow
                                    try {
                                        dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                                    } catch {
                                        Write-Host "Manual cleanup may be required." -ForegroundColor Red
                                    }
                                }
                            }
                            
                            Write-Host "`nOverall Results:" -ForegroundColor Cyan
                            Write-Host "Total tweaks applied successfully: $totalSuccess" -ForegroundColor Green
                            if ($totalFailed -gt 0) {
                                Write-Host "Total tweaks failed: $totalFailed" -ForegroundColor Red
                            }
                        } elseif ($applyChoice -eq "4") {
                            Write-Host "Operation cancelled." -ForegroundColor Yellow
                        } else {
                            Write-Host "Invalid selection." -ForegroundColor Red
                        }
                        
                        Read-Host "Press Enter to continue"
                        
                    } catch {
                        Write-Host "Error during tweak application: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host "Attempting to discard changes and unmount..." -ForegroundColor Yellow
                        try {
                            dism /Unmount-Wim /MountDir:$Global:MountPath /Discard /CheckIntegrity
                        } catch {
                            Write-Host "Manual cleanup may be required." -ForegroundColor Red
                        }
                        Read-Host "Press Enter to continue"
                    } finally {
                        $Global:CurrentOperation = "None"
                        $Global:CleanupRequired = $false
                    }
                } else {
                    Write-Host "No tweaks selected." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                }
            }
            "20" { return }
            default {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($true)
}

function Apply-TweaksToMountedImage {
    param(
        [string]$IndexNumber,
        [hashtable]$TweakStates
    )
    
    Write-Host "Applying selected tweaks to index ${IndexNumber}..." -ForegroundColor Yellow
    
    $successCount = 0
    $failCount = 0
    
    # Registry hive paths
    $softwareHive = "$Global:MountPath\Windows\System32\config\SOFTWARE"
    $systemHive = "$Global:MountPath\Windows\System32\config\SYSTEM"
    $defaultUserHive = "$Global:MountPath\Users\Default\NTUSER.DAT"
    
    $hivesLoaded = @{}
    
    try {
        # Load registry hives
        if (Test-Path $softwareHive) {
            $tempSoftware = "TEMP_SOFTWARE_$IndexNumber"
            $loadResult = reg load "HKLM\$tempSoftware" $softwareHive 2>&1
            if ($LASTEXITCODE -eq 0) {
                $hivesLoaded[$tempSoftware] = $true
                $softwareLoaded = $true
            } else {
                Write-Host "⚠ Failed to load SOFTWARE hive" -ForegroundColor Yellow
                $softwareLoaded = $false
            }
        } else {
            $softwareLoaded = $false
        }
        
        if (Test-Path $systemHive) {
            $tempSystem = "TEMP_SYSTEM_$IndexNumber"
            $loadResult = reg load "HKLM\$tempSystem" $systemHive 2>&1
            if ($LASTEXITCODE -eq 0) {
                $hivesLoaded[$tempSystem] = $true
                $systemLoaded = $true
            } else {
                Write-Host "⚠ Failed to load SYSTEM hive" -ForegroundColor Yellow
                $systemLoaded = $false
            }
        } else {
            $systemLoaded = $false
        }
        
        if (Test-Path $defaultUserHive) {
            $tempNTUser = "TEMP_NTUSER_$IndexNumber"
            $loadResult = reg load "HKLM\$tempNTUser" $defaultUserHive 2>&1
            if ($LASTEXITCODE -eq 0) {
                $hivesLoaded[$tempNTUser] = $true
                $defaultUserLoaded = $true
            } else {
                Write-Host "⚠ Failed to load NTUSER hive" -ForegroundColor Yellow
                $defaultUserLoaded = $false
            }
        } else {
            $defaultUserLoaded = $false
        }
        
        # Apply tweaks based on selection
        if ($TweakStates.DisableSearchIndexing -and $systemLoaded) {
            Write-Host "Applying: Disable Windows Search Indexing..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\ControlSet001\Services\WSearch" /v "Start" /t REG_DWORD /d 4 /f 2>&1 | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "✓ Windows Search Indexing disabled" -ForegroundColor Green
                    $successCount++
                } else {
                    throw "Registry operation failed"
                }
            } catch {
                Write-Host "✗ Failed to disable search indexing" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.DisableTelemetry -and $softwareLoaded) {
            Write-Host "Applying: Disable Telemetry/Data Collection..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Policies\Microsoft\Windows\DataCollection" /v "AllowTelemetry" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Microsoft\Windows\CurrentVersion\Policies\DataCollection" /v "AllowTelemetry" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Policies\Microsoft\Windows\AppCompat" /v "AITEnable" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Policies\Microsoft\Windows\AppCompat" /v "DisableInventory" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ Telemetry/Data Collection disabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to disable telemetry" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.ClassicContextMenu -and $defaultUserLoaded) {
            Write-Host "Applying: Classic Context Menu (Windows 11)..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_NTUSER_$IndexNumber\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /ve /t REG_SZ /d "" /f 2>&1 | Out-Null
                Write-Host "✓ Classic Context Menu enabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to enable classic context menu" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.DisableAutoRun -and $softwareLoaded) {
            Write-Host "Applying: Disable AutoRun/AutoPlay..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "NoDriveTypeAutoRun" /t REG_DWORD /d 255 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "NoAutorun" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ AutoRun/AutoPlay disabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to disable AutoRun/AutoPlay" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.EnableLongPaths -and $softwareLoaded) {
            Write-Host "Applying: Enable Win32 Long Paths..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Microsoft\Windows\CurrentVersion\Policies\System" /v "EnableLongPathsWin32" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ Win32 Long Paths enabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to enable long paths" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.BypassTPM -and $systemLoaded) {
            Write-Host "Applying: Bypass TPM Requirement..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\Setup\LabConfig" /v "BypassTPMCheck" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ TPM Requirement bypassed" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to bypass TPM requirement" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.BypassRAM -and $systemLoaded) {
            Write-Host "Applying: Bypass RAM Requirement..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\Setup\LabConfig" /v "BypassRAMCheck" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ RAM Requirement bypassed" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to bypass RAM requirement" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.BypassSecureBoot -and $systemLoaded) {
            Write-Host "Applying: Bypass Secure Boot Requirement..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\Setup\LabConfig" /v "BypassSecureBootCheck" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ Secure Boot Requirement bypassed" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to bypass Secure Boot requirement" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.BypassCPU -and $systemLoaded) {
            Write-Host "Applying: Bypass CPU Requirement..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\Setup\LabConfig" /v "BypassCPUCheck" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ CPU Requirement bypassed" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to bypass CPU requirement" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.DisableMSAccount -and $softwareLoaded) {
            Write-Host "Applying: Disable Microsoft Account Requirement..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Microsoft\Windows\CurrentVersion\OOBE" /v "BypassNRO" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Microsoft\PolicyManager\current\device\Accounts" /v "AllowMicrosoftAccountSignInAssistant" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                Write-Host "✓ Microsoft Account Requirement disabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to disable Microsoft Account requirement" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.DisableWeakSSLTLS -and $systemLoaded) {
            Write-Host "Applying: Disable Weak SSL/TLS Protocols..." -ForegroundColor Cyan
            try {
                
                # Create parent protocol keys first
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols" /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0" /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0" /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0" /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1" /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2" /f 2>&1 | Out-Null
                
                # Disable SSL 2.0
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0\Client" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0\Client" /v "DisabledByDefault" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0\Server" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                
                # Disable SSL 3.0
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Client" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                
                # Disable TLS 1.0
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                
                # Disable TLS 1.1
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client" /v "DisabledByDefault" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server" /v "DisabledByDefault" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                
                # Enable TLS 1.2 (ensure it's enabled)
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" /v "Enabled" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" /v "DisabledByDefault" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server" /v "Enabled" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server" /v "DisabledByDefault" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                
                Write-Host "✓ Weak SSL/TLS Protocols disabled, TLS 1.2 enabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to configure SSL/TLS protocols" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.EnableSMBSigning -and $systemLoaded) {
            Write-Host "Applying: Enable SMB Security Signing..." -ForegroundColor Cyan
            try {
                # Enable SMB client signing
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Services\LanManWorkstation\Parameters" /v "RequireSecuritySignature" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                
                # Enable SMB server signing
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Services\LanManServer\Parameters" /v "RequireSecuritySignature" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                
                Write-Host "✓ SMB Security Signing enabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to enable SMB security signing" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.DisableIE11 -and $softwareLoaded) {
            Write-Host "Applying: Disable Internet Explorer 11..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Microsoft\Internet Explorer\Setup\11.0" /v "DoNotAllowIE11" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ Internet Explorer 11 disabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to disable Internet Explorer 11" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.EnableCertPaddingCheck -and $softwareLoaded) {
            Write-Host "Applying: Enable Certificate Padding Check..." -ForegroundColor Cyan
            try {
                # Enable for 64-bit
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Microsoft\Cryptography\Wintrust\Config" /v "EnableCertPaddingCheck" /t REG_SZ /d "1" /f 2>&1 | Out-Null
                
                # Enable for 32-bit on 64-bit systems
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Wow6432Node\Microsoft\Cryptography\Wintrust\Config" /v "EnableCertPaddingCheck" /t REG_SZ /d "1" /f 2>&1 | Out-Null
                
                Write-Host "✓ Certificate Padding Check enabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to enable certificate padding check" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.RestrictNetworkLocation -and $softwareLoaded) {
            Write-Host "Applying: Restrict Network Location Changes..." -ForegroundColor Cyan
            try {
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Policies\Microsoft\Windows\Network Connections" /v "NC_StdDomainUserSetLocation" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ Network Location Changes restricted" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to restrict network location changes" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.PreventDeviceEncryption -and $systemLoaded) {
            Write-Host "Applying: Enable Prevent Device Encryption (Disable BitLocker)..." -ForegroundColor Cyan
            try {
                # Microsoft's recommended method to prevent automatic device encryption
                reg add "HKLM\TEMP_SYSTEM_$IndexNumber\CurrentControlSet\Control\BitLocker" /v "PreventDeviceEncryption" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                
                Write-Host "✓ Device Encryption Prevention enabled (BitLocker auto-start disabled)" -ForegroundColor Green
                Write-Host "  Note: Not recommended on devices with Windows Recall feature" -ForegroundColor Yellow
                Write-Host "  Note: Ensure compliance with Windows 11 licensing requirements" -ForegroundColor Yellow
                $successCount++
            } catch {
                Write-Host "✗ Failed to enable device encryption prevention" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.EnableADRecoveryBackup -and $softwareLoaded) {
            Write-Host "Applying: Enable AD Recovery Key Backup..." -ForegroundColor Cyan
            try {
                # Enable Active Directory backup of BitLocker recovery keys
                reg add "HKLM\TEMP_SOFTWARE_$IndexNumber\Policies\Microsoft\FVE" /v "OSRequireActiveDirectoryBackup" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                
                Write-Host "✓ AD Recovery Key Backup enabled" -ForegroundColor Green
                Write-Host "  Note: Requires proper Active Directory configuration" -ForegroundColor Yellow
                $successCount++
            } catch {
                Write-Host "✗ Failed to enable AD recovery key backup" -ForegroundColor Red
                $failCount++
            }
        }
        
    } finally {
        # Unload all loaded hives
        foreach ($tempName in $hivesLoaded.Keys) {
            try {
                reg unload "HKLM\$tempName" 2>&1 | Out-Null
            } catch {}
        }
    }
    
    Write-Host "Tweaks application completed for index ${IndexNumber}: $successCount successful, $failCount failed" -ForegroundColor Cyan
    
    return @{
        Success = $successCount
        Failed = $failCount
    }
}

function Apply-TweaksToCurrentSystem {
    param(
        [hashtable]$TweakStates
    )
    
    Write-Host "Applying selected tweaks to current system..." -ForegroundColor Yellow
    
    $successCount = 0
    $failCount = 0
    $requiresReboot = $false
    
    try {
        # Apply tweaks based on selection to live registry
        if ($TweakStates.DisableSearchIndexing) {
            Write-Host "Applying: Disable Windows Search Indexing..." -ForegroundColor Cyan
            try {
                # Stop Windows Search service first
                Stop-Service -Name "WSearch" -Force -ErrorAction SilentlyContinue
                reg add "HKLM\SYSTEM\CurrentControlSet\Services\WSearch" /v "Start" /t REG_DWORD /d 4 /f 2>&1 | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "✓ Windows Search Indexing disabled" -ForegroundColor Green
                    $successCount++
                    $requiresReboot = $true
                } else {
                    throw "Registry operation failed"
                }
            } catch {
                Write-Host "✗ Failed to disable search indexing: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.DisableTelemetry) {
            Write-Host "Applying: Disable Telemetry/Data Collection..." -ForegroundColor Cyan
            try {
                reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection" /v "AllowTelemetry" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" /v "AllowTelemetry" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat" /v "AITEnable" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat" /v "DisableInventory" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ Telemetry/Data Collection disabled" -ForegroundColor Green
                $successCount++
                $requiresReboot = $true
            } catch {
                Write-Host "✗ Failed to disable telemetry: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.ClassicContextMenu) {
            Write-Host "Applying: Classic Context Menu (Windows 11)..." -ForegroundColor Cyan
            try {
                # Apply to current user
                reg add "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /ve /t REG_SZ /d "" /f 2>&1 | Out-Null
                # Apply to default user (for new users)
                reg add "HKU\.DEFAULT\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /ve /t REG_SZ /d "" /f 2>&1 | Out-Null
                Write-Host "✓ Classic Context Menu enabled" -ForegroundColor Green
                $successCount++
                Write-Host "  Note: You may need to restart Windows Explorer to see changes" -ForegroundColor Gray
            } catch {
                Write-Host "✗ Failed to enable classic context menu: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.DisableAutoRun) {
            Write-Host "Applying: Disable AutoRun/AutoPlay..." -ForegroundColor Cyan
            try {
                reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "NoDriveTypeAutoRun" /t REG_DWORD /d 255 /f 2>&1 | Out-Null
                reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "NoAutorun" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                reg add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "NoDriveTypeAutoRun" /t REG_DWORD /d 255 /f 2>&1 | Out-Null
                Write-Host "✓ AutoRun/AutoPlay disabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to disable AutoRun/AutoPlay: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.EnableLongPaths) {
            Write-Host "Applying: Enable Win32 Long Paths..." -ForegroundColor Cyan
            try {
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\FileSystem" /v "LongPathsEnabled" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ Win32 Long Paths enabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to enable long paths: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        # Note: Bypass tweaks (TPM, RAM, Secure Boot, CPU) are not applicable to live system
        if ($TweakStates.BypassTPM -or $TweakStates.BypassRAM -or $TweakStates.BypassSecureBoot -or $TweakStates.BypassCPU) {
            Write-Host "⚠ Bypass tweaks (TPM/RAM/SecureBoot/CPU) are only applicable to WIM images, not live systems" -ForegroundColor Yellow
        }
        
        # Note: MS Account bypass is not applicable to live system
        if ($TweakStates.DisableMSAccount) {
            Write-Host "⚠ Microsoft Account bypass is only applicable to WIM images, not live systems" -ForegroundColor Yellow
        }
        
        if ($TweakStates.DisableWeakSSLTLS) {
            Write-Host "Applying: Disable Weak SSL/TLS Protocols..." -ForegroundColor Cyan
            try {
              
                # Create parent protocol keys first
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols" /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0" /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0" /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0" /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1" /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2" /f 2>&1 | Out-Null
                
                # Disable SSL 2.0
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0\Client" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0\Client" /v "DisabledByDefault" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0\Server" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                
                # Disable SSL 3.0
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Client" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                
                # Disable TLS 1.0
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                
                # Disable TLS 1.1
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client" /v "DisabledByDefault" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server" /v "Enabled" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server" /v "DisabledByDefault" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                
                # Enable TLS 1.2 (ensure it's enabled)
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" /v "Enabled" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" /v "DisabledByDefault" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server" /v "Enabled" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server" /v "DisabledByDefault" /t REG_DWORD /d 0 /f 2>&1 | Out-Null
                
                Write-Host "✓ Weak SSL/TLS Protocols disabled, TLS 1.2 enabled" -ForegroundColor Green
                $successCount++
                $requiresReboot = $true
            } catch {
                Write-Host "✗ Failed to configure SSL/TLS protocols: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.EnableSMBSigning) {
            Write-Host "Applying: Enable SMB Security Signing..." -ForegroundColor Cyan
            try {
                # Enable SMB client signing
                reg add "HKLM\SYSTEM\CurrentControlSet\Services\LanManWorkstation\Parameters" /v "RequireSecuritySignature" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                
                # Enable SMB server signing
                reg add "HKLM\SYSTEM\CurrentControlSet\Services\LanManServer\Parameters" /v "RequireSecuritySignature" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                
                Write-Host "✓ SMB Security Signing enabled" -ForegroundColor Green
                $successCount++
                $requiresReboot = $true
            } catch {
                Write-Host "✗ Failed to enable SMB security signing: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.DisableIE11) {
            Write-Host "Applying: Disable Internet Explorer 11..." -ForegroundColor Cyan
            try {
                reg add "HKLM\SOFTWARE\Microsoft\Internet Explorer\Setup\11.0" /v "DoNotAllowIE11" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ Internet Explorer 11 disabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to disable Internet Explorer 11: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.EnableCertPaddingCheck) {
            Write-Host "Applying: Enable Certificate Padding Check..." -ForegroundColor Cyan
            try {
                # Enable for 64-bit
                reg add "HKLM\SOFTWARE\Microsoft\Cryptography\Wintrust\Config" /v "EnableCertPaddingCheck" /t REG_SZ /d "1" /f 2>&1 | Out-Null
                
                # Enable for 32-bit on 64-bit systems
                reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\Cryptography\Wintrust\Config" /v "EnableCertPaddingCheck" /t REG_SZ /d "1" /f 2>&1 | Out-Null
                
                Write-Host "✓ Certificate Padding Check enabled" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to enable certificate padding check: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.RestrictNetworkLocation) {
            Write-Host "Applying: Restrict Network Location Changes..." -ForegroundColor Cyan
            try {
                reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Network Connections" /v "NC_StdDomainUserSetLocation" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                Write-Host "✓ Network Location Changes restricted" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "✗ Failed to restrict network location changes: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.PreventDeviceEncryption) {
            Write-Host "Applying: Enable Prevent Device Encryption (Disable BitLocker)..." -ForegroundColor Cyan
            try {
                # Microsoft's recommended method to prevent automatic device encryption
                reg add "HKLM\SYSTEM\CurrentControlSet\Control\BitLocker" /v "PreventDeviceEncryption" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                
                Write-Host "✓ Device Encryption Prevention enabled (BitLocker auto-start disabled)" -ForegroundColor Green
                Write-Host "  Note: Not recommended on devices with Windows Recall feature" -ForegroundColor Yellow
                Write-Host "  Note: Ensure compliance with Windows 11 licensing requirements" -ForegroundColor Yellow
                $successCount++
            } catch {
                Write-Host "✗ Failed to enable device encryption prevention: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
        if ($TweakStates.EnableADRecoveryBackup) {
            Write-Host "Applying: Enable AD Recovery Key Backup..." -ForegroundColor Cyan
            try {
                # Enable Active Directory backup of BitLocker recovery keys
                reg add "HKLM\SOFTWARE\Policies\Microsoft\FVE" /v "OSRequireActiveDirectoryBackup" /t REG_DWORD /d 1 /f 2>&1 | Out-Null
                
                Write-Host "✓ AD Recovery Key Backup enabled" -ForegroundColor Green
                Write-Host "  Note: Requires proper Active Directory configuration" -ForegroundColor Yellow
                $successCount++
            } catch {
                Write-Host "✗ Failed to enable AD recovery key backup: $($_.Exception.Message)" -ForegroundColor Red
                $failCount++
            }
        }
        
    } catch {
        Write-Host "✗ Unexpected error during tweak application: $($_.Exception.Message)" -ForegroundColor Red
        $failCount++
    }
    
    Write-Host "Tweaks application completed for current system: $successCount successful, $failCount failed" -ForegroundColor Cyan
    
    return @{
        Success = $successCount
        Failed = $failCount
        RequiresReboot = $requiresReboot
    }
}

function Main {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "This script requires administrator privileges. Please run PowerShell as administrator." -ForegroundColor Red
        exit 1
    }

    do {
        Show-Menu
        $choice = Read-Host
        
        switch ($choice) {
            "1" { Manage-Configuration }
            "2" { Manage-Drivers }
            "3" { 
                if (-not (Test-RequiredPaths)) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } elseif (-not (Test-MountPathEmpty)) {
                    Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Manage-PackagesAndFeatures
                }
            }
            "4" { 
                if (-not (Test-RequiredPaths)) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } elseif (-not (Test-MountPathEmpty)) {
                    Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Manage-Indexes
                }
            }
            "5" { 
                if (-not (Test-RequiredPaths)) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } elseif (-not (Test-MountPathEmpty)) {
                    Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Manage-Editions
                }
            }
            "6" { 
                if (-not (Test-RequiredPaths)) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } elseif (-not (Test-MountPathEmpty)) {
                    Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Manage-Conversions
                }
            }
            "7" { 
                if (-not (Test-RequiredPaths)) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Manage-ISOCreation
                }
            }
            "8" { 
                if (-not (Test-RequiredPaths)) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Get-WimInfo
                }
            }
            "9" { 
                if (-not (Test-RequiredPaths)) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } elseif (-not (Test-MountPathEmpty)) {
                    Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Manage-DesktopFiles
                }
            }
            "10" { 
                if (-not (Test-RequiredPaths)) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } elseif (-not (Test-MountPathEmpty)) {
                    Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Manage-Registry
                }
            }
            "11" { 
                if (-not (Test-RequiredPaths)) {
                    Write-Host "Please set all required paths first (ISO Path, Mount Path)." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } elseif (-not (Test-MountPathEmpty)) {
                    Write-Host "Mount path is not empty! Please use the cleanup option first." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                } else {
                    Mount-IndexOnly
                }
            }
            "12" { Invoke-ManualCleanup }
            "13" { 
                Write-Host "Exiting..." -ForegroundColor Yellow
                Invoke-Cleanup
                exit 
            }
            default { 
                Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($true)
}

Main