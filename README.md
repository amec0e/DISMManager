# DISMManager
A tool for helping to create custom recovery ISOs using windows DISM

**IMPORTANT:** This script includes Generic RTM and KMS Keys these are used to set the product key when upgrading an edition of windows. **THESE KEYS WILL NOT ACTIVATE YOUR MACHINE** and the keys provided can be found publicly available, they are to help with installations and upgrades only.

##### Script Structure:
```
MAIN MENU
├── 1. Configuration Settings
│   ├── 1. Set Extracted ISO Path
│   ├── 2. Set Mount Path  
│   ├── 3. Set ISO Output Path
│   ├── 4. Set Driver Path
│   ├── 5. Set Package Path
│   ├── 6. Set Conversion Backup Path
│   ├── 7. Setup Wimlib (Advanced Features)
│   │   ├── 1. Download Wimlib (opens browser)
│   │   ├── 2. Set Wimlib Path (after download)
│   │   ├── 3. Test Wimlib Installation
│   │   ├── 4. Clear Wimlib Path
│   │   └── 5. Return to Configuration Menu
│   ├── 8. Toggle Force Unsigned Drivers
│   └── 9. Return to Main Menu
│
├── 2. Manage Drivers (Add/Remove)
│   ├── 1. Add Drivers
│   │   ├── • Add to specific index
│   │   └── • Add to all indexes
│   ├── 2. Remove Drivers
│   │   ├── • Remove from specific index
│   │   └── • Remove from all indexes
│   ├── 3. Export Drivers from Current System
│   ├── 4. Export Drivers from Mounted Image
│   └── 5. Return to Main Menu
│
├── 3. Manage Packages and Features
│   ├── 1. Manage for specific index
│   ├── 2. Manage for all indexes
│   └── Per-Index Management Menu:
│       ├── PACKAGE MANAGEMENT:
│       │   ├── 1. View all installed packages
│       │   ├── 2. View specific package info
│       │   ├── 3. Add package(s)
│       │   └── 4. Remove package(s)
│       ├── FEATURE MANAGEMENT:
│       │   ├── 5. View all available features
│       │   ├── 6. View specific feature info
│       │   ├── 7. Enable feature(s)
│       │   └── 8. Disable feature(s)
│       ├── 9. Commit changes and continue
│       └── 10. Discard changes and exit
│
├── 4. Manage Indexes (Add/Remove)
│   ├── 1. View Current Indexes
│   ├── 2. Add Index
│   │   ├── 1. Export from another WIM file
│   │   └── 2. Export from ESD file
│   ├── 3. Remove Index
│   └── 4. Return to Main Menu
│
├── 5. Manage Editions
│   ├── 1. View Current Edition (Mounted Image)
│   ├── 2. View Available Target Editions (Mounted Image)
│   ├── 3. Upgrade Edition (Mounted Image)
│   ├── 4. View Index Metadata (requires wimlib)
│   ├── 5. Update Index Metadata (requires wimlib)
│   ├── 6. Set Product Keys (Generic/RTM)
│   │   ├── 1. Set key for specific index
│   │   ├── 2. Set key for all indexes
│   │   ├── 3. View available RTM keys
│   │   ├── 4. View available KMS Client keys (Official)
│   │   └── 5. Return to Edition Management
│   └── 7. Return to Main Menu
│
├── 6. File Conversion (WIM/ESD)
│   ├── 1. Convert to WIM
│   ├── 2. Convert to ESD
│   ├── 3. Recompress current format
│   └── 4. Return to Main Menu
│
├── 7. Create ISO (UEFI/BIOS/Hybrid)
│   ├── 1. Create UEFI-only ISO
│   ├── 2. Create BIOS-only ISO
│   ├── 3. Create Hybrid ISO (UEFI + BIOS)
│   ├── 4. Detect and Recommend Best Option
│   └── 5. Return to Main Menu
│
├── 8. View WIM/ESD Information
│
├── 9. Mount Index
│   ├── • Mount Read-Write (for making changes)
│   ├── • Mount Read-Only (for viewing only)
│   └── • Cancel
│
├── 10. Cleanup/Unmount
│
└── 11. Exit
```