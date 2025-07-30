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
│   ├── 6. Set Backup Path
│   ├── 7. Set Desktop Files Path
│   ├── 8. Setup Wimlib (Advanced Features)
│   │   ├── 1. Download Wimlib (opens browser)
│   │   ├── 2. Set Wimlib Path (after download)
│   │   ├── 3. Test Wimlib Installation
│   │   ├── 4. Clear Wimlib Path
│   │   └── 5. Return to Configuration Menu
│   ├── 9. Toggle Force Unsigned Drivers
│   └── 10. Return to Main Menu
│
├── 2. Manage Drivers (Add/Remove/Export)
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
│   ├── 3. Manage Current Live System ⭐ (NEW)
│   │   ├── SYSTEM PACKAGE MANAGEMENT (.cab/.msu):
│   │   │   ├── 1. View all installed system packages
│   │   │   ├── 2. View specific system package info
│   │   │   ├── 3. Add system package(s)
│   │   │   └── 4. Remove system package(s)
│   │   ├── APPLICATION PACKAGE MANAGEMENT (.appx/.appxbundle):
│   │   │   ├── 5. View all provisioned application packages
│   │   │   ├── 6. Add application package(s)
│   │   │   └── 7. Remove application package(s)
│   │   ├── FEATURE MANAGEMENT:
│   │   │   ├── 8. View all available features
│   │   │   ├── 9. View specific feature info
│   │   │   ├── 10. Enable feature(s)
│   │   │   └── 11. Disable feature(s)
│   │   └── 12. Return to Main Menu
│   └── 4. Return to Main Menu
│   │
│   └── Per-Index Management Menu (Options 1 & 2):
│       ├── SYSTEM PACKAGE MANAGEMENT (.cab/.msu):
│       │   ├── 1. View all installed system packages
│       │   ├── 2. View specific system package info
│       │   ├── 3. Add system package(s)
│       │   └── 4. Remove system package(s)
│       ├── APPLICATION PACKAGE MANAGEMENT (.appx/.appxbundle):
│       │   ├── 5. View all provisioned application packages
│       │   ├── 6. Add application package(s)
│       │   └── 7. Remove application package(s)
│       ├── FEATURE MANAGEMENT:
│       │   ├── 8. View all available features
│       │   ├── 9. View specific feature info
│       │   ├── 10. Enable feature(s)
│       │   └── 11. Disable feature(s)
│       ├── 12. Commit changes and continue
│       └── 13. Discard changes and exit
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
├── 9. Manage Desktop Files
│   ├── 1. Add Desktop Files to Images
│   │   ├── • Select files to add
│   │   ├── • Add to specific index
│   │   └── • Add to all indexes
│   ├── 2. Remove Desktop Files from Images
│   └── 3. Return to Main Menu
│
├── 10. Manage Registry
│   ├── 1. Apply Common Tweaks
│   │   ├── Tweak Selection Interface (19 toggleable options):
│   │   │   ├── 1. [✓/✗] Disable Windows Search Indexing
│   │   │   ├── 2. [✓/✗] Disable Telemetry/Data Collection
│   │   │   ├── 3. [✓/✗] Classic Context Menu (Windows 11)
│   │   │   ├── 4. [✓/✗] Disable AutoRun/AutoPlay
│   │   │   ├── 5. [✓/✗] Enable Win32 Long Paths
│   │   │   ├── 6. [✓/✗] Bypass TPM Requirements (Image only)
│   │   │   ├── 7. [✓/✗] Bypass RAM Requirements (Image only)
│   │   │   ├── 8. [✓/✗] Bypass Secure Boot (Image only)
│   │   │   ├── 9. [✓/✗] Bypass CPU Requirements (Image only)
│   │   │   ├── 10. [✓/✗] Disable MS Account Requirements (Image only)
│   │   │   ├── 11. [✓/✗] Disable Weak SSL/TLS Protocols
│   │   │   ├── 12. [✓/✗] Enable SMB Signing
│   │   │   ├── 13. [✓/✗] Disable Internet Explorer 11
│   │   │   ├── 14. [✓/✗] Enable Certificate Padding Check
│   │   │   ├── 15. [✓/✗] Restrict Network Location Changes
│   │   │   ├── 16. [✓/✗] Prevent Device Encryption (Disable BitLocker)
│   │   │   ├── 17. [✓/✗] Enable AD Recovery Key Backup
│   │   │   ├── 18. Toggle All (Select/Deselect All)
│   │   │   └── 19. Apply Selected Tweaks
│   │   │       ├── 1. Apply to current system (live)
│   │   │       ├── 2. Apply to specific index
│   │   │       ├── 3. Apply to all indexes
│   │   │       └── 4. Cancel
│   │   └── 20. Return to Registry Management
│   └── 2. Return to Main Menu
│
├── 11. Mount Index
│   ├── • Mount Read-Write (for making changes)
│   ├── • Mount Read-Only (for viewing only)
│   └── • Cancel
│
├── 12. Cleanup/Unmount
│
└── 13. Exit
```