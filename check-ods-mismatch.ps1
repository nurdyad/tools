<#
.SYNOPSIS
Validates consistency between practice work-item directories and their
corresponding work-items.json configuration.

.DESCRIPTION
This script performs two types of validation:

1. Practice-specific directories
   - Ensures directory names follow the format:
       [Practice Name] (ODS Code)
   - Confirms the ODS code in the directory name matches payload.ods_code
   - Confirms the practice name matches payload.docman_practice_display_name
     (case-insensitive, title case vs ALL CAPS supported)

2. Practice Count
   - Treated as a reference container (not a practice directory)
   - Each entry is validated to ensure it maps to an existing practice directory
   - Missing mappings are reported as warnings (not errors)

This script is non-destructive and intended as a lightweight guardrail to
detect configuration drift during onboarding or manual updates.
#>

# =========================
# Configuration
# =========================

# This block allows Python to pass the folder path to PowerShell
param (
    [Parameter(Mandatory=$false)]
    [string]$BasePath = "C:\rpa\postie\postie-bots\devdata\work-items-in"
)

$PracticeCountDir = "Practice Count"

# Non-practice folders that should be excluded from validation
$ExcludedFolders = @(
    "Platform upload",
    "Scanner",
    "Verification testing",
    "Practice Count"
)

# =========================
# Helper Functions
# =========================

function Format-PracticeName {
    param ([string]$Name)
    # Remove ODS suffix if present and normalise for comparison
    return ($Name -replace "\s*\([A-Z0-9]+\)$", "").Trim().ToUpper()
}

function GetOdsFromFolder {
    param ([string]$FolderName)

    # This regex ensures we only extract exactly 6 alphanumeric characters 
    # inside parentheses at the very END of the folder name.
    if ($FolderName -match "\(([A-Z0-9]{6})\)$") {
        return $Matches[1]
    }
    return $null
}

# =========================
# Load practice directories
# =========================

if (-not (Test-Path $BasePath)) {
    Write-Warning "Target path does not exist: $BasePath"
    return
}

$PracticeDirs = Get-ChildItem $BasePath -Directory |
    Where-Object { $ExcludedFolders -notcontains $_.Name }

# Map of normalised practice name → folder name
$PracticeDirectoryMap = @{}

foreach ($Dir in $PracticeDirs) {
    $PracticeDirectoryMap[(Format-PracticeName $Dir.Name)] = $Dir.Name
}

# =========================
# 1️⃣ Validate practice-specific directories
# =========================

foreach ($Dir in $PracticeDirs) {
    $WorkItemsPath = Join-Path $Dir.FullName "work-items.json"

    if (-not (Test-Path $WorkItemsPath)) {
        Write-Warning "Missing work-items.json in folder: $($Dir.Name)"
        continue
    }

    try {
        $Json = Get-Content $WorkItemsPath -Raw | ConvertFrom-Json
        $Payload = $Json[0].payload
    }
    catch {
        Write-Warning "Invalid JSON in $($Dir.Name)\work-items.json"
        continue
    }

    # --- ODS validation ---
    $FolderOds   = GetOdsFromFolder $Dir.Name
    $ExpectedOds = $Payload.ods_code

    if (-not $FolderOds) {
        Write-Warning "Folder missing ODS code: $($Dir.Name)"
        continue
    }

    if ($FolderOds -ne $ExpectedOds) {
        Write-Warning "ODS mismatch in '$($Dir.Name)': folder=$FolderOds json=$ExpectedOds"
    }

    # --- Name validation ---
    if ($Payload.docman_practice_display_name) {
        $FolderNameNormalized = Format-PracticeName $Dir.Name
        $DocmanNameNormalized = Format-PracticeName $Payload.docman_practice_display_name

        if ($FolderNameNormalized -ne $DocmanNameNormalized) {
            Write-Warning "Name mismatch in '$($Dir.Name)': folder='$FolderNameNormalized' docman='$DocmanNameNormalized'"
        }
    }
}

# =========================
# 2️⃣ Validate Practice Count entries
# =========================

# Corrected for PowerShell version compatibility
$PracticeCountPath = Join-Path (Join-Path $BasePath $PracticeCountDir) "work-items.json"

if (Test-Path $PracticeCountPath) {
    try {
        # FIX: Using $PracticeCountPath to avoid null errors
        $PracticeCountJson = Get-Content $PracticeCountPath -Raw | ConvertFrom-Json
    }
    catch {
        Write-Warning "Invalid JSON in Practice Count work-items.json"
        return
    }

    foreach ($Entry in $PracticeCountJson) {
        if (-not $Entry.payload.docman_practice_display_name) { continue }

        $ExpectedName = Format-PracticeName $Entry.payload.docman_practice_display_name

        if (-not $PracticeDirectoryMap.ContainsKey($ExpectedName)) {
            Write-Warning "Practice Count entry not found as directory: $($Entry.payload.docman_practice_display_name)"
        }
    }
} else {
    Write-Warning "Practice Count work-items.json not found at $PracticeCountPath"
}

Write-Host "ODS and practice directory validation completed."