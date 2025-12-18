# Migration Validation Script for SharePoint 2013 to SharePoint SE
# Korrigierte Version

param(
    [Parameter(Mandatory=$true)][string]$SourceUrl,
    [Parameter(Mandatory=$true)][string]$TargetUrl,
    [string]$OutputPath = ".\MigrationValidation_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [switch]$UseCurrentCredentials = $true # Schalter für Windows Auth (Standard bei OnPrem)
)

# Prüfung auf Modul (Warnung bei SP2013)
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    Write-Warning "Das Modul 'PnP.PowerShell' fehlt. Bitte installieren (für SP SE)."
    # Hinweis: Für SP2013 benötigen Sie ggf. das Legacy Modul, was zu Konflikten führen kann.
    # Dieses Skript geht davon aus, dass die Verbindung trotzdem hergestellt werden kann.
}
Import-Module PnP.PowerShell -ErrorAction SilentlyContinue

$results = @()
$validationErrors = @()

# Cleanup URLs (Slash am Ende entfernen)
$SourceUrl = $SourceUrl.TrimEnd('/')
$TargetUrl = $TargetUrl.TrimEnd('/')

# ===== VERBINDUNG ZU SHAREPOINT =====
Write-Host "Verbinde mit SharePoint Instanzen..." -ForegroundColor Cyan

# Funktion für Connect (On-Premises Optimierung)
function Connect-SPOnPrem {
    param($Url)
    try {
        if ($UseCurrentCredentials) {
            # Nutzt den eingeloggten Windows User (Standard für OnPrem zu OnPrem)
            return Connect-PnPOnline -Url $Url -CurrentCredentials -ReturnConnection -ErrorAction Stop
        } else {
            # Fragt nach Credentials (Legacy Auth)
            $creds = Get-Credential
            return Connect-PnPOnline -Url $Url -Credentials $creds -ReturnConnection -ErrorAction Stop
        }
    } catch {
        throw $_
    }
}

try {
    $sourceContext = Connect-SPOnPrem -Url $SourceUrl
    # ServerRelativeUrl des Source Root Webs holen für spätere Pfad-Berechnungen
    $sourceRootWeb = Get-PnPWeb -Connection $sourceContext
    $sourceRootRelUrl = $sourceRootWeb.ServerRelativeUrl
    Write-Host "✓ Mit Quelle verbunden: $SourceUrl ($($sourceRootWeb.Title))" -ForegroundColor Green
} catch {
    Write-Host "✗ Fehler bei Verbindung zur Quelle ($SourceUrl): $_" -ForegroundColor Red
    Write-Host "  Hinweis: Für SP2013 kann PnP.PowerShell Probleme machen. Stellen Sie sicher, dass CSOM aktiviert ist." -ForegroundColor Gray
    exit 1
}

try {
    $targetContext = Connect-SPOnPrem -Url $TargetUrl
    # ServerRelativeUrl des Target Root Webs holen
    $targetRootWeb = Get-PnPWeb -Connection $targetContext
    $targetRootRelUrl = $targetRootWeb.ServerRelativeUrl
    Write-Host "✓ Mit Ziel verbunden: $TargetUrl ($($targetRootWeb.Title))" -ForegroundColor Green
} catch {
    Write-Host "✗ Fehler bei Verbindung zum Ziel: $_" -ForegroundColor Red
    exit 1
}

# ===== FUNKTIONEN (Optimiert) =====

function Get-FilesInLibrary {
    param(
        [object]$Context,
        [string]$LibraryTitle
    )
    
    try {
        # Performance-Optimierung: ItemCount Eigenschaft der Liste zuerst prüfen
        # Das Iterieren über Items ist sehr langsam bei großen Listen.
        $list = Get-PnPList -Identity $LibraryTitle -Connection $Context -Includes ItemCount, RootFolder.Folders
        
        # Grobe Zählung über ItemCount (schneller)
        # Wenn exakte Trennung File/Folder nötig ist, muss iteriert werden, was bei >5000 Items dauert.
        # Hier nutzen wir eine CAML Query für bessere Performance bei großen Listen
        
        $filesQuery = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>0</Value></Eq></Where></Query></View>"
        $foldersQuery = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>1</Value></Eq></Where></Query></View>"
        
        $files = Get-PnPListItem -List $LibraryTitle -Connection $Context -Query $filesQuery -PageSize 2000
        $folders = Get-PnPListItem -List $LibraryTitle -Connection $Context -Query $foldersQuery -PageSize 2000
        
        return @{
            FileCount = $files.Count
            FolderCount = $folders.Count
        }
    } catch {
        Write-Host "  ⚠ Fehler beim Lesen der Bibliothek '$LibraryTitle': $_" -ForegroundColor Red
        return @{ FileCount = -1; FolderCount = -1 }
    }
}

function Compare-Permissions {
    param($SourceContext, $TargetContext, $ListName)
    # (Logik beibehalten, da okay für Übersicht)
    try {
        $sourcePerms = Get-PnPRoleAssignment -List $ListName -Connection $SourceContext
        $targetPerms = Get-PnPRoleAssignment -List $ListName -Connection $TargetContext
        
        # Vergleich auf Basis von LoginName kann tricky sein bei Domain-Wechseln. 
        # Wir entfernen Claims-Präfixe für den Vergleich (z.B. i:0#.w|)
        $srcUsers = $sourcePerms.Member.LoginName | ForEach-Object { $_ -replace "^.*\|", "" }
        $tgtUsers = $targetPerms.Member.LoginName | ForEach-Object { $_ -replace "^.*\|", "" }
        
        $missing = $srcUsers | Where-Object { $_ -notin $tgtUsers }
        
        if ($missing) { return "Fehlend: $($missing -join ', ')" }
        return $null
    } catch {
        return "Fehler beim Perm-Check"
    }
}

# ===== LOGIK: ROOT BIBLIOTHEKEN =====
Write-Host "`n=== VALIDIERE ROOT-BIBLIOTHEKEN ===" -ForegroundColor Cyan

$sourceLibs = Get-PnPList -Connection $sourceContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden -and $_.Title -ne "Microfeed" }
$targetLibs = Get-PnPList -Connection $targetContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }

foreach ($lib in $sourceLibs) {
    Write-Host -NoNewline "Prüfe $($lib.Title)... "
    
    # Check if exists in Target (Case Insensitive)
    $targetLib = $targetLibs | Where-Object { $_.Title -eq $lib.Title }
    
    if ($targetLib) {
        $sInfo = Get-FilesInLibrary -Context $sourceContext -LibraryTitle $lib.Title
        $tInfo = Get-FilesInLibrary -Context $targetContext -LibraryTitle $targetLib.Title
        
        $status = if ($sInfo.FileCount -eq $tInfo.FileCount) { "OK" } else { "DIFF" }
        $color = if ($status -eq "OK") { "Green" } else { "Red" }
        
        Write-Host "$status (Q: $($sInfo.FileCount), Z: $($tInfo.FileCount))" -ForegroundColor $color
        
        $permRes = Compare-Permissions -SourceContext $sourceContext -TargetContext $targetContext -ListName $lib.Title
        
        $results += [PSCustomObject]@{
            Level = "Root"
            Name = $lib.Title
            Source_Files = $sInfo.FileCount
            Target_Files = $tInfo.FileCount
            Status = $status
            Permissions = $permRes
        }
    } else {
        Write-Host "FEHLT IM ZIEL" -ForegroundColor Red
        $validationErrors += "Root-Lib fehlt: $($lib.Title)"
        $results += [PSCustomObject]@{ Level="Root"; Name=$lib.Title; Status="MISSING"; Source_Files=""; Target_Files="" }
    }
}

# ===== LOGIK: SUBSITES (Rekursiv) =====
Write-Host "`n=== VALIDIERE SUBSITES ===" -ForegroundColor Cyan

# Wir holen alle Webs. Wichtig: Include ServerRelativeUrl
$sourceSubwebs = Get-PnPSubWeb -Connection $sourceContext -Recurse -Includes ServerRelativeUrl, Title

foreach ($web in $sourceSubwebs) {
    # 1. Berechne den relativen Pfad ab dem Source-Root
    # Bsp Source: /sites/team/sub1
    # Bsp Root:   /sites/team
    # Relative:   /sub1
    $relativePath = $web.ServerRelativeUrl.Substring($sourceRootRelUrl.Length)
    
    # 2. Baue die erwartete Ziel-URL
    # Bsp TargetRoot: /sites/teamSE
    # Erwartet:       /sites/teamSE/sub1
    $expectedTargetUrl = $targetRootRelUrl + $relativePath
    
    # Die volle absolute URL für Connect-PnPOnline berechnen
    # Dazu nehmen wir den Host der TargetUrl
    $targetUriObj = [System.Uri]$TargetUrl
    $fullTargetWebUrl = $targetUriObj.Scheme + "://" + $targetUriObj.Authority + $expectedTargetUrl
    
    Write-Host "`nSubsite: $($web.Title)" -ForegroundColor Yellow
    Write-Host "  Pfad: $relativePath" -ForegroundColor Gray
    
    try {
        # Versuche Verbindung zur Subsite im Ziel
        $targetSubContext = Connect-SPOnPrem -Url $fullTargetWebUrl
        Write-Host "  ✓ Subsite im Ziel gefunden" -ForegroundColor Green
        
        # Verbinde zur Source Subsite für Details
        $sourceSubContext = Connect-SPOnPrem -Url $web.Url
        
        # Bibliotheken vergleichen
        $subSourceLibs = Get-PnPList -Connection $sourceSubContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden -and $_.Title -ne "Microfeed" }
        $subTargetLibs = Get-PnPList -Connection $targetSubContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }
        
        foreach ($slib in $subSourceLibs) {
            $tlib = $subTargetLibs | Where-Object { $_.Title -eq $slib.Title }
            
            if ($tlib) {
                $sInfo = Get-FilesInLibrary -Context $sourceSubContext -LibraryTitle $slib.Title
                $tInfo = Get-FilesInLibrary -Context $targetSubContext -LibraryTitle $tlib.Title
                
                $status = if ($sInfo.FileCount -eq $tInfo.FileCount) { "OK" } else { "DIFF" }
                
                $results += [PSCustomObject]@{
                    Level = "Subsite: $($web.Title)"
                    Name = $slib.Title
                    Source_Files = $sInfo.FileCount
                    Target_Files = $tInfo.FileCount
                    Status = $status
                    Permissions = ""
                }
                Write-Host "    Lib '$($slib.Title)': $status (Q:$($sInfo.FileCount)/Z:$($tInfo.FileCount))"
            } else {
                Write-Host "    Lib '$($slib.Title)': FEHLT" -ForegroundColor Red
                $validationErrors += "Subsite '$($web.Title)' - Lib fehlt: $($slib.Title)"
            }
        }
        
    } catch {
        Write-Host "  ✗ Subsite konnte im Ziel nicht verbunden werden (Fehlt vermutlich oder URL anders)" -ForegroundColor Red
        $validationErrors += "Subsite fehlt oder nicht erreichbar: $($web.Title) ($relativePath)"
        $results += [PSCustomObject]@{
            Level = "Subsite"
            Name = $web.Title
            Source_Files = "N/A"
            Target_Files = "MISSING"
            Status = "SITE_MISSING"
        }
    }
}

# ===== ABSCHLUSS =====
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Write-Host "`nFertig. Report: $OutputPath" -ForegroundColor Cyan