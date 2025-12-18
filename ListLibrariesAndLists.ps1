# Migration Validation Script for SharePoint 2013 to SharePoint SE
# Vergleicht Quelle und Ziel Teamraum inklusive Struktur, Dateien und Berechtigungen

param(
    [Parameter(Mandatory=$true)][string]$SourceUrl,
    [Parameter(Mandatory=$true)][string]$TargetUrl,
    [string]$OutputPath = ".\MigrationValidation_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Laden Sie PnP PowerShell Module
$modules = @("PnP.PowerShell")
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installiere Modul: $module"
        Install-Module -Name $module -Force -AllowClobber
    }
    Import-Module $module
}

$results = @()
$validationErrors = @()

# ===== VERBINDUNG ZU SHAREPOINT =====
Write-Host "Verbinde mit SharePoint Instanzen..." -ForegroundColor Cyan

try {
    $sourceContext = Connect-PnPOnline -Url $SourceUrl -Interactive -ReturnConnection
    Write-Host "✓ Mit Quelle verbunden: $SourceUrl" -ForegroundColor Green
} catch {
    Write-Host "✗ Fehler bei Verbindung zur Quelle: $_" -ForegroundColor Red
    exit 1
}

try {
    $targetContext = Connect-PnPOnline -Url $TargetUrl -Interactive -ReturnConnection
    Write-Host "✓ Mit Ziel verbunden: $TargetUrl" -ForegroundColor Green
} catch {
    Write-Host "✗ Fehler bei Verbindung zum Ziel: $_" -ForegroundColor Red
    exit 1
}

# ===== FUNKTION: DATEIEN IN BIBLIOTHEK ZÄHLEN =====
function Get-FilesInLibrary {
    param(
        [object]$Context,
        [string]$LibraryName,
        [string]$Folder = ""
    )
    
    $fileCount = 0
    $folderCount = 0
    $details = @()
    
    try {
        if ([string]::IsNullOrEmpty($Folder)) {
            $items = Get-PnPListItem -List $LibraryName -PageSize 5000 -Connection $Context | Where-Object { $_.FileSystemObjectType -eq "File" }
        } else {
            $items = Get-PnPListItem -List $LibraryName -PageSize 5000 -Connection $Context -FolderServerRelativeUrl $Folder | Where-Object { $_.FileSystemObjectType -eq "File" }
        }
        
        $fileCount = @($items).Count
        
        # Zähle Ordner
        if ([string]::IsNullOrEmpty($Folder)) {
            $folders = Get-PnPListItem -List $LibraryName -PageSize 5000 -Connection $Context | Where-Object { $_.FileSystemObjectType -eq "Folder" }
        } else {
            $folders = Get-PnPListItem -List $LibraryName -PageSize 5000 -Connection $Context -FolderServerRelativeUrl $Folder | Where-Object { $_.FileSystemObjectType -eq "Folder" }
        }
        
        $folderCount = @($folders).Count
        
        return @{
            FileCount = $fileCount
            FolderCount = $folderCount
            Items = $items
            Folders = $folders
        }
    } catch {
        Write-Host "Fehler beim Zählen von Dateien in $LibraryName : $_" -ForegroundColor Yellow
        return @{ FileCount = 0; FolderCount = 0; Items = @(); Folders = @() }
    }
}

# ===== FUNKTION: SUBSITES ABRUFEN =====
function Get-AllSubsites {
    param([object]$Context)
    
    $subsites = @()
    try {
        $webs = Get-PnPSubWeb -Connection $Context -Recurse
        $subsites = @($webs)
    } catch {
        Write-Host "Fehler beim Abrufen von Subsites: $_" -ForegroundColor Yellow
    }
    
    return $subsites
}

# ===== FUNKTION: BERECHTIGUNGEN VERGLEICHEN =====
function Compare-Permissions {
    param(
        [object]$SourceContext,
        [object]$TargetContext,
        [string]$ListName
    )
    
    $permissionDifferences = @()
    
    try {
        $sourcePerms = Get-PnPRoleAssignment -List $ListName -Connection $SourceContext
        $targetPerms = Get-PnPRoleAssignment -List $ListName -Connection $TargetContext
        
        $sourcePrincipalIds = $sourcePerms.Member.LoginName | Sort-Object
        $targetPrincipalIds = $targetPerms.Member.LoginName | Sort-Object
        
        $missing = @($sourcePrincipalIds | Where-Object { $_ -notin $targetPrincipalIds })
        $extra = @($targetPrincipalIds | Where-Object { $_ -notin $sourcePrincipalIds })
        
        if ($missing.Count -gt 0) {
            $permissionDifferences += "FEHLEND (im Ziel): $($missing -join ', ')"
        }
        if ($extra.Count -gt 0) {
            $permissionDifferences += "EXTRA (im Ziel): $($extra -join ', ')"
        }
        
        return $permissionDifferences
    } catch {
        return @("Fehler beim Abrufen von Berechtigungen")
    }
}

# ===== HAUPTVERGLEICH: ROOT-BIBLIOTHEKEN UND LISTEN =====
Write-Host "`n=== VALIDIERE ROOT-BIBLIOTHEKEN UND LISTEN ===" -ForegroundColor Cyan

$sourceLibs = Get-PnPList -Connection $sourceContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }
$targetLibs = Get-PnPList -Connection $targetContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }

$sourceLibNames = $sourceLibs.Title | Sort-Object
$targetLibNames = $targetLibs.Title | Sort-Object

Write-Host "Quell-Bibliotheken: $($sourceLibNames.Count)"
Write-Host "Ziel-Bibliotheken: $($targetLibNames.Count)"

foreach ($lib in $sourceLibs) {
    Write-Host "`nBibliothek: $($lib.Title)" -ForegroundColor Yellow
    
    $sourceFileInfo = Get-FilesInLibrary -Context $sourceContext -LibraryName $lib.Title
    
    if ($targetLibs.Title -contains $lib.Title) {
        $targetFileInfo = Get-FilesInLibrary -Context $targetContext -LibraryName $lib.Title
        
        $status = if ($sourceFileInfo.FileCount -eq $targetFileInfo.FileCount -and $sourceFileInfo.FolderCount -eq $targetFileInfo.FolderCount) { "✓ OK" } else { "✗ UNTERSCHIED" }
        
        Write-Host "  Status: $status"
        Write-Host "  Quelle: $($sourceFileInfo.FileCount) Dateien, $($sourceFileInfo.FolderCount) Ordner"
        Write-Host "  Ziel:   $($targetFileInfo.FileCount) Dateien, $($targetFileInfo.FolderCount) Ordner"
        
        $permDiff = Compare-Permissions -SourceContext $sourceContext -TargetContext $targetContext -ListName $lib.Title
        if ($permDiff.Count -gt 0) {
            Write-Host "  ⚠ Berechtigungen unterscheiden sich:"
            $permDiff | ForEach-Object { Write-Host "    - $_" }
        }
        
        $results += [PSCustomObject]@{
            Typ = "Bibliothek (Root)"
            Name = $lib.Title
            Quelle_Dateien = $sourceFileInfo.FileCount
            Ziel_Dateien = $targetFileInfo.FileCount
            Quelle_Ordner = $sourceFileInfo.FolderCount
            Ziel_Ordner = $targetFileInfo.FolderCount
            Status = $status
            Berechtigungsprobleme = $permDiff -join "; "
        }
    } else {
        Write-Host "  ✗ BIBLIOTHEK FEHLT IM ZIEL!"
        $results += [PSCustomObject]@{
            Typ = "Bibliothek (Root)"
            Name = $lib.Title
            Quelle_Dateien = $sourceFileInfo.FileCount
            Ziel_Dateien = "N/A - FEHLT"
            Quelle_Ordner = $sourceFileInfo.FolderCount
            Ziel_Ordner = "N/A - FEHLT"
            Status = "✗ FEHLT"
            Berechtigungsprobleme = ""
        }
        $validationErrors += "Bibliothek fehlt: $($lib.Title)"
    }
}

# Prüfe auf Extra-Bibliotheken im Ziel
$extraLibs = $targetLibNames | Where-Object { $_ -notin $sourceLibNames }
if ($extraLibs.Count -gt 0) {
    Write-Host "`n⚠ Extra-Bibliotheken im Ziel: $($extraLibs -join ', ')" -ForegroundColor Yellow
    $extraLibs | ForEach-Object {
        $results += [PSCustomObject]@{
            Typ = "Bibliothek (Root)"
            Name = $_
            Quelle_Dateien = "N/A"
            Ziel_Dateien = "N/A"
            Quelle_Ordner = "N/A"
            Ziel_Ordner = "N/A"
            Status = "⚠ EXTRA IM ZIEL"
            Berechtigungsprobleme = ""
        }
    }
}

# ===== VALIDIERE SUBSITES =====
Write-Host "`n=== VALIDIERE SUBSITES ===" -ForegroundColor Cyan

$sourceSubsites = Get-AllSubsites -Context $sourceContext
$targetSubsites = Get-AllSubsites -Context $targetContext

Write-Host "Quell-Subsites: $($sourceSubsites.Count)"
Write-Host "Ziel-Subsites: $($targetSubsites.Count)"

$sourceSubsiteNames = $sourceSubsites.Title | Sort-Object
$targetSubsiteNames = $targetSubsites.Title | Sort-Object

foreach ($subsite in $sourceSubsites) {
    Write-Host "`nSubsite: $($subsite.Title) [$($subsite.Url)]" -ForegroundColor Yellow
    
    if ($targetSubsites.Title -contains $subsite.Title) {
        Write-Host "  ✓ Subsite existiert im Ziel"
        
        # Verbinde mit Subsite und validiere Inhalte
        try {
            $sourceSubContext = Connect-PnPOnline -Url $subsite.Url -Interactive -ReturnConnection
            $targetSubUrl = $TargetUrl.TrimEnd('/') + "/" + $subsite.Title
            $targetSubContext = Connect-PnPOnline -Url $targetSubUrl -Interactive -ReturnConnection
            
            $sourceSubLibs = Get-PnPList -Connection $sourceSubContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }
            $targetSubLibs = Get-PnPList -Connection $targetSubContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }
            
            Write-Host "    Bibliotheken in Subsite - Quelle: $($sourceSubLibs.Count), Ziel: $($targetSubLibs.Count)"
            
            foreach ($subLib in $sourceSubLibs) {
                $sourceSubFileInfo = Get-FilesInLibrary -Context $sourceSubContext -LibraryName $subLib.Title
                
                if ($targetSubLibs.Title -contains $subLib.Title) {
                    $targetSubFileInfo = Get-FilesInLibrary -Context $targetSubContext -LibraryName $subLib.Title
                    
                    $status = if ($sourceSubFileInfo.FileCount -eq $targetSubFileInfo.FileCount -and $sourceSubFileInfo.FolderCount -eq $targetSubFileInfo.FolderCount) { "✓ OK" } else { "✗ UNTERSCHIED" }
                    
                    Write-Host "      $($subLib.Title): $status (Q: $($sourceSubFileInfo.FileCount) Dateien, Z: $($targetSubFileInfo.FileCount) Dateien)"
                    
                    $results += [PSCustomObject]@{
                        Typ = "Bibliothek (Subsite)"
                        Name = "$($subsite.Title)/$($subLib.Title)"
                        Quelle_Dateien = $sourceSubFileInfo.FileCount
                        Ziel_Dateien = $targetSubFileInfo.FileCount
                        Quelle_Ordner = $sourceSubFileInfo.FolderCount
                        Ziel_Ordner = $targetSubFileInfo.FolderCount
                        Status = $status
                        Berechtigungsprobleme = ""
                    }
                } else {
                    Write-Host "      ✗ $($subLib.Title): FEHLT IM ZIEL" -ForegroundColor Red
                    $validationErrors += "Bibliothek fehlt in Subsite $($subsite.Title): $($subLib.Title)"
                }
            }
        } catch {
            Write-Host "    ⚠ Fehler beim Zugriff auf Subsite: $_" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  ✗ Subsite fehlt im Ziel!" -ForegroundColor Red
        $validationErrors += "Subsite fehlt: $($subsite.Title)"
    }
}

# ===== ZUSAMMENFASSUNG =====
Write-Host "`n=== VALIDIERUNGSBERICHT ===" -ForegroundColor Cyan
Write-Host "Gesamtprobleme gefunden: $($validationErrors.Count)"

if ($validationErrors.Count -gt 0) {
    Write-Host "`nProbleme:" -ForegroundColor Red
    $validationErrors | ForEach-Object { Write-Host "  ✗ $_" }
} else {
    Write-Host "✓ Keine Probleme gefunden!" -ForegroundColor Green
}

# Exportiere Ergebnisse in CSV
$results | Export-Csv -Path $OutputPath -Encoding UTF8 -NoTypeInformation -Delimiter ";"
Write-Host "`n✓ Detaillierter Report exportiert zu: $OutputPath" -ForegroundColor Green

# Zeige Zusammenfassung
Write-Host "`nZusammenfassung:" -ForegroundColor Cyan
$results | Group-Object -Property Status | Select-Object Name, Count

# Disconnect
Disconnect-PnPOnline -Connection $sourceContext
Disconnect-PnPOnline -Connection $targetContext

Write-Host "`n✓ Validierung abgeschlossen!" -ForegroundColor Green
