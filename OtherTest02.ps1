Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# --- KONFIGURATION ---
$SiteCollectionUrl = "http://sp2013/sites/marketing" # Ihre SP2013 URL
$ReportPath = "C:\Migration_Reports\Inventory_OLD_Fixed.csv"
# ---------------------

$results = @()

try {
    Write-Host "Verbinde mit Site Collection: $SiteCollectionUrl" -ForegroundColor Cyan
    $site = Get-SPSite $SiteCollectionUrl
    
    foreach ($web in $site.AllWebs) {
        Write-Host "Scanne Subsite: $($web.Title)" -ForegroundColor Yellow
        
        foreach ($list in $web.Lists) {
            # Wir geben ALLES aus, was nicht versteckt ist. 
            # Wir entfernen den 'IsCatalog' Filter vorerst, um sicherzugehen, dass wir nichts verpassen.
            if ($list.Hidden -eq $false) {
                
                # 1. Typ sicher bestimmen (Über Enum statt String)
                $typeLabel = "Liste"
                $isLibrary = $false
                
                if ($list.BaseType -eq [Microsoft.SharePoint.SPBaseType]::DocumentLibrary) { 
                    $typeLabel = "Bibliothek" 
                    $isLibrary = $true
                }

                # 2. Größe berechnen (Fehleranfällig bei Listen in SP2013, daher Try-Catch)
                $sizeMB = 0
                try {
                    # Größe ist nur bei Bibliotheken wirklich relevant/zuverlässig abrufbar
                    if ($isLibrary) {
                        $sizeMB = [math]::Round(($list.RootFolder.Size / 1MB), 2)
                    }
                } catch {
                    $sizeMB = -1 # Markierung für Fehler bei Berechnung
                }

                # Debug-Ausgabe in die Konsole, damit Sie sofort sehen, ob Listen gefunden werden
                # Write-Host " -> Gefunden: [$typeLabel] $($list.Title)" -ForegroundColor Gray

                $item = New-Object PSObject -Property @{
                    "SubsiteTitel"   = $web.Title
                    "SubsiteUrl"     = $web.ServerRelativeUrl
                    "ListenName"     = $list.Title
                    "Typ"            = $typeLabel
                    "ElementAnzahl"  = $list.ItemCount
                    "GroesseMB"      = $sizeMB
                    "LetzteAenderung"= $list.LastItemModifiedDate
                    "ListenUrl"      = $list.RootFolder.ServerRelativeUrl
                }
                $results += $item
            }
        }
        $web.Dispose()
    }

    # Exportieren
    $results | Select-Object SubsiteTitel, ListenName, Typ, ElementAnzahl, GroesseMB, LetzteAenderung, SubsiteUrl | 
               Sort-Object SubsiteUrl, ListenName | 
               Export-Csv -Path $ReportPath -NoTypeInformation -Delimiter ";" -Encoding UTF8

    Write-Host "------------------------------------------------"
    Write-Host "Bericht gespeichert: $ReportPath" -ForegroundColor Green

} catch {
    Write-Error "Kritischer Fehler: $_"
} finally {
    if ($site) { $site.Dispose() }
}