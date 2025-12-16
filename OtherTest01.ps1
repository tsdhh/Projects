Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# --- KONFIGURATION ---
$SiteCollectionUrl = "http://sp2013/sites/marketing" # URL der Site Collection (Wurzel)
$ReportPath = "C:\Migration_Reports\FullInventory_OLD.csv" 
# ---------------------

$results = @()

try {
    Write-Host "Verbinde mit Site Collection: $SiteCollectionUrl" -ForegroundColor Cyan
    $site = Get-SPSite $SiteCollectionUrl
    
    # Wir iterieren durch ALLE Webs (Hauptseite + alle Subsites)
    foreach ($web in $site.AllWebs) {
        Write-Host "Scanne Subsite: $($web.Title) ($($web.Url))" -ForegroundColor Yellow
        
        try {
            foreach ($list in $web.Lists) {
                # Systemlisten und Kataloge ausschließen, um den Bericht lesbar zu halten
                if ($list.Hidden -eq $false -and $list.IsCatalog -eq $false) {
                    
                    # Größe berechnen (RootFolder Size gibt Bytes zurück)
                    $sizeMB = [math]::Round(($list.RootFolder.Size / 1MB), 2)
                    
                    # Listentyp bestimmen (Bibliothek oder Liste)
                    $type = "Liste"
                    if ($list.BaseType -eq "DocumentLibrary") { $type = "Bibliothek" }

                    $item = New-Object PSObject -Property @{
                        "SubsiteTitel"   = $web.Title
                        "SubsiteUrl"     = $web.ServerRelativeUrl
                        "ListenName"     = $list.Title
                        "Typ"            = $type
                        "ElementAnzahl"  = $list.ItemCount
                        "GroesseMB"      = $sizeMB
                        "LetzteAenderung"= $list.LastItemModifiedDate
                        "ListenUrl"      = $list.RootFolder.ServerRelativeUrl
                    }
                    $results += $item
                }
            }
        } catch {
            Write-Error "Fehler beim Lesen von Web $($web.Url): $_"
        } finally {
            $web.Dispose()
        }
    }

    # Exportieren
    # Sortierung: Erst nach Subsite-URL, dann nach Listenname
    $results | Select-Object SubsiteTitel, ListenName, Typ, ElementAnzahl, GroesseMB, LetzteAenderung, SubsiteUrl, ListenUrl | 
               Sort-Object SubsiteUrl, ListenName | 
               Export-Csv -Path $ReportPath -NoTypeInformation -Delimiter ";" -Encoding UTF8

    Write-Host "------------------------------------------------"
    Write-Host "Vollständiger Bericht gespeichert: $ReportPath" -ForegroundColor Green

} catch {
    Write-Error "Kritischer Fehler: $_"
} finally {
    if ($site) { $site.Dispose() }
}