Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# --- KONFIGURATION ---
$SiteUrl = "http://sp2013/sites/marketing"  # <-- Hier jeweils anpassen (Alt oder Neu)
$ReportPath = "C:\Migration_Reports\Inventory_OLD.csv" # <-- Dateiname anpassen (z.B. _OLD.csv oder _NEW.csv)
# ---------------------

try {
    $web = Get-SPWeb $SiteUrl
    $lists = $web.Lists
    
    $results = @()

    Write-Host "Analysiere Inhalte von: $($web.Title) ($($web.Url))" -ForegroundColor Cyan

    foreach ($list in $lists) {
        # Wir ignorieren versteckte Systemlisten (Kataloge, interne Listen), um den Bericht sauber zu halten
        if ($list.Hidden -eq $false -and $list.IsCatalog -eq $false) {
            
            # Objekt für den Export bauen
            $item = New-Object PSObject -Property @{
                "ListenName"     = $list.Title
                "Typ"            = $list.BaseType
                "ElementAnzahl"  = $list.ItemCount
                "Url"            = $list.RootFolder.ServerRelativeUrl
                "LetzteAenderung"= $list.LastItemModifiedDate
            }
            $results += $item
            
            Write-Host "Gefunden: $($list.Title) - Anzahl: $($list.ItemCount)"
        }
    }

    # Exportieren als CSV
    # Wir sortieren nach Name für einfacheren Vergleich
    $results | Select-Object ListenName, ElementAnzahl, Typ, LetzteAenderung, Url | Sort-Object ListenName | Export-Csv -Path $ReportPath -NoTypeInformation -Delimiter ";" -Encoding UTF8

    Write-Host "------------------------------------------------"
    Write-Host "Bericht gespeichert unter: $ReportPath" -ForegroundColor Green

} catch {
    Write-Error "Fehler: $_"
} finally {
    if ($web) { $web.Dispose() }
}