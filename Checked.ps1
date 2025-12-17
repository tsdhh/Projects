Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# --- KONFIGURATION ---
# Geben Sie hier die URL Ihrer Webanwendung an (z.B. http://portal)
$WebAppURL = "http://ihre-sharepoint-url" 
# Pfad für die CSV-Ausgabe
$OutputPath = "C:\Temp\AusgecheckteDateien.csv"

# --- SKRIPT START ---
$results = @()
$counter = 0

Write-Host "Starte Analyse für WebApplication: $WebAppURL" -ForegroundColor Cyan

Start-SPAssignment -Global # Speichermanagement für SP-Objekte

try {
    $webApp = Get-SPWebApplication $WebAppURL
    $sites = $webApp.Sites
    $totalSites = $sites.Count

    foreach ($site in $sites) {
        $counter++
        Write-Progress -Activity "Scanne Site Collections" -Status "$($site.Url)" -PercentComplete (($counter / $totalSites) * 100)

        try {
            foreach ($web in $site.AllWebs) {
                # Nur Dokumentenbibliotheken prüfen, keine Kataloge/Systemlisten
                $lists = $web.Lists | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false }

                foreach ($list in $lists) {
                    
                    # 1. PRÜFUNG: Regulär ausgecheckte Dateien
                    # Wir nutzen eine CAML Query für Performance (filtert serverseitig)
                    $query = New-Object Microsoft.SharePoint.SPQuery
                    $query.ViewAttributes = "Scope='RecursiveAll'"
                    $query.Query = "<Where><IsNotNull><FieldRef Name='CheckoutUser' /></IsNotNull></Where>"
                    
                    $items = $list.GetItems($query)

                    foreach ($item in $items) {
                        if ($item.File.CheckOutType -ne "None") {
                            $obj = New-Object PSObject
                            $obj | Add-Member -MemberType NoteProperty -Name "Teamraum Name" -Value $web.Title
                            $obj | Add-Member -MemberType NoteProperty -Name "Teamraum URL" -Value $web.Url
                            $obj | Add-Member -MemberType NoteProperty -Name "Bibliothek" -Value $list.Title
                            $obj | Add-Member -MemberType NoteProperty -Name "Dateiname" -Value $item.Name
                            $obj | Add-Member -MemberType NoteProperty -Name "Pfad" -Value $item.File.ServerRelativeUrl
                            $obj | Add-Member -MemberType NoteProperty -Name "Ausgecheckt von" -Value $item.File.CheckedOutByUser.LoginName
                            $obj | Add-Member -MemberType NoteProperty -Name "Status Typ" -Value "Regulär Ausgecheckt"
                            $obj | Add-Member -MemberType NoteProperty -Name "Letzte Änderung" -Value $item.File.TimeLastModified
                            
                            $results += $obj
                            Write-Host "Gefunden: $($item.Name) in $($web.Title)" -ForegroundColor Yellow
                        }
                    }

                    # 2. PRÜFUNG: Dateien, die noch NIE eingecheckt wurden (No Checked In Version)
                    # Diese tauchen in $list.Items NICHT auf!
                    $checkedOutFiles = $list.CheckedOutFiles
                    
                    if ($checkedOutFiles.Count -gt 0) {
                        foreach ($cof in $checkedOutFiles) {
                            $obj = New-Object PSObject
                            $obj | Add-Member -MemberType NoteProperty -Name "Teamraum Name" -Value $web.Title
                            $obj | Add-Member -MemberType NoteProperty -Name "Teamraum URL" -Value $web.Url
                            $obj | Add-Member -MemberType NoteProperty -Name "Bibliothek" -Value $list.Title
                            $obj | Add-Member -MemberType NoteProperty -Name "Dateiname" -Value $cof.LeafName
                            # Pfad bauen, da ServerRelativeUrl hier anders ist
                            $fullPath = $web.ServerRelativeUrl.TrimEnd('/') + "/" + $list.RootFolder.Url + "/" + $cof.LeafName
                            $obj | Add-Member -MemberType NoteProperty -Name "Pfad" -Value $fullPath
                            $obj | Add-Member -MemberType NoteProperty -Name "Ausgecheckt von" -Value $cof.CheckedOutBy.LoginName
                            $obj | Add-Member -MemberType NoteProperty -Name "Status Typ" -Value "Nie eingecheckt (Geisterdatei)"
                            $obj | Add-Member -MemberType NoteProperty -Name "Letzte Änderung" -Value "N/A"

                            $results += $obj
                            Write-Host "Geisterdatei gefunden: $($cof.LeafName)" -ForegroundColor Red
                        }
                    }
                }
            }
        }
        catch {
            Write-Host "Fehler beim Zugriff auf Web: $($web.Url) - $_" -ForegroundColor Red
        }
    }
}
finally {
    Stop-SPAssignment -Global
}

# --- EXPORT ---
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    Write-Host "Fertig! Bericht gespeichert unter: $OutputPath" -ForegroundColor Green
}
else {
    Write-Host "Keine ausgecheckten Dateien gefunden." -ForegroundColor Green
}