Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# --- KONFIGURATION ---
$WebAppURL = "http://ihre-sharepoint-url" 
$OutputPath = "C:\Temp\AusgecheckteDateien.csv"

# --- SKRIPT START ---
$results = @()
$counter = 0

Write-Host "Starte Analyse fuer WebApplication: $WebAppURL" -ForegroundColor Cyan

Start-SPAssignment -Global

try {
    $webApp = Get-SPWebApplication $WebAppURL
    $sites = $webApp.Sites
    $totalSites = $sites.Count

    foreach ($site in $sites) {
        $counter++
        Write-Progress -Activity "Scanne Site Collections" -Status "$($site.Url)" -PercentComplete (($counter / $totalSites) * 100)

        try {
            foreach ($web in $site.AllWebs) {
                # Nur Dokumentenbibliotheken
                $lists = $web.Lists | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false }

                foreach ($list in $lists) {
                    
                    # 1. PRÜFUNG: Regulär ausgecheckte Dateien
                    $query = New-Object Microsoft.SharePoint.SPQuery
                    $query.ViewAttributes = "Scope='RecursiveAll'"
                    $query.Query = "<Where><IsNotNull><FieldRef Name='CheckoutUser' /></IsNotNull></Where>"
                    
                    $items = $list.GetItems($query)

                    foreach ($item in $items) {
                        if ($item.File.CheckOutType -ne "None") {
                            # Variablen setzen für bessere Lesbarkeit
                            $checkOutUser = $item.File.CheckedOutByUser.LoginName
                            $timeLastMod = $item.File.TimeLastModified
                            
                            $obj = New-Object PSObject
                            $obj | Add-Member -MemberType NoteProperty -Name "Teamraum" -Value $web.Title
                            $obj | Add-Member -MemberType NoteProperty -Name "URL" -Value $web.Url
                            $obj | Add-Member -MemberType NoteProperty -Name "Liste" -Value $list.Title
                            $obj | Add-Member -MemberType NoteProperty -Name "Datei" -Value $item.Name
                            $obj | Add-Member -MemberType NoteProperty -Name "Pfad" -Value $item.File.ServerRelativeUrl
                            $obj | Add-Member -MemberType NoteProperty -Name "User" -Value $checkOutUser
                            $obj | Add-Member -MemberType NoteProperty -Name "Typ" -Value "Normal"
                            $obj | Add-Member -MemberType NoteProperty -Name "Datum" -Value $timeLastMod
                            
                            $results += $obj
                            Write-Host "Gefunden: $($item.Name)" -ForegroundColor Yellow
                        }
                    }

                    # 2. PRÜFUNG: Nie eingecheckte Dateien (Geisterdateien)
                    $checkedOutFiles = $list.CheckedOutFiles
                    
                    if ($checkedOutFiles.Count -gt 0) {
                        foreach ($cof in $checkedOutFiles) {
                            # Pfad zusammenbauen
                            $relWeb = $web.ServerRelativeUrl.TrimEnd('/')
                            $listUrl = $list.RootFolder.Url
                            $fPath = "$relWeb/$listUrl/$($cof.LeafName)"
                            $cUser = $cof.CheckedOutBy.LoginName

                            $obj = New-Object PSObject
                            $obj | Add-Member -MemberType NoteProperty -Name "Teamraum" -Value $web.Title
                            $obj | Add-Member -MemberType NoteProperty -Name "URL" -Value $web.Url
                            $obj | Add-Member -MemberType NoteProperty -Name "Liste" -Value $list.Title
                            $obj | Add-Member -MemberType NoteProperty -Name "Datei" -Value $cof.LeafName
                            $obj | Add-Member -MemberType NoteProperty -Name "Pfad" -Value $fPath
                            $obj | Add-Member -MemberType NoteProperty -Name "User" -Value $cUser
                            $obj | Add-Member -MemberType NoteProperty -Name "Typ" -Value "Nie eingecheckt"
                            $obj | Add-Member -MemberType NoteProperty -Name "Datum" -Value "Unbekannt"

                            $results += $obj
                            Write-Host "Nie eingecheckt: $($cof.LeafName)" -ForegroundColor Red
                        }
                    }
                }
            }
        }
        catch {
            Write-Host "Fehler bei Web: $($web.Url)" -ForegroundColor Red
        }
    }
}
finally {
    Stop-SPAssignment -Global
}

# --- EXPORT ---
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    Write-Host "CSV gespeichert: $OutputPath" -ForegroundColor Green
}
else {
    Write-Host "Nichts gefunden." -ForegroundColor Green
}