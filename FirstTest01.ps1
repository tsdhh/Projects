Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# --- KONFIGURATION ---
$OldSiteUrl = "http://sp2013/sites/marketing"  # Die URL des alten Teamraums
$NewSiteUrl = "https://sp-se/sites/marketing"  # Die URL des neuen Teamraums
# ---------------------

try {
    Write-Host "Verbinde mit Seite: $OldSiteUrl" -ForegroundColor Cyan
    $web = Get-SPWeb $OldSiteUrl
    
    # 1. Die Startseite (Welcome Page) ermitteln
    $welcomePageUrl = $web.RootFolder.WelcomePage
    $file = $web.GetFile($welcomePageUrl)
    
    # Prüfen, ob die Datei ausgecheckt werden muss
    if ($file.CheckOutType -ne "None") {
        Write-Warning "Seite ist bereits ausgecheckt von: $($file.CheckedOutByUser.LoginName)"
        # Option: $file.UndoCheckOut() # Falls gewünscht, Check-Out erzwingen
    }
    
    # Seite auschecken (falls Versionierung aktiviert ist)
    if ($file.Level -eq "Published" -and $web.Lists[$file.ParentFolder.ParentListId].EnableVersioning) {
        $file.CheckOut()
    }

    # 2. WebPartManager holen
    $wpm = $file.GetLimitedWebPartManager([System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)

    # 3. HTML für das Banner vorbereiten
    $htmlContent = @"
<div style="background-color: #f44336; color: white; padding: 20px; text-align: center; font-size: 18px; border: 2px solid darkred; margin-bottom: 20px;">
    <strong>ACHTUNG:</strong> Dieser Teamraum wurde migriert! <br/>
    Bitte arbeiten Sie ab sofort nur noch im neuen System.<br/><br/>
    <a href="$NewSiteUrl" style="color: yellow; text-decoration: underline; font-weight: bold; font-size: 20px;">
        Hier klicken, um zum neuen Teamraum zu gelangen
    </a>
</div>
"@

    # 4. Script Editor Webpart erstellen
    $webPart = New-Object Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart
    $webPart.Title = "Migrations-Hinweis"
    $webPart.ChromeType = [System.Web.UI.WebControls.WebParts.PartChromeType]::None # Kein Rahmen um das Webpart
    $webPart.Content = $htmlContent

    # 5. Webpart zur Seite hinzufügen
    # Wir suchen die erste verfügbare Zone. Bei Wiki-Pages oft "Main", "Body" oder "Right".
    # ZoneID null oder erste Zone nutzen, Index 0 setzt es ganz nach oben.
    
    $zoneId = "Main" # Standard für viele SP2013 Seiten
    if ($wpm.Zones.Count -gt 0) {
        # Falls "Main" nicht existiert, nehmen wir die erste, die wir finden
        if (-not $wpm.Zones["Main"]) {
            $zoneId = $wpm.Zones[0].ID 
        }
    }
    
    Write-Host "Füge Webpart zur Zone '$zoneId' hinzu..."
    $wpm.AddWebPart($webPart, $zoneId, 0)

    # 6. Speichern und Veröffentlichen
    $file.Update()
    
    if ($file.Level -ne "Published" -and $web.Lists[$file.ParentFolder.ParentListId].EnableVersioning) {
        $file.CheckIn("Banner für Migration hinzugefügt via PowerShell")
        $file.Publish("Banner für Migration hinzugefügt via PowerShell")
    }
    
    # Optional: Schreibschutz setzen (Empfohlen!)
    # Set-SPSite -Identity $OldSiteUrl -LockState ReadOnly
    
    Write-Host "Erfolg! Banner wurde gesetzt." -ForegroundColor Green

} catch {
    Write-Error "Fehler aufgetreten: $_"
} finally {
    if ($web) { $web.Dispose() }
}