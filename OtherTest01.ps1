# Pfade zu den DLLs laden (Pfade können je nach Installation variieren, dies ist Standard für SP2013/2016/SE/Online)
try {
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
} catch {
    Write-Error "DLLs nicht gefunden. Bitte Pfade anpassen."
    return
}

# --- KONFIGURATION ---
$SiteUrl = "https://sp-se/sites/marketing"   # URL der neuen Site Collection
$ReportPath = "C:\Migration_Reports\Inventory_NEW_CSOM.csv"
# ---------------------

# Credentials abfragen
$creds = Get-Credential

$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$ctx.Credentials = New-Object Microsoft.SharePoint.Client.NetworkCredential($creds.UserName, $creds.GetNetworkCredential().Password, $creds.GetNetworkCredential().Domain)

$results = @()

# Funktion für Rekursion
function Get-WebInventoryCSOM {
    param([Microsoft.SharePoint.Client.Web]$currentWeb)

    Write-Host "Scanne: $($currentWeb.Url)" -ForegroundColor Yellow

    # Listen laden
    $lists = $currentWeb.Lists
    $ctx.Load($lists)
    $ctx.Load($currentWeb.Webs)
    $ctx.ExecuteQuery()

    foreach ($list in $lists) {
        # Wir müssen sicherstellen, dass wir alle Properties haben, die wir brauchen
        # Da wir oben nur die Collection geladen haben, sind einfache Properties da, aber sicherheitshalber:
        if ($list.Hidden -eq $false -and $list.IsCatalog -eq $false) {
            
            # Um die Ordnergröße zu bekommen, müssen wir den RootFolder explizit laden
            $ctx.Load($list.RootFolder)
            $ctx.ExecuteQuery()

            $sizeMB = [math]::Round(($list.RootFolder.Size / 1MB), 2)
            
            $type = "Liste"
            if ($list.BaseType -eq [Microsoft.SharePoint.Client.BaseType]::DocumentLibrary) { $type = "Bibliothek" }

            $item = New-Object PSObject -Property @{
                "SubsiteTitel"   = $currentWeb.Title
                "SubsiteUrl"     = $currentWeb.ServerRelativeUrl
                "ListenName"     = $list.Title
                "Typ"            = $type
                "ElementAnzahl"  = $list.ItemCount
                "GroesseMB"      = $sizeMB
                "LetzteAenderung"= $list.LastItemModifiedDate
            }
            $global:results += $item
        }
    }

    # Rekursion für Subsites
    foreach ($subWeb in $currentWeb.Webs) {
        # Subweb Kontext laden, sonst fehlen Properties beim nächsten Loop
        $ctx.Load($subWeb)
        $ctx.ExecuteQuery()
        Get-WebInventoryCSOM -currentWeb $subWeb
    }
}

try {
    Write-Host "Verbinde mit: $SiteUrl" -ForegroundColor Cyan
    $rootWeb = $ctx.Web
    $ctx.Load($rootWeb)
    $ctx.ExecuteQuery()

    # Start der Rekursion
    Get-WebInventoryCSOM -currentWeb $rootWeb

    # Export
    $global:results | Select-Object SubsiteTitel, ListenName, Typ, ElementAnzahl, GroesseMB, LetzteAenderung, SubsiteUrl | 
                      Sort-Object SubsiteUrl, ListenName | 
                      Export-Csv -Path $ReportPath -NoTypeInformation -Delimiter ";" -Encoding UTF8

    Write-Host "Fertig! Bericht unter: $ReportPath" -ForegroundColor Green

} catch {
    Write-Error "Fehler: $_"
}