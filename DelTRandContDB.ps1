param (
    [string]$SiteUrl,
    [string]$DatabaseName
)

# SharePoint-Umgebung laden
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Überprüfen, ob die Parameter übergeben wurden
if (-not $SiteUrl) {
    Write-Host "Bitte geben Sie die URL der zu löschenden Webseite als Parameter an."
    exit
}

if (-not $DatabaseName) {
    Write-Host "Bitte geben Sie den Namen der zu löschenden Content-Datenbank als Parameter an."
    exit
}

# Löschen der Teamraum-Webseite
$site = Get-SPSite $SiteUrl
if ($site -ne $null) {
    $site.Delete()
    Write-Host "Der Teamraum '$SiteUrl' wurde gelöscht."

    # Löschen der Content-Datenbank
    $database = Get-SPContentDatabase -Identity $DatabaseName
    if ($database -ne $null) {
        Remove-SPContentDatabase -Identity $database -Confirm:$false
        Write-Host "Die Content-Datenbank '$DatabaseName' wurde gelöscht."
    } else {
        Write-Host "Die Content-Datenbank '$DatabaseName' wurde nicht gefunden."
    }
} else {
    Write-Host "Die Website '$SiteUrl' wurde nicht gefunden."
}
