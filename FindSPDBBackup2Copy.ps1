aram (
    [string]$TeileDesDateinamens,
    [string]$NetzwerkShare = '\\DC01\Verzeichnis10\FS01$Verzeichnis5',
    [string]$Verzeichnis3 = "C:\Pfad\zu\Verzeichnis3"
)

# Überprüfen, ob der Teilsuchbegriff übergeben wurde
if (-not $TeileDesDateinamens) {
    Write-Host "Bitte geben Sie Teile des Dateinamens als Parameter an."
    exit
}

# Suchen der Datei im Netzwerkshare
$gefunden = Get-ChildItem -Path $NetzwerkShare -Filter $TeileDesDateinamens -Recurse -ErrorAction SilentlyContinue

# Überprüfen, ob die Datei gefunden wurde
if ($gefunden) {
    foreach ($datei in $gefunden) {
        # Kopieren der gefundenen Datei nach Verzeichnis3
        Copy-Item -Path $datei.FullName -Destination $Verzeichnis3
        Write-Host "Datei '$($datei.Name)' wurde nach '$Verzeichnis3' kopiert."
    }
} else {
    Write-Host "Keine Datei mit dem Muster '$TeileDesDateinamens' wurde im Netzwerkshare '$NetzwerkShare' gefunden."
}
