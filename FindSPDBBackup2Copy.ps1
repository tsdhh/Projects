param (
    [string]$fileName,
    [string]$Verzeichnis1 = "C:\Pfad\zu\Verzeichnis1",
    [string]$Verzeichnis2 = "C:\Pfad\zu\Verzeichnis2",
    [string]$Verzeichnis3 = "C:\Pfad\zu\Verzeichnis3"
)

# Überprüfen, ob der Dateiname übergeben wurde
if (-not $fileName) {
    Write-Host "Bitte geben Sie einen Dateinamen als Parameter an."
    exit
}

# Suchen der Datei in Verzeichnis1 und Verzeichnis2
$gefunden = Get-ChildItem -Path $Verzeichnis1, $Verzeichnis2 -Filter $fileName -Recurse -ErrorAction SilentlyContinue

# Überprüfen, ob die Datei gefunden wurde
if ($gefunden) {
    # Kopieren der gefundenen Datei nach Verzeichnis3
    Copy-Item -Path $gefunden.FullName -Destination $Verzeichnis3
    Write-Host "Datei '$fileName' wurde nach '$Verzeichnis3' kopiert."
} else {
    Write-Host "Datei '$fileName' wurde nicht in den angegebenen Verzeichnissen gefunden."
}