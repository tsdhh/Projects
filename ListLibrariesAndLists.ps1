# Migration Validation Script for SharePoint 2013 to SharePoint SE
# Version: FIXED BRACKETS

param(
    [Parameter(Mandatory=$true)][string]$SourceUrl,
    [Parameter(Mandatory=$true)][string]$TargetUrl,
    [string]$OutputPath = ".\MigrationValidation_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [switch]$PromptForCredentials 
)

# Prüfung auf Modul
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    Write-Warning "Das Modul 'PnP.PowerShell' fehlt."
}
Import-Module PnP.PowerShell -ErrorAction SilentlyContinue

$results = @()
$validationErrors = @()

$SourceUrl = $SourceUrl.TrimEnd('/')
$TargetUrl = $TargetUrl.TrimEnd('/')

# ===== VERBINDUNG ZU SHAREPOINT =====
Write-Host "Verbinde mit SharePoint Instanzen..." -ForegroundColor Cyan

function Connect-SPOnPrem {
    param($Url)
    try {
        if ($PromptForCredentials) {
            $creds = Get-Credential
            return Connect-PnPOnline -Url $Url -Credentials $creds -ReturnConnection -ErrorAction Stop
        } else {
            return Connect-PnPOnline -Url $Url -CurrentCredentials -ReturnConnection -ErrorAction Stop
        }
    } catch {
        throw $_
    }
}

try {
    $sourceContext = Connect-SPOnPrem -Url $SourceUrl
    $sourceRootWeb = Get-PnPWeb -Connection $sourceContext
    $sourceRootRelUrl = $sourceRootWeb.ServerRelativeUrl
    Write-Host "✓ Mit Quelle verbunden: $SourceUrl" -ForegroundColor Green
} catch {
    Write-Host "✗ Fehler bei Verbindung zur Quelle: $_" -ForegroundColor Red
    exit 1
}

try {
    $targetContext = Connect-SPOnPrem -Url $TargetUrl
    $targetRootWeb = Get-PnPWeb -Connection $targetContext
    $targetRootRelUrl = $targetRootWeb.ServerRelativeUrl
    Write-Host "✓ Mit Ziel verbunden: $TargetUrl" -ForegroundColor Green
} catch {
    Write-Host "✗ Fehler bei Verbindung zum Ziel: $_" -ForegroundColor Red
    exit 1
}

# ===== FUNKTIONEN =====

function Get-FilesInLibrary {
    param(
        [object]$Context,
        [string]$LibraryTitle
    )
    
    try {
        Get-PnPList -Identity $LibraryTitle -Connection $Context -ErrorAction Stop | Out-Null
        
        $filesQuery = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>0</Value></Eq></Where></Query></View>"
        $foldersQuery = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>1</Value></Eq></Where></Query></View>"
        
        $files = Get-PnPListItem -List $LibraryTitle -Connection $Context -Query $filesQuery -PageSize 2000
        $folders = Get-PnPListItem -List $LibraryTitle -Connection $Context -Query $foldersQuery -PageSize 2000
        
        return @{
            FileCount = $files.Count
            FolderCount = $folders.Count
        }
    } catch {
        return @{ FileCount = -1; FolderCount = -1 }
    }
}

function Compare-Permissions {
    param($SourceContext, $TargetContext, $ListName)
    try {
        $sourcePerms = Get-PnPRoleAssignment -List $ListName -Connection $SourceContext
        $targetPerms = Get-PnPRoleAssignment -List $ListName -Connection $TargetContext
        
        $srcUsers = $sourcePerms.Member.LoginName | ForEach-Object { $_ -replace "^.*\|", "" }
        $tgtUsers = $targetPerms.Member.LoginName | ForEach-Object { $_ -replace "^.*\|", "" }
        
        $missing = $srcUsers | Where-Object { $_ -notin $tgtUsers }
        
        if ($missing) { return "Fehlend: $($missing -join ', ')" }
        return $null
    } catch {
        return "Fehler beim Perm-Check"
    }
}

# ===== LOGIK: ROOT BIBLIOTHEKEN =====
Write-Host "`n=== VALIDIERE ROOT-BIBLIOTHEKEN ===" -ForegroundColor Cyan

$sourceLibs = Get-PnPList -Connection $sourceContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden -and $_.Title -ne "Microfeed" }
$targetLibs = Get-PnPList -Connection $targetContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }

foreach ($lib in $sourceLibs) {
    Write-Host -NoNewline "Prüfe $($lib.Title)... "
    
    $targetLib = $targetLibs | Where-Object { $_.Title -eq $lib.Title }
    
    if ($targetLib) {
        $sInfo = Get-FilesInLibrary -Context $sourceContext -LibraryTitle $lib.Title
        $tInfo = Get-FilesInLibrary -Context $targetContext -LibraryTitle $targetLib.Title
        
        $status = if ($sInfo.FileCount -eq $tInfo.FileCount) { "OK" } else { "DIFF" }
        $color = if ($status -eq "OK") { "Green" } else { "Red" }
        
        Write-Host "$status (Q: $($sInfo.FileCount), Z: $($tInfo.FileCount))" -ForegroundColor $color
        
        $permRes = Compare-Permissions -SourceContext $sourceContext -TargetContext $targetContext -ListName $lib.Title
        
        $results += [PSCustomObject]@{
            Level = "Root"
            Name = $lib.Title
            Source_Files = $sInfo.FileCount
            Target_Files = $tInfo.FileCount
            Status = $status
            Permissions = $permRes
        }
    } else {
        Write-Host "FEHLT IM ZIEL" -ForegroundColor Red
        $validationErrors += "Root-Lib fehlt: $($lib.Title)"
        $results += [PSCustomObject]@{ Level="Root"; Name=$lib.Title; Status="MISSING"; Source_Files=""; Target_Files="" }
    }
}

# ===== LOGIK: SUBSITES (Rekursiv) =====
Write-Host "`n=== VALIDIERE SUBSITES ===" -ForegroundColor Cyan

$sourceSubwebs = Get-PnPSubWeb -Connection $sourceContext -Recurse -Includes ServerRelativeUrl, Title

foreach ($web in $sourceSubwebs) {
    # URL Berechnung
    $relativePath = $web.ServerRelativeUrl.Substring($sourceRootRelUrl.Length)
    $expectedTargetUrl = $targetRootRelUrl + $relativePath
    
    $targetUriObj = [System.Uri]$TargetUrl
    $fullTargetWebUrl = $targetUriObj.Scheme + "://" + $targetUriObj.Authority + $expectedTargetUrl
    
    Write-Host "`nSubsite: $($web.Title)" -ForegroundColor Yellow
    
    try {
        # 1. Verbindung Ziel
        $targetSubContext = Connect-SPOnPrem -Url $fullTargetWebUrl
        Write-Host "  ✓ Subsite im Ziel gefunden" -ForegroundColor Green
        
        # 2. Verbindung Quelle
        $sourceSubContext = Connect-SPOnPrem -Url $web.Url
        
        # 3. Bibliotheken vergleichen
        $subSourceLibs = Get-PnPList -Connection $sourceSubContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden -and $_.Title -ne "Microfeed" }
        $subTargetLibs = Get-PnPList -Connection $targetSubContext | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }
        
        foreach ($slib in $subSourceLibs) {
            $tlib = $subTargetLibs | Where-Object { $_.Title -eq $slib.Title }
            
            if ($tlib) {
                $sInfo = Get-FilesInLibrary -Context $sourceSubContext -LibraryTitle $slib.Title
                $tInfo = Get-FilesInLibrary -Context $targetSubContext -LibraryTitle $tlib.Title
                
                $status = if ($sInfo.FileCount -eq $tInfo.FileCount) { "OK" } else { "DIFF" }
                
                $results += [PSCustomObject]@{
                    Level = "Subsite: $($web.Title)"
                    Name = $slib.Title
                    Source_Files = $sInfo.FileCount
                    Target_Files = $tInfo.FileCount
                    Status = $status
                    Permissions = ""
                }
                Write-Host "    Lib '$($slib.Title)': $status (Q:$($sInfo.FileCount)/Z:$($tInfo.FileCount))"
            } else {
                Write-Host "    Lib '$($slib.Title)': FEHLT" -ForegroundColor Red
                $validationErrors += "Subsite '$($web.Title)' - Lib fehlt: $($slib.Title)"
            }
        } # Ende Foreach Libs
        
    }
    catch {
        Write-Host "  ✗ Subsite im Ziel nicht erreichbar/vorhanden" -ForegroundColor Red
        $validationErrors += "Subsite fehlt: $($web.Title)"
        $results += [PSCustomObject]@{ Level = "Subsite"; Name = $web.Title; Source_Files = "N/A"; Target_Files = "MISSING"; Status = "SITE_MISSING" }
    } # Ende Try/Catch

} # Ende Foreach Webs

# ===== ABSCHLUSS =====
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

Write-Host ""
Write-Host "Fertig. Report: $OutputPath" -ForegroundColor Cyan