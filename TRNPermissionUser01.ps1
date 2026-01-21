# SharePoint Subscription Edition - Berechtigungen auslesen - Basis Script
# Voraussetzung: Script muss auf dem SharePoint Server ausgeführt werden
# Als Administrator ausführen!

# UTF-8 Encoding für Konsole setzen (behebt Umlaut-Problem)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Parameter definieren
$SiteUrl = "http://sharepoint-server/sites/DeinTeamraum"
# Optional: Ausgabe in Datei exportieren
$ExportToFile = $true
$OutputPath = "C:\Temp\SharePoint-Permissions.html"

# Erweiterte Optionen
$IncludeLists = $true  # Listen/Bibliotheken einbeziehen
$OnlyUniquePermissions = $true  # Nur Listen mit eigenen Berechtigungen (nicht geerbt)
$AnalyzeLimitedAccess = $true  # Limited Access Berechtigungen detailliert analysieren
$ShowLimitedAccessReasons = $true  # Gründe für Limited Access anzeigen
$IncludeFolders = $true  # Ordner-Berechtigungen analysieren
$MaxFolderDepth = 3  # Maximale Ordner-Tiefe (Performance-Schutz)

# SharePoint Snap-In laden (nur wenn noch nicht geladen)
# Robuste Prüfung mit mehreren Methoden
$snapinLoaded = $false

# Methode 1: Prüfen ob Snap-In geladen ist
$snapin = Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
if ($null -ne $snapin) {
    $snapinLoaded = $true
}

# Methode 2: Prüfen ob SharePoint Cmdlets verfügbar sind
if (-not $snapinLoaded) {
    if (Get-Command Get-SPWeb -ErrorAction SilentlyContinue) {
        $snapinLoaded = $true
    }
}

# Snap-In laden falls nicht geladen
if (-not $snapinLoaded) {
    try {
        Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
        Write-Host "✓ SharePoint PowerShell erfolgreich geladen" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Fehler beim Laden von SharePoint PowerShell" -ForegroundColor Red
        Write-Host "  Fehlerdetails: $($_.Exception.Message)" -ForegroundColor DarkRed
        Write-Host "  Bitte als Administrator ausführen!" -ForegroundColor Yellow
        exit
    }
}
else {
    Write-Host "✓ SharePoint PowerShell bereits geladen" -ForegroundColor Cyan
}

# Funktion: Ordner mit eigenen Berechtigungen rekursiv finden
function Get-FoldersWithUniquePermissions {
    param(
        [Parameter(Mandatory=$true)]
        $List,
        [Parameter(Mandatory=$true)]
        $Web,
        [Parameter(Mandatory=$true)]
        [int]$Depth,
        [Parameter(Mandatory=$true)]
        [int]$MaxDepth,
        [Parameter(Mandatory=$true)]
        [string]$ParentPath
    )
    
    $foldersWithPerms = @()
    
    # Maximale Tiefe erreicht?
    if ($Depth -ge $MaxDepth) {
        return $foldersWithPerms
    }
    
    try {
        # Ordner in der aktuellen Ebene abrufen
        $folderQuery = New-Object Microsoft.SharePoint.SPQuery
        $folderQuery.Folder = $List.RootFolder
        if ($ParentPath) {
            $folderQuery.Folder = $Web.GetFolder("$($List.RootFolder.ServerRelativeUrl)/$ParentPath")
        }
        $folderQuery.ViewAttributes = "Scope='Recursive'"
        $folderQuery.Query = "<Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq></Where>"
        
        $folders = $List.GetItems($folderQuery)
        
        foreach ($folder in $folders) {
            # Nur Ordner mit eigenen Berechtigungen
            if ($folder.HasUniqueRoleAssignments) {
                $folderPath = $folder["FileRef"]
                $folderName = $folder["FileLeafRef"]
                $relativePath = $folderPath.Replace($List.RootFolder.ServerRelativeUrl + "/", "")
                
                $folderPerms = @()
                
                foreach ($roleAssignment in $folder.RoleAssignments) {
                    $member = $roleAssignment.Member
                    $permissions = @()
                    
                    foreach ($roleDefinition in $roleAssignment.RoleDefinitionBindings) {
                        $permissions += $roleDefinition.Name
                    }
                    
                    $folderPerms += [PSCustomObject]@{
                        Name = $member.Name
                        Type = $member.GetType().Name
                        Permissions = $permissions
                    }
                }
                
                $foldersWithPerms += [PSCustomObject]@{
                    Name = $folderName
                    Path = $relativePath
                    Depth = $Depth + 1
                    Permissions = $folderPerms
                }
            }
        }
    }
    catch {
        # Fehler bei Ordner-Zugriff ignorieren (z.B. keine Berechtigung)
    }
    
    return $foldersWithPerms
}

# Funktion: Berechtigungen mit Server Object Model abrufen
function Get-SitePermissions-ServerOM {
    param([string]$Url)
    
    Write-Host "`n╔══════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║  Berechtigungsanalyse für: $Url" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    
    # Strukturierte Daten für Export
    $siteData = [PSCustomObject]@{
        SiteUrl = $Url
        SitePermissions = @()
        Lists = @()
        LimitedAccessReasons = @()
    }
    
    # SPWeb-Objekt abrufen
    $web = Get-SPWeb $Url
    
    try {
        Write-Host "`n[WEBSITE-EBENE]" -ForegroundColor Magenta
        Write-Host "─" * 80
        
        # Website-Berechtigungen durchgehen
        foreach ($roleAssignment in $web.RoleAssignments) {
            $member = $roleAssignment.Member
            
            # Berechtigungsstufen sammeln
            $permissions = @()
            foreach ($roleDefinition in $roleAssignment.RoleDefinitionBindings) {
                $permissions += $roleDefinition.Name
            }
            
            # Daten für Export sammeln
            $permissionEntry = [PSCustomObject]@{
                Name = $member.Name
                Type = $member.GetType().Name
                LoginName = $member.LoginName
                Permissions = $permissions -join ', '
                IsLimitedAccess = ($permissions -contains 'Limited Access')
                Members = @()
            }
            
            # Konsolen-Ausgabe
            $nameDisplay = if ($permissionEntry.IsLimitedAccess) {
                "👥 $($member.Name) ⚠️"
            } else {
                "👥 $($member.Name)"
            }
            
            Write-Host "`n├─ $nameDisplay" -ForegroundColor Yellow
            Write-Host "│  ├─ Typ: $($member.GetType().Name)" -ForegroundColor Gray
            
            if ($member.LoginName) {
                Write-Host "│  ├─ Login: $($member.LoginName)" -ForegroundColor Gray
            }
            
            $permColor = if ($permissionEntry.IsLimitedAccess) { "DarkYellow" } else { "Cyan" }
            Write-Host "│  └─ 🔐 Berechtigungen: $($permissions -join ', ')" -ForegroundColor $permColor
            
            # Limited Access Erklärung anzeigen
            if ($permissionEntry.IsLimitedAccess -and $ShowLimitedAccessReasons) {
                Write-Host "│     ℹ️  Limited Access = Technischer Zugriff (automatisch vergeben)" -ForegroundColor DarkGray
                Write-Host "│     ℹ️  Ermöglicht Navigation zu Inhalten mit expliziten Rechten" -ForegroundColor DarkGray
            }
            
            # Gruppenmitglieder anzeigen
            if ($member -is [Microsoft.SharePoint.SPGroup]) {
                Write-Host "│     └─ Gruppenmitglieder:" -ForegroundColor Magenta
                
                foreach ($user in $member.Users) {
                    $userInfo = "├─ 👤 $($user.Name)"
                    if ($user.Email) {
                        $userInfo += " ($($user.Email))"
                    }
                    Write-Host "│        $userInfo" -ForegroundColor White
                    
                    # Für Export speichern
                    $permissionEntry.Members += [PSCustomObject]@{
                        Name = $user.Name
                        Email = $user.Email
                        LoginName = $user.LoginName
                    }
                }
            }
            
            $siteData.SitePermissions += $permissionEntry
        }
        
        # Listen/Bibliotheken analysieren
        if ($IncludeLists) {
            Write-Host "`n`n[LISTEN & BIBLIOTHEKEN]" -ForegroundColor Magenta
            Write-Host "─" * 80
            
            $lists = $web.Lists | Where-Object { -not $_.Hidden }
            $listCount = 0
            $limitedAccessReasons = @()
            
            foreach ($list in $lists) {
                # Prüfen ob Liste eigene Berechtigungen hat
                $hasUniquePermissions = $list.HasUniqueRoleAssignments
                
                # Nur Listen mit eigenen Berechtigungen anzeigen, wenn gewünscht
                if ($OnlyUniquePermissions -and -not $hasUniquePermissions) {
                    continue
                }
                
                $listCount++
                $inheritanceInfo = if ($hasUniquePermissions) { "🔒 Eigene Berechtigungen" } else { "🔓 Geerbt von Website" }
                
                Write-Host "`n┌─ 📚 $($list.Title) [$($list.BaseType)]" -ForegroundColor Green
                Write-Host "│  └─ $inheritanceInfo" -ForegroundColor $(if ($hasUniquePermissions) { "Yellow" } else { "Gray" })
                
                $listData = [PSCustomObject]@{
                    Title = $list.Title
                    BaseType = $list.BaseType.ToString()
                    HasUniquePermissions = $hasUniquePermissions
                    ItemCount = $list.ItemCount
                    Permissions = @()
                    Folders = @()
                }
                
                # Wenn eigene Berechtigungen, diese anzeigen
                if ($hasUniquePermissions) {
                    foreach ($roleAssignment in $list.RoleAssignments) {
                        $member = $roleAssignment.Member
                        
                        $permissions = @()
                        foreach ($roleDefinition in $roleAssignment.RoleDefinitionBindings) {
                            $permissions += $roleDefinition.Name
                        }
                        
                        $isLimitedAccess = $permissions -contains 'Limited Access'
                        
                        # Grund für Limited Access sammeln
                        if ($isLimitedAccess -and $AnalyzeLimitedAccess) {
                            $limitedAccessEntry = [PSCustomObject]@{
                                Member = $member.Name
                                Location = "Liste: $($list.Title)"
                                Reason = "Hat vermutlich Rechte auf Ordner/Items in dieser Liste"
                            }
                            $limitedAccessReasons += $limitedAccessEntry
                            $siteData.LimitedAccessReasons += $limitedAccessEntry
                        }
                        
                        $memberDisplay = if ($isLimitedAccess) { "👥 $($member.Name) ⚠️" } else { "👥 $($member.Name)" }
                        $permColor = if ($isLimitedAccess) { "DarkYellow" } else { "Cyan" }
                        
                        Write-Host "│  ├─ $memberDisplay" -ForegroundColor Yellow
                        Write-Host "│  │  └─ 🔐 $($permissions -join ', ')" -ForegroundColor $permColor
                        
                        $listPermEntry = [PSCustomObject]@{
                            Name = $member.Name
                            Type = $member.GetType().Name
                            Permissions = $permissions -join ', '
                            IsLimitedAccess = $isLimitedAccess
                            Members = @()
                        }
                        
                        # Gruppenmitglieder
                        if ($member -is [Microsoft.SharePoint.SPGroup]) {
                            foreach ($user in $member.Users) {
                                Write-Host "│  │     ├─ 👤 $($user.Name)" -ForegroundColor White
                                
                                $listPermEntry.Members += [PSCustomObject]@{
                                    Name = $user.Name
                                    Email = $user.Email
                                }
                            }
                        }
                        
                        $listData.Permissions += $listPermEntry
                    }
                    
                    # Ordner in Bibliotheken analysieren (nur DocumentLibrary)
                    if ($IncludeFolders -and $list.BaseType -eq "DocumentLibrary") {
                        Write-Host "│  └─ 📁 Analysiere Ordner..." -ForegroundColor DarkCyan
                        
                        try {
                            $folders = Get-FoldersWithUniquePermissions -List $list -Web $web -Depth 0 -MaxDepth $MaxFolderDepth -ParentPath ""
                            
                            foreach ($folderInfo in $folders) {
                                $indent = "│  " + ("   " * $folderInfo.Depth)
                                Write-Host "$indent├─ 📂 $($folderInfo.Name) 🔒" -ForegroundColor Cyan
                                
                                foreach ($perm in $folderInfo.Permissions) {
                                    $isLA = $perm.Permissions -contains 'Limited Access'
                                    $permColor = if ($isLA) { "DarkYellow" } else { "Green" }
                                    Write-Host "$indent│  └─ 👥 $($perm.Name): $($perm.Permissions -join ', ')" -ForegroundColor $permColor
                                    
                                    # Limited Access für Ordner tracken
                                    if ($isLA -and $AnalyzeLimitedAccess) {
                                        $limitedAccessEntry = [PSCustomObject]@{
                                            Member = $perm.Name
                                            Location = "Ordner: $($list.Title)/$($folderInfo.Path)"
                                            Reason = "Hat vermutlich Rechte auf Dateien in diesem Ordner"
                                        }
                                        $limitedAccessReasons += $limitedAccessEntry
                                        $siteData.LimitedAccessReasons += $limitedAccessEntry
                                    }
                                }
                                
                                $listData.Folders += $folderInfo
                            }
                            
                            if ($folders.Count -gt 0) {
                                Write-Host "$indent└─ ✓ $($folders.Count) Ordner mit eigenen Berechtigungen gefunden" -ForegroundColor DarkGray
                            }
                        }
                        catch {
                            Write-Host "│     └─ ⚠️ Fehler bei Ordner-Analyse: $($_.Exception.Message)" -ForegroundColor DarkRed
                        }
                    }
                }
                
                $siteData.Lists += $listData
            }
            
            Write-Host "`n"
            Write-Host "─" * 80
            Write-Host "📊 Statistik: $listCount Listen/Bibliotheken mit $(if($OnlyUniquePermissions){'eigenen'}else{'allen'}) Berechtigungen analysiert" -ForegroundColor Cyan
            
            # Limited Access Zusammenfassung anzeigen
            if ($AnalyzeLimitedAccess -and $limitedAccessReasons.Count -gt 0) {
                Write-Host "`n"
                Write-Host "═" * 80 -ForegroundColor Yellow
                Write-Host "⚠️  LIMITED ACCESS ANALYSE" -ForegroundColor Yellow
                Write-Host "═" * 80 -ForegroundColor Yellow
                Write-Host "`nFolgende Benutzer/Gruppen haben Limited Access:" -ForegroundColor Cyan
                
                $groupedReasons = $limitedAccessReasons | Group-Object -Property Member
                
                foreach ($group in $groupedReasons) {
                    Write-Host "`n👤 $($group.Name)" -ForegroundColor Yellow
                    foreach ($reason in $group.Group) {
                        Write-Host "   └─ $($reason.Location)" -ForegroundColor Gray
                        Write-Host "      └─ $($reason.Reason)" -ForegroundColor DarkGray
                    }
                }
                
                Write-Host "`n💡 HINWEIS:" -ForegroundColor Cyan
                Write-Host "   Limited Access ist NORMAL und wird automatisch von SharePoint vergeben." -ForegroundColor Gray
                Write-Host "   Es bedeutet: Der Benutzer hat auf untergeordnete Elemente (Ordner/Dateien)" -ForegroundColor Gray
                Write-Host "   explizite Berechtigungen, benötigt aber technischen Zugriff auf die" -ForegroundColor Gray
                Write-Host "   übergeordnete Struktur, um dorthin zu navigieren." -ForegroundColor Gray
                Write-Host "═" * 80 -ForegroundColor Yellow
            }
        }
        
        Write-Host "`n" + ("═" * 80)
        
        return $siteData
    }
    finally {
        # SPWeb-Objekt freigeben
        $web.Dispose()
    }
}

# Funktion: Export als HTML (Vorschau für Webpart-Design)
function Export-PermissionsToHTML {
    param(
        [Parameter(Mandatory=$true)]
        $Data,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>SharePoint Berechtigungen - Baumansicht</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            padding: 20px;
            background-color: #f3f2f1;
            margin: 0;
        }
        .container {
            background-color: white;
            padding: 20px;
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            max-width: 1400px;
            margin: 0 auto;
        }
        h1 {
            color: #0078d4;
            border-bottom: 3px solid #0078d4;
            padding-bottom: 10px;
            margin-bottom: 10px;
        }
        h2 {
            color: #005a9e;
            margin-top: 30px;
            padding: 10px;
            background-color: #f3f2f1;
            border-left: 4px solid #0078d4;
        }
        .site-info {
            background-color: #e1f5ff;
            padding: 10px;
            border-radius: 4px;
            margin-bottom: 20px;
            font-size: 14px;
        }
        .tree {
            margin-top: 20px;
        }
        .tree-item {
            margin: 10px 0;
            padding: 15px;
            border-left: 3px solid #0078d4;
            background-color: #faf9f8;
            transition: all 0.2s;
        }
        .tree-item:hover {
            background-color: #f3f2f1;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .tree-item-header {
            font-weight: bold;
            color: #323130;
            font-size: 16px;
            margin-bottom: 8px;
        }
        .tree-item-detail {
            color: #605e5c;
            font-size: 14px;
            margin: 5px 0;
        }
        .permissions {
            color: #0078d4;
            font-weight: 500;
            background-color: #e1f5ff;
            padding: 5px 10px;
            border-radius: 3px;
            display: inline-block;
            margin-top: 5px;
        }
        .members {
            margin-left: 20px;
            margin-top: 10px;
            border-left: 2px solid #00bcf2;
            padding-left: 15px;
        }
        .member {
            padding: 8px;
            margin: 5px 0;
            background-color: white;
            border-radius: 3px;
            border-left: 3px solid #00bcf2;
        }
        .list-item {
            margin: 15px 0;
            padding: 15px;
            border-left: 4px solid #107c10;
            background-color: #f0fdf4;
        }
        .list-item.inherited {
            border-left-color: #8a8886;
            background-color: #f8f9fa;
            opacity: 0.8;
        }
        .list-header {
            font-weight: bold;
            font-size: 16px;
            color: #107c10;
            margin-bottom: 10px;
        }
        .list-item.inherited .list-header {
            color: #605e5c;
        }
        .badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 500;
            margin-left: 10px;
        }
        .badge-unique {
            background-color: #fde047;
            color: #854d0e;
        }
        .badge-inherited {
            background-color: #e5e7eb;
            color: #6b7280;
        }
        .stats {
            background-color: #f3f2f1;
            padding: 15px;
            border-radius: 4px;
            margin-top: 20px;
            border-left: 4px solid #00bcf2;
        }
        .icon {
            margin-right: 8px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 SharePoint Berechtigungsübersicht</h1>
        <div class="site-info">
            <strong>🌐 Website:</strong> $($Data.SiteUrl)
        </div>
        
        <h2>🏢 Website-Berechtigungen</h2>
        <div class="tree">
"@
    
    # Website-Berechtigungen
    foreach ($item in $Data.SitePermissions) {
        $html += @"
            <div class="tree-item">
                <div class="tree-item-header">
                    <span class="icon">👥</span>$($item.Name)
                </div>
                <div class="tree-item-detail">
                    Typ: $($item.Type)
                </div>
                <div class="permissions">
                    🔐 $($item.Permissions)
                </div>
"@
        
        if ($item.Members.Count -gt 0) {
            $html += "<div class='members'><strong>Mitglieder ($($item.Members.Count)):</strong>"
            foreach ($member in $item.Members) {
                $email = if ($member.Email) { " ($($member.Email))" } else { "" }
                $html += "<div class='member'>👤 $($member.Name)$email</div>"
            }
            $html += "</div>"
        }
        
        $html += "</div>"
    }
    
    $html += "</div>"
    
    # Listen/Bibliotheken
    if ($Data.Lists.Count -gt 0) {
        $html += "<h2>📚 Listen & Bibliotheken</h2>"
        
        foreach ($list in $Data.Lists) {
            $isInherited = -not $list.HasUniquePermissions
            $cssClass = if ($isInherited) { "list-item inherited" } else { "list-item" }
            $badge = if ($isInherited) { 
                "<span class='badge badge-inherited'>🔓 Geerbt</span>" 
            } else { 
                "<span class='badge badge-unique'>🔒 Eigene Berechtigungen</span>" 
            }
            
            $html += @"
            <div class="$cssClass">
                <div class="list-header">
                    📑 $($list.Title) $badge
                </div>
                <div class="tree-item-detail">
                    Typ: $($list.BaseType) | Elemente: $($list.ItemCount)
                </div>
"@
            
            if ($list.Permissions.Count -gt 0) {
                $html += "<div style='margin-top: 10px;'>"
                foreach ($perm in $list.Permissions) {
                    $isLA = $perm.IsLimitedAccess
                    $laWarning = if ($isLA) { " ⚠️" } else { "" }
                    $permClass = if ($isLA) { "permissions" } else { "permissions" }
                    
                    $html += @"
                    <div class="tree-item" style="margin: 10px 0;">
                        <div class="tree-item-header">👥 $($perm.Name)$laWarning</div>
                        <div class="$permClass">🔐 $($perm.Permissions)</div>
"@
                    
                    if ($perm.Members.Count -gt 0) {
                        $html += "<div class='members'>"
                        foreach ($member in $perm.Members) {
                            $email = if ($member.Email) { " ($($member.Email))" } else { "" }
                            $html += "<div class='member'>👤 $($member.Name)$email</div>"
                        }
                        $html += "</div>"
                    }
                    
                    $html += "</div>"
                }
                $html += "</div>"
            }
            
            # Ordner anzeigen
            if ($list.Folders.Count -gt 0) {
                $html += "<div style='margin-top: 15px; padding: 10px; background-color: #e8f4f8; border-left: 4px solid #0078d4; border-radius: 4px;'>"
                $html += "<strong>📁 Ordner mit eigenen Berechtigungen ($($list.Folders.Count)):</strong>"
                
                foreach ($folder in $list.Folders) {
                    $indent = "margin-left: " + ($folder.Depth * 20) + "px;"
                    $html += @"
                    <div style="$indent margin-top: 10px; padding: 8px; background-color: white; border-left: 3px solid #00bcf2; border-radius: 3px;">
                        <div style="font-weight: bold; color: #0078d4;">📂 $($folder.Name)</div>
                        <div style="font-size: 12px; color: #605e5c;">Pfad: $($folder.Path)</div>
"@
                    
                    foreach ($perm in $folder.Permissions) {
                        $isLA = $perm.Permissions -contains 'Limited Access'
                        $laWarning = if ($isLA) { " ⚠️" } else { "" }
                        $permColor = if ($isLA) { "#d97706" } else { "#0078d4" }
                        $permName = $perm.Name
                        $permList = $perm.Permissions -join ', '
                        
                        $html += "<div style='margin-top: 5px; padding: 5px; background-color: #f8f9fa;'>"
                        $html += "<span style='color: #323130;'>👥 ${permName}${laWarning}:</span>"
                        $html += "<span style='color: $permColor; font-weight: 500;'> $permList</span>"
                        $html += "</div>"
                    }
                    
                    $html += "</div>"
                }
                
                $html += "</div>"
            }
            
            $html += "</div>"
        }
        
        # Statistik
        $uniqueCount = ($Data.Lists | Where-Object { $_.HasUniquePermissions }).Count
        $inheritedCount = ($Data.Lists | Where-Object { -not $_.HasUniquePermissions }).Count
        $folderCount = ($Data.Lists.Folders | Measure-Object).Count
        
        $html += @"
        <div class="stats">
            <strong>📊 Statistik:</strong><br>
            Gesamt: $($Data.Lists.Count) Listen/Bibliotheken |
            🔒 Eigene Berechtigungen: $uniqueCount |
            🔓 Geerbt: $inheritedCount |
            📁 Ordner mit Berechtigungen: $folderCount
        </div>
"@
    }
    
    # Limited Access Analyse
    if ($Data.LimitedAccessReasons.Count -gt 0) {
        $html += @"
        <h2 style="color: #d97706; border-left-color: #d97706;">⚠️ Limited Access Analyse</h2>
        <div style="background-color: #fffbeb; padding: 15px; border-left: 4px solid #d97706; border-radius: 4px; margin-bottom: 15px;">
            <strong style="color: #d97706;">ℹ️ Was ist Limited Access?</strong><br>
            <span style="color: #78716c;">Limited Access wird automatisch von SharePoint vergeben und gibt KEINE direkten Berechtigungen auf Inhalte. 
            Es ermöglicht lediglich den technischen Zugriff auf übergeordnete Strukturen, damit Benutzer zu Inhalten navigieren können, 
            für die sie explizite Berechtigungen haben.</span>
        </div>
"@
        
        # Gruppiere nach Benutzer
        $groupedReasons = $Data.LimitedAccessReasons | Group-Object -Property Member
        
        foreach ($group in $groupedReasons) {
            $html += @"
            <div class="tree-item" style="border-left-color: #d97706; background-color: #fffbeb;">
                <div class="tree-item-header" style="color: #d97706;">👤 $($group.Name)</div>
                <div style="margin-top: 10px;">
                    <strong>Orte mit Limited Access:</strong>
"@
            
            foreach ($reason in $group.Group) {
                $html += @"
                    <div style="margin: 8px 0; padding: 8px; background-color: white; border-left: 3px solid #fbbf24; border-radius: 3px;">
                        <div style="color: #0078d4; font-weight: 500;">📍 $($reason.Location)</div>
                        <div style="color: #78716c; font-size: 14px; margin-top: 3px;">→ $($reason.Reason)</div>
                    </div>
"@
            }
            
            $html += @"
                </div>
            </div>
"@
        }
    }
    
    $html += @"
    </div>
</body>
</html>
"@
    
    # UTF-8 ohne BOM für Datei
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($OutputPath, $html, $utf8NoBom)
}

# Funktion: Berechtigungen mit Client Object Model abrufen (CSOM)
function Get-SitePermissions-CSOM {
    param([string]$Url)
    
    Write-Host "`n=== Berechtigungen für: $Url ===" -ForegroundColor Green
    Write-Host "─" * 80
    
    # CSOM Assemblies laden
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    
    # Kontext erstellen
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    
    # Windows-Authentifizierung verwenden
    $ctx.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    
    try {
        # Web-Objekt abrufen
        $web = $ctx.Web
        $ctx.Load($web)
        $ctx.Load($web.RoleAssignments)
        $ctx.ExecuteQuery()
        
        # Alle Rollenzuweisungen durchgehen
        foreach ($roleAssignment in $web.RoleAssignments) {
            $ctx.Load($roleAssignment.Member)
            $ctx.Load($roleAssignment.RoleDefinitionBindings)
            $ctx.ExecuteQuery()
            
            $member = $roleAssignment.Member
            
            # Berechtigungsstufen sammeln
            $permissions = @()
            foreach ($roleDefinition in $roleAssignment.RoleDefinitionBindings) {
                $permissions += $roleDefinition.Name
            }
            
            # Ausgabe formatieren
            Write-Host "`n├─ $($member.Title)" -ForegroundColor Yellow
            Write-Host "│  ├─ Typ: $($member.PrincipalType)" -ForegroundColor Gray
            Write-Host "│  ├─ Login: $($member.LoginName)" -ForegroundColor Gray
            Write-Host "│  └─ Berechtigungen: $($permissions -join ', ')" -ForegroundColor Cyan
            
            # Wenn es sich um eine Gruppe handelt, Mitglieder anzeigen
            if ($member.PrincipalType -eq "SharePointGroup") {
                try {
                    $group = $ctx.Web.SiteGroups.GetById($member.Id)
                    $ctx.Load($group.Users)
                    $ctx.ExecuteQuery()
                    
                    if ($group.Users.Count -gt 0) {
                        Write-Host "│     └─ Gruppenmitglieder:" -ForegroundColor Magenta
                        
                        foreach ($user in $group.Users) {
                            $userInfo = "├─ $($user.Title)"
                            if ($user.Email) {
                                $userInfo += " ($($user.Email))"
                            }
                            Write-Host "│        $userInfo" -ForegroundColor White
                        }
                    }
                }
                catch {
                    Write-Host "│     └─ (Mitglieder konnten nicht abgerufen werden)" -ForegroundColor DarkGray
                }
            }
        }
        
        Write-Host "`n" + ("─" * 80)
    }
    finally {
        $ctx.Dispose()
    }
}

# Hauptausführung
try {
    # Server Object Model verwenden
    $permissionsData = Get-SitePermissions-ServerOM -Url $SiteUrl
    
    Write-Host "`n✓ Script erfolgreich ausgeführt!" -ForegroundColor Green
    
    # Optional: HTML-Export für Webpart-Vorschau
    if ($ExportToFile) {
        Export-PermissionsToHTML -Data $permissionsData -OutputPath $OutputPath
        Write-Host "✓ HTML-Datei erstellt: $OutputPath" -ForegroundColor Green
        Write-Host "  Diese kannst du im Browser öffnen, um zu sehen, wie es im Webpart aussehen könnte." -ForegroundColor Cyan
    }
}
catch {
    Write-Host "`n✗ Fehler: $_" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor DarkRed
}

# Zusätzliche Informationen
Write-Host "`n📌 Erklärung der Berechtigungsstufen:" -ForegroundColor Cyan
Write-Host "   • Vollzugriff (Full Control): Alle Rechte"
Write-Host "   • Entwerfen (Design): Kann Listen und Seiten erstellen/ändern"
Write-Host "   • Bearbeiten (Edit): Kann Elemente hinzufügen/bearbeiten/löschen"
Write-Host "   • Mitwirken (Contribute): Kann Elemente hinzufügen/bearbeiten"
Write-Host "   • Lesen (Read): Nur Lesezugriff"
Write-Host "   • Eingeschränkter Lesezugriff (Limited Access): Minimaler Zugriff"