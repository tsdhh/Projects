# SharePoint Subscription Edition - Berechtigungen auslesen - Basis Script
# Voraussetzung: Script muss auf dem SharePoint Server ausgeführt werden
# Als Administrator ausführen!

# Parameter definieren
$SiteUrl = "http://sharepoint-server/sites/DeinTeamraum"
# Optional: Ausgabe in Datei exportieren
$ExportToFile = $true
$OutputPath = "C:\Temp\SharePoint-Permissions.html"

# SharePoint Snap-In laden
try {
    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
    Write-Host "✓ SharePoint PowerShell erfolgreich geladen" -ForegroundColor Green
}
catch {
    Write-Host "✗ Fehler beim Laden von SharePoint PowerShell" -ForegroundColor Red
    Write-Host "  Bitte als Administrator ausführen!" -ForegroundColor Yellow
    exit
}

# Funktion: Berechtigungen mit Server Object Model abrufen
function Get-SitePermissions-ServerOM {
    param([string]$Url)
    
    Write-Host "`n=== Berechtigungen für: $Url ===" -ForegroundColor Green
    Write-Host "─" * 80
    
    # Array für strukturierte Daten (später für HTML-Export)
    $permissionsData = @()
    
    # SPWeb-Objekt abrufen
    $web = Get-SPWeb $Url
    
    try {
        # Alle Rollenzuweisungen durchgehen
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
                Members = @()
            }
            
            # Ausgabe formatieren
            Write-Host "`n├─ $($member.Name)" -ForegroundColor Yellow
            Write-Host "│  ├─ Typ: $($member.GetType().Name)" -ForegroundColor Gray
            
            if ($member.LoginName) {
                Write-Host "│  ├─ Login: $($member.LoginName)" -ForegroundColor Gray
            }
            
            Write-Host "│  └─ Berechtigungen: $($permissions -join ', ')" -ForegroundColor Cyan
            
            # Wenn es sich um eine Gruppe handelt, Mitglieder anzeigen
            if ($member -is [Microsoft.SharePoint.SPGroup]) {
                Write-Host "│     └─ Gruppenmitglieder:" -ForegroundColor Magenta
                
                foreach ($user in $member.Users) {
                    $userInfo = "├─ $($user.Name)"
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
            
            $permissionsData += $permissionEntry
        }
        
        Write-Host "`n" + ("─" * 80)
        
        return $permissionsData
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
        }
        .container {
            background-color: white;
            padding: 20px;
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #0078d4;
            border-bottom: 2px solid #0078d4;
            padding-bottom: 10px;
        }
        .tree {
            margin-top: 20px;
        }
        .tree-item {
            margin: 10px 0;
            padding: 10px;
            border-left: 3px solid #0078d4;
            background-color: #f8f9fa;
        }
        .tree-item-header {
            font-weight: bold;
            color: #323130;
            font-size: 16px;
        }
        .tree-item-detail {
            color: #605e5c;
            font-size: 14px;
            margin: 5px 0;
        }
        .permissions {
            color: #0078d4;
            font-weight: 500;
        }
        .members {
            margin-left: 20px;
            margin-top: 10px;
        }
        .member {
            padding: 5px;
            margin: 3px 0;
            background-color: white;
            border-left: 2px solid #00bcf2;
        }
        .icon {
            margin-right: 5px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 SharePoint Berechtigungsübersicht</h1>
        <div class="tree">
"@
    
    foreach ($item in $Data) {
        $html += @"
            <div class="tree-item">
                <div class="tree-item-header">
                    <span class="icon">👥</span>$($item.Name)
                </div>
                <div class="tree-item-detail">
                    Typ: $($item.Type)
                </div>
                <div class="tree-item-detail permissions">
                    🔐 Berechtigungen: $($item.Permissions)
                </div>
"@
        
        if ($item.Members.Count -gt 0) {
            $html += "<div class='members'><strong>Mitglieder:</strong>"
            foreach ($member in $item.Members) {
                $email = if ($member.Email) { " ($($member.Email))" } else { "" }
                $html += "<div class='member'>👤 $($member.Name)$email</div>"
            }
            $html += "</div>"
        }
        
        $html += "</div>"
    }
    
    $html += @"
        </div>
    </div>
</body>
</html>
"@
    
    # HTML-Datei speichern
    $html | Out-File -FilePath $OutputPath -Encoding UTF8
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