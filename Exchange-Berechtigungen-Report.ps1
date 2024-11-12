<#
.SYNOPSIS
    Erstellt einen Bericht über die Berechtigungen von Exchange-Postfächern.
.DESCRIPTION
    Dieses PowerShell-Skript sammelt Postfachberechtigungen (Postfachzugriffsrechte, "Send-As" und "Send on Behalf") für alle Benutzer- und freigegebene Postfächer in einem Exchange-Server und erstellt einen HTML-Report, der die Berechtigungen anzeigt.

.EXAMPLE
    PS> .\ExchangePermissionsReport.ps1
    (Erstellt einen Bericht über die Berechtigungen und speichert ihn als HTML-Datei.)

.LINK
    https://github.com/chris-20/Exchange-Berechtigungen-Report

.NOTES
    Lizenz: MIT
#>

# Exchange Server PowerShell Modul laden
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue

# Datum fuer Report
$datum = Get-Date -Format "dd.MM.yyyy HH:mm"
$reportPfad = "Exchange_Berechtigungen_Report_$($datum.Replace(':', '-')).html"

# HTML-Styling
$htmlStyle = @"
<style>
    body { font-family: Arial, sans-serif; margin: 20px; }
    h1, h2 { color: #2c3e50; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
    th { background-color: #2c3e50; color: white; padding: 10px; text-align: left; }
    td { padding: 8px; border: 1px solid #ddd; }
    tr:nth-child(even) { background-color: #f2f2f2; }
    tr:hover { background-color: #e9e9e9; }
    .section { margin-bottom: 30px; }
</style>
"@

# HTML-Header
$htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Exchange Berechtigungen Report</title>
    $htmlStyle
</head>
<body>
    <h1>Exchange Berechtigungen Report</h1>
    <p>Erstellt am: $datum</p>
"@

# Funktion zum Abrufen der Berechtigungen fuer ein Postfach
function Get-MailboxPermissions {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Mailbox
    )
    
    $permissions = @()
    
    # Postfachberechtigungen
    Get-MailboxPermission -Identity $Mailbox | 
        Where-Object {$_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-5-*" -and $_.IsInherited -eq $false} |
        ForEach-Object {
            $permissions += [PSCustomObject]@{
                Typ = "Postfachberechtigung"
                Berechtigung = $_.AccessRights -join ", "
                Benutzer = $_.User
            }
        }
    
    # Send-As Berechtigungen
    Get-ADPermission -Identity $Mailbox | 
        Where-Object {$_.ExtendedRights -like "*Send-As*" -and $_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-5-*" -and $_.IsInherited -eq $false} |
        ForEach-Object {
            $permissions += [PSCustomObject]@{
                Typ = "Send-As"
                Berechtigung = "Send As"
                Benutzer = $_.User
            }
        }
    
    # Stellvertretungen
    $sendOnBehalf = Get-Mailbox -Identity $Mailbox | Select-Object -ExpandProperty GrantSendOnBehalfTo
    if ($sendOnBehalf) {
        foreach ($delegate in $sendOnBehalf) {
            $permissions += [PSCustomObject]@{
                Typ = "Send on Behalf"
                Berechtigung = "Send on Behalf"
                Benutzer = $delegate
            }
        }
    }
    
    return $permissions
}

# Hauptteil des Reports
$reportContent = ""

# Teil 1: Benutzerpostfuecher
$reportContent += "<div class='section'><h2>Benutzerpostfuecher - Berechtigungen</h2>"
$alleBenutzerPostfaecher = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "UserMailbox"}

foreach ($postfach in $alleBenutzerPostfaecher) {
    $berechtigungen = Get-MailboxPermissions -Mailbox $postfach.Identity
    
    if ($berechtigungen) {
        $reportContent += "<h3>Postfach: $($postfach.DisplayName) ($($postfach.PrimarySmtpAddress))</h3>"
        $reportContent += "<table><tr><th>Berechtigungstyp</th><th>Berechtigung</th><th>Berechtigter Benutzer</th></tr>"
        
        foreach ($berechtigung in $berechtigungen) {
            $reportContent += "<tr><td>$($berechtigung.Typ)</td><td>$($berechtigung.Berechtigung)</td><td>$($berechtigung.Benutzer)</td></tr>"
        }
        
        $reportContent += "</table>"
    }
}

# Teil 2: Freigegebene Postfuecher
$reportContent += "<div class='section'><h2>Freigegebene Postfaecher - Berechtigungen</h2>"
$alleFreigegebenenPostfaecher = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "SharedMailbox"}

foreach ($postfach in $alleFreigegebenenPostfaecher) {
    $berechtigungen = Get-MailboxPermissions -Mailbox $postfach.Identity
    
    if ($berechtigungen) {
        $reportContent += "<h3>Freigegebenes Postfach: $($postfach.DisplayName) ($($postfach.PrimarySmtpAddress))</h3>"
        $reportContent += "<table><tr><th>Berechtigungstyp</th><th>Berechtigung</th><th>Berechtigter Benutzer</th></tr>"
        
        foreach ($berechtigung in $berechtigungen) {
            $reportContent += "<tr><td>$($berechtigung.Typ)</td><td>$($berechtigung.Berechtigung)</td><td>$($berechtigung.Benutzer)</td></tr>"
        }
        
        $reportContent += "</table>"
    }
}

# HTML Footer
$htmlFooter = @"
</body>
</html>
"@

# Gesamten HTML-Report zusammenstellen und speichern
$htmlReport = $htmlHeader + $reportContent + $htmlFooter
$htmlReport | Out-File -FilePath $reportPfad -Encoding UTF8

Write-Host "Report wurde erstellt unter: $reportPfad"
