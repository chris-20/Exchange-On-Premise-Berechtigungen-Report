<#
.SYNOPSIS
    Erstellt einen Bericht über die Berechtigungen von Exchange-Postfächern.
.DESCRIPTION
    Dieses PowerShell-Skript sammelt Postfachberechtigungen (Postfachzugriffsrechte, "Send-As" und "Send on Behalf") für alle Benutzer- und freigegebene Postfächer in einem Exchange-Server und erstellt einen HTML-Report, der die Berechtigungen anzeigt.

.EXAMPLE
    PS> .\ExchangePermissionsReport.ps1
    (Erstellt einen Bericht über die Berechtigungen und speichert ihn als HTML-Datei.)

.LINK
    https://github.com/chris-20/Exchange-On-Premise-Berechtigungen-Report/

.NOTES
    Lizenz: MIT
    Version: 1.0
#>

# Funktion für Statusanzeige
function Write-ProgressStatus {
    param (
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete
    )
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
    Write-Host "[$($PercentComplete)%] $Activity - $Status" -ForegroundColor Cyan
}

# Exchange Server PowerShell Modul laden
Write-ProgressStatus -Activity "Initialisierung" -Status "Lade Exchange PowerShell Modul..." -PercentComplete 0
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue

# Datum fuer Report
$datum = Get-Date -Format "dd.MM.yyyy HH:mm"
$reportPfad = "Exchange_Berechtigungen_Report_$($datum.Replace(':', '-')).html"

# HTML-Styling
$htmlStyle = @"
<style>
    :root {
        --primary-gradient-start: #2B5876;
        --primary-gradient-mid: #3B4B76;
        --primary-gradient-end: #4E4376;
        --accent-gradient-start: #00c6ff;
        --accent-gradient-end: #0072ff;
        --background-color: #f8fafc;
        --card-background: #ffffff;
        --text-primary: #1e293b;
        --text-secondary: #64748b;
        --border-radius: 12px;
        --transition-smooth: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        --shadow-sm: 0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.24);
        --shadow-md: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -2px rgba(0,0,0,0.1);
        --shadow-lg: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -4px rgba(0,0,0,0.1);
    }

    body { 
        font-family: 'Segoe UI', system-ui, sans-serif;
        line-height: 1.6; 
        margin: 0; 
        padding: 32px; 
        background: linear-gradient(135deg, var(--background-color), #ffffff);
        color: var(--text-primary);
        min-height: 100vh;
    }

    .container { 
        max-width: 1200px; 
        margin: 0 auto; 
        background-color: var(--card-background); 
        padding: 32px;
        border-radius: var(--border-radius);
        box-shadow: var(--shadow-lg);
        border: 1px solid rgba(43, 88, 118, 0.08);
        position: relative;
        overflow: hidden;
    }

    .container::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 4px;
        background: linear-gradient(90deg, 
            var(--primary-gradient-start), 
            var(--primary-gradient-end));
        opacity: 0.8;
    }

    h1 { 
        color: var(--text-primary);
        font-size: 2.25rem;
        font-weight: 700;
        letter-spacing: -0.025em;
        margin-bottom: 2rem;
        padding-bottom: 1rem;
        background: linear-gradient(135deg, 
            var(--primary-gradient-start), 
            var(--primary-gradient-end));
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        position: relative;
    }

    h1::after {
        content: '';
        position: absolute;
        bottom: 0;
        left: 0;
        right: 0;
        height: 2px;
        background: linear-gradient(90deg, 
            var(--accent-gradient-start), 
            var(--accent-gradient-end));
        border-radius: 2px;
        opacity: 0.8;
    }

    h2 {
        color: var(--primary-gradient-start);
        font-size: 1.5rem;
        font-weight: 600;
        margin: 2rem 0 1.5rem;
        padding: 0.5rem 1rem;
        background: linear-gradient(135deg, 
            rgba(43, 88, 118, 0.08), 
            rgba(78, 67, 118, 0.08));
        border-radius: var(--border-radius);
        position: relative;
        overflow: hidden;
    }

    h2::before {
        content: '';
        position: absolute;
        left: 0;
        top: 0;
        bottom: 0;
        width: 4px;
        background: var(--primary-gradient-start);
        border-radius: 2px;
    }

    h3 {
        color: var(--text-primary);
        font-size: 1.25rem;
        font-weight: 500;
        margin: 1.5rem 0 1rem;
        padding: 0.75rem 1rem;
        background: rgba(43, 88, 118, 0.03);
        border-radius: var(--border-radius);
        border: 1px solid rgba(43, 88, 118, 0.06);
        transition: var(--transition-smooth);
    }

    h3:hover {
        transform: translateX(4px);
        border-color: var(--primary-gradient-start);
        background: rgba(43, 88, 118, 0.05);
    }

    .section {
        margin: 2rem 0;
        padding: 1.5rem;
        background: var(--card-background);
        border-radius: var(--border-radius);
        box-shadow: var(--shadow-md);
        border: 1px solid rgba(43, 88, 118, 0.08);
        transition: var(--transition-smooth);
    }

    .section:hover {
        box-shadow: var(--shadow-lg);
    }

    table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        margin: 1rem 0;
        background: var(--card-background);
        border-radius: var(--border-radius);
        overflow: hidden;
        box-shadow: var(--shadow-sm);
        border: 1px solid rgba(43, 88, 118, 0.08);
    }

    th {
        background: var(--primary-gradient-start);
        color: white;
        padding: 16px;
        text-align: left;
        font-weight: 500;
        font-size: 0.875rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }

    tr th:first-child {
        background: linear-gradient(90deg,
            var(--primary-gradient-start),
            var(--primary-gradient-start) 50%,
            var(--primary-gradient-end) 100%);
    }

    tr th:not(:first-child):not(:last-child) {
        background: var(--primary-gradient-end);
    }

    tr th:last-child {
        background: linear-gradient(90deg,
            var(--primary-gradient-end) 0%,
            var(--primary-gradient-start) 100%);
    }

    td {
        padding: 12px 16px;
        border-bottom: 1px solid rgba(43, 88, 118, 0.08);
        font-size: 0.875rem;
        color: var(--text-primary);
        transition: var(--transition-smooth);
    }

    tr:last-child td {
        border-bottom: none;
    }

    tr {
        transition: var(--transition-smooth);
    }

    tr:nth-child(even) {
        background: rgba(43, 88, 118, 0.02);
    }

    tr:hover {
        background: rgba(43, 88, 118, 0.05);
    }

    .timestamp {
        color: var(--text-secondary);
        font-size: 0.875rem;
        margin-top: 2rem;
        padding: 1rem;
        background: rgba(43, 88, 118, 0.02);
        border-radius: var(--border-radius);
        text-align: right;
        border: 1px solid rgba(43, 88, 118, 0.06);
    }

    @media (max-width: 768px) {
        body {
            padding: 16px;
        }

        .container {
            padding: 16px;
        }

        td, th {
            padding: 12px;
        }
    }
</style>
"@

$htmlHeader = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange Berechtigungen Report</title>
    <link rel="icon" type="image/svg+xml" href="data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNDAgMjQwIj48ZGVmcz48bGluZWFyR3JhZGllbnQgaWQ9InByaW1hcnlHcmFkaWVudCIgeDE9IjAlIiB5MT0iMCUiIHgyPSIxMDAlIiB5Mj0iMTAwJSI+PHN0b3Agb2Zmc2V0PSIwJSIgc3R5bGU9InN0b3AtY29sb3I6IzJCNTg3NiIvPjxzdG9wIG9mZnNldD0iMTAwJSIgc3R5bGU9InN0b3AtY29sb3I6IzRFNDM3NiIvPjwvbGluZWFyR3JhZGllbnQ+PGxpbmVhckdyYWRpZW50IGlkPSJhY2NlbnRHcmFkaWVudCIgeDE9IjAlIiB5MT0iMCUiIHgyPSIxMDAlIiB5Mj0iMCUiPjxzdG9wIG9mZnNldD0iMCUiIHN0eWxlPSJzdG9wLWNvbG9yOiMwMGM2ZmYiLz48c3RvcCBvZmZzZXQ9IjEwMCUiIHN0eWxlPSJzdG9wLWNvbG9yOiMwMDcyZmYiLz48L2xpbmVhckdyYWRpZW50PjwvZGVmcz48Y2lyY2xlIGN4PSIxMjAiIGN5PSIxMjAiIHI9IjExMCIgZmlsbD0idXJsKCNwcmltYXJ5R3JhZGllbnQpIi8+PHBhdGggZD0iTSA2MCwxMjAgTCA5MCwxMjAgTCAxMDUsNzAgTCAxMzUsMTcwIEwgMTUwLDEyMCBMIDE4MCwxMjAiIGZpbGw9Im5vbmUiIHN0cm9rZT0idXJsKCNhY2NlbnRHcmFkaWVudCkiIHN0cm9rZS13aWR0aD0iMTQiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIvPjwvc3ZnPg==">
    $htmlStyle
</head>
<body>
    <div class="container">
        <h1>Exchange Berechtigungen Report</h1>
        <div class="timestamp">Erstellt am: $datum</div>
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
Write-ProgressStatus -Activity "Report-Erstellung" -Status "Initialisiere Report..." -PercentComplete 10
$reportContent = ""

# Teil 1: Benutzerpostfuecher
Write-ProgressStatus -Activity "Report-Erstellung" -Status "Sammle Benutzerpostfächer..." -PercentComplete 20
$reportContent += "<div class='section'><h2>Benutzerpostfächer - Berechtigungen</h2>"
$alleBenutzerPostfaecher = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "UserMailbox"}
$totalMailboxCount = $alleBenutzerPostfaecher.Count
$currentMailbox = 0

foreach ($postfach in $alleBenutzerPostfaecher) {
    $currentMailbox++
    $percentComplete = 20 + (30 * ($currentMailbox / $totalMailboxCount))
    $statusMessage = "Verarbeite Postfach ${currentMailbox} von ${totalMailboxCount}: $($postfach.DisplayName)"
    Write-ProgressStatus -Activity "Benutzerpostfächer" -Status $statusMessage -PercentComplete $percentComplete
    
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
Write-ProgressStatus -Activity "Report-Erstellung" -Status "Sammle freigegebene Postfächer..." -PercentComplete 60
$reportContent += "<div class='section'><h2>Freigegebene Postfächer - Berechtigungen</h2>"
$alleFreigegebenenPostfaecher = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "SharedMailbox"}
$totalSharedMailboxCount = $alleFreigegebenenPostfaecher.Count
$currentSharedMailbox = 0

foreach ($postfach in $alleFreigegebenenPostfaecher) {
    $currentSharedMailbox++
    $percentComplete = 60 + (30 * ($currentSharedMailbox / $totalSharedMailboxCount))
    $statusMessage = "Verarbeite Postfach ${currentSharedMailbox} von ${totalSharedMailboxCount}: $($postfach.DisplayName)"
    Write-ProgressStatus -Activity "Freigegebene Postfächer" -Status $statusMessage -PercentComplete $percentComplete
    
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
    </div>
</body>
</html>
"@

# Gesamten HTML-Report zusammenstellen und speichern
Write-ProgressStatus -Activity "Report-Erstellung" -Status "Speichere Report..." -PercentComplete 90
$htmlReport = $htmlHeader + $reportContent + $htmlFooter
$htmlReport | Out-File -FilePath $reportPfad -Encoding UTF8

Write-ProgressStatus -Activity "Report-Erstellung" -Status "Fertig!" -PercentComplete 100
Write-Host "`nReport wurde erstellt unter: $reportPfad" -ForegroundColor Green
