## 📊 Exchange-Berechtigungen-Report

Dieses PowerShell-Skript erstellt einen umfassenden Berechtigungsbericht für Benutzer- und freigegebene Postfächer auf einem Exchange Server. Es erfasst detailliert Postfachberechtigungen, `Send-As`- und `Send-on-Behalf`-Berechtigungen und erstellt eine übersichtliche HTML-Datei. Der Report dient IT-Administratoren als wertvolle Hilfe, um Zugriffsbeschränkungen zu prüfen, Sicherheitsanforderungen einzuhalten und die Übersicht zu wahren.

### ✨ Funktionen

- **🔎 Detaillierte Berechtigungsberichte**: Alle relevanten Berechtigungen werden für jedes Postfach strukturiert angezeigt.
- **📄 Stilvolle HTML-Darstellung**: Der Report nutzt klares HTML und CSS für eine lesbare und moderne Darstellung.
- **📅 Automatisierter Zeitstempel**: Jeder Bericht enthält das Erstellungsdatum zur Nachverfolgung.
- **👥 Benutzer- und freigegebene Postfächer**: Der Bericht listet die Berechtigungen getrennt für unterschiedliche Postfachtypen.

### 📋 Voraussetzungen

- PowerShell-Zugriff auf den Exchange Server
- Berechtigungen zur Nutzung der Exchange Server PowerShell-Module

### 🚀 Verwendung

1. Führen Sie das Skript auf dem Exchange Server mit den erforderlichen Administratorrechten aus.
2. Der Bericht wird als HTML-Datei im gleichen Verzeichnis gespeichert und enthält im Dateinamen das Erstellungsdatum.

### 📘 Beispielausgabe

Der HTML-Bericht zeigt:
- 🛠 **Berechtigungstyp** (Postfachzugriff, `Send-As`, `Send-on-Behalf`)
- 🧾 **Spezifische Berechtigungen**
- 👤 **Zugewiesener Benutzer**

