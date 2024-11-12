📊 **Exchange-On-Premise-Berechtigungen-Report**

Dieses PowerShell-Skript erstellt einen umfassenden Berechtigungsbericht für Benutzer- und freigegebene Postfächer auf einem Exchange Server. Es erfasst detailliert Postfachberechtigungen, Send-As- und Send-on-Behalf-Berechtigungen und erstellt eine übersichtliche HTML-Datei. Der Report dient dazu, Zugriffsbeschränkungen zu prüfen, Sicherheitsanforderungen einzuhalten und die Übersicht zu wahren.

✨ **Funktionen**

- 🔎 **Detaillierte Berechtigungsberichte**: Alle relevanten Berechtigungen werden für jedes Postfach strukturiert angezeigt.
- 📄 **Stilvolle HTML-Darstellung**: Der Report nutzt klares HTML und CSS für eine lesbare und moderne Darstellung.
- 📅 **Automatisierter Zeitstempel**: Jeder Bericht enthält das Erstellungsdatum zur Nachverfolgung.
- 👥 **Benutzer- und freigegebene Postfächer**: Der Bericht listet die Berechtigungen getrennt für unterschiedliche Postfachtypen auf.

📋 **Voraussetzungen**

- Zugriff auf den Exchange On-Premise Server
- Berechtigungen zur Nutzung der Exchange Server PowerShell-Module

🚀 **Verwendung**

1. Lade das Skript herunter.
2. Öffne die **Exchange Management Shell**.
3. Navigiere in das Verzeichnis, in dem sich das Skript befindet.
4. Gib folgenden Befehl ein, um das Skript auszuführen:  
   `powershell -ExecutionPolicy Bypass -File .\Exchange-Berechtigungen-Report.ps1`
5. Der Bericht wird als HTML-Datei im gleichen Verzeichnis gespeichert und enthält einen Zeitstempel im Dateinamen.

📘 **Beispielausgabe**

Der HTML-Bericht zeigt:

- 🛠 **Berechtigungstyp** (Postfachzugriff, Send-As, Send-on-Behalf)
- 🧾 **Spezifische Berechtigungen**
- 👤 **Zugewiesene Benutzer** und deren Berechtigungen

