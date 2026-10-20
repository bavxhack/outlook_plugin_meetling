# Prüfbericht – Outlook-Plugin-API

Basis: GitHub main 0ca8de3d9741af5378dde2989239e78e551b5066, am
20.10.2026 über GitHub-Connector verifiziert. Lokaler Ausgangscheckout war älter
und wurde vor Implementierung mit dieser Basis abgeglichen. Shell-Netzwerkproxy
ist in der bereitgestellten Linux-Umgebung nicht erreichbar; .NET/Windows/Outlook
stehen lokal nicht zur Verfügung. Builds und Ausführung erfolgen in Windows-CI.

## Automatisierte Prüfungen

Kern-Testläufer prüft URLs/Unterpfade und Pfadabwehr, GET/POST-Konvention,
unterschiedliche Tokenheader, ausschließlich server_id im Formular, alle fünf
fachlichen Fehler bei HTTP 200, ungültiges JSON/HTML/leere Antworten, Schemas,
0/1/2/4 Server, gleiche Namen und getrennte IDs, große numerische IDs ohne
Rundung, fehlerhafte/duplizierte IDs, Login-Endpointbindung, Tokenvalidierung,
Konferenz-URLs unverändert, Logout-Schemata/Fehler, HTTP-Status und keine Retries,
verwaltete Priorität/Sperrwirkung/kaputte Datei ohne Fallback, Session-Epochen.

Windows-Standardbenutzer-Smoke prüft COM-Aktivierung, erhaltene DISPIDs 1–4,
zusätzliche DISPIDs 5–13, IDispatch-Aufrufgrenze, Logo/IPictureDisp, Ribbon im
Explorer, Termin und Mail-Compose, Lifecycle, DPAPI-Roundtrip/Neustart,
Installationstokenbindung und Löschen, installierbare Pakete für x86/x64,
Neuinstallation und Deinstallation. Die Windows-Runner ersetzen kein Outlook.

Ausführungsergebnisse: siehe verlinkten Actions-Lauf des PR. Ein Build gilt erst
mit grünen Tests und beiden hochgeladenen Architektur-Artefakten als geliefert.

## Nicht ausgeführte Live-Abnahme

Alle Punkte in docs/smoke-test.md benötigen einen Windows-Outlook-Testplatz mit
realer Meetling-/SSO-Installation. SSO-Rückkehr und Moderatoridentität sind offen.
Keine simulierte Antwort wird als produktive Integration bezeichnet. Auch
native Word-Editor-Einfügung, Feld-/Formatierungserhalt und Outlook-Bitnessen
müssen real geprüft werden. Keine Zugangsdaten oder Testinstanz bereitgestellt.
