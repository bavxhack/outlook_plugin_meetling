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

## Rückmeldung aus Live-Versuch (20.10.2026)

Nutzer meldet: /outlooklogin liefert nach OTP im normalen Edge die Token-JSON-
Antwort; im Add-in bleibt die OTP-Seite stehen. Die konkrete Ursache ist ohne
Browserablauf/Diagnose noch unbestätigt. Im Client wurde eine pauschale Sperre
zusätzlicher SSO-Fenster gefunden. Ergänzt wurden SSO-Popups im gleichen
WebView2-Kontext, sichere Navigationsdiagnose und eine explizite Rückkehr zum
konfigurierten Endpunkt. Diese Änderungen sind kein Nachweis einer behobenen
Live-OTP-Störung. Erneuter Versuch mit neuem Paket erforderlich.

## Weiteres Live-Log und Korrektur (Version 1.3.1)

Nutzerlog zeigt nach OTP: installation redirected → blocked redirected →
endpoint-response HTTP=200 → failed. Daraus ist weder ein konkretes Zielschema
noch der Exceptiontyp ablesbar. Antwortstream wurde bisher ungeprüft an
StreamReader übergeben; unerwartete Fehler beendeten die Browsersitzung ohne
technische Fehlerdiagnose. Fehlende Streams werden nun explizit behandelt,
Laufzeitfehler erhalten die Sitzung, UI-Abschluss wird explizit dispatcht,
Fehlerphase/Exceptiontyp/HRESULT und feste Sperrgründe werden protokolliert.

Zusätzlich wird ausschließlich eine parameterlose HTTP-Rückleitung auf exakt
den konfigurierten Login-Host/-Pfad durch einen einmaligen Aufruf des bekannten
HTTPS-Endpunkts ersetzt; HTTP wird dabei niemals angefragt. Die tatsächliche
Sperrursache im Nutzerlog bleibt bis zum neuen Versuch offen. Regressionstests
prüfen fehlende/valide/ungültige/abgebrochene Antwortstreams und die engen
Grenzen der HTTPS-Rückkehr. Keine erfolgreiche Live-Anmeldung behauptet.

## Anmeldung bestätigt, Server-Transportfehler (Version 1.3.2)

Der Nutzer bestätigt eine funktionierende Anmeldung. Das Log dokumentiert die
blockierte HTTP-Rückleitung, den Aufruf des konfigurierten HTTPS-Endpunkts und
`token-response-validated`. Anschließend scheitert der Serverabruf mit
HttpRequestException 0x80131500. Dieser äußere Fehler allein belegt weder eine
leere Serverliste noch einen bestimmten TLS-/Proxyfehler.

Der Nutzer meldet anschließend einen geänderten API-Vertrag: vor jedem
Authorization-Token ist nun Bearer erforderlich. Serverliste und Logout senden
daher `Authorization: Bearer <token>`; der separate token-Header der Erstellung
behält den dokumentierten Wert. Tests prüfen beide Authorization-Endpunkte und
den getrennten Create-Header. Gespeicherte Tokens werden unverändert weitergenutzt.

Globale Netzwerkvorgaben und Zertifikatsprüfung bleiben erhalten. Zusätzlich
werden innere TLS-/DNS-/Proxy-/Timeoutursachen sekretfrei protokolliert. Neue
ApiRequest-Einträge unterscheiden Transportfehler von HTTP-Status und einer
validen leeren Liste. Die konkrete Ursache des äußeren HttpRequestException ist
im Nutzerlog nicht belegt. Echte Konferenzerstellung und Moderatorzuordnung sind
weiterhin unbestätigt.

## TLS-Verbindungsfehler eingegrenzt (Version 1.3.3)

Das Nutzerlog von Version 1.3.2 meldet beim Serverabruf SecureChannelFailure
vor einer HTTP-Antwort. Der TLS-Verbindungsaufbau scheitert; ein bestimmter
Protokoll-/Cipher-/Zertifikatsgrund ist daraus noch nicht bewiesen. Für den
API-Handler werden nun TLS 1.2/1.3 explizit angeboten. Der Outlook-Prozess
behält seine globalen Einstellungen; Zertifikate werden normal geprüft.

Windows-CI simuliert globale TLS-1.0-Hostvorgaben und prüft moderne Protokolle
am tatsächlichen .NET-Framework-Handler, ausgeschlossene Altprotokolle, normale
Zertifikatsprüfung und unveränderte globale Einstellungen. Dies ist ein
Konfigurations-/Runtime-Regressionstest, kein echter TLS-Handshake gegen die
Nutzerinstallation. Beide Architekturpakete sind als 1.3.3 gekennzeichnet.
Die erfolgreiche Serververbindung nach dieser Änderung bleibt live zu prüfen.

## Neue getrennte Controller-Routen (Version 1.3.4)

Aktuelle Wiki-Rohfassung vom 20.10.2026 geprüft (die normale Wiki-Ansicht lieferte
zunächst einen veralteten Cache). Login bleibt unter room/outlook/plugin;
Server, Erstellung und Logout verwenden api/v1/outlook/plugin. Alle drei
API-Aktionen verlangen Authorization: Bearer, auch die Erstellung. GET/POST sind
nun verbindliche Routenattribute. Antwortschemata und Moderatorgrenze bleiben
wie beschrieben. Quelle:
https://raw.githubusercontent.com/wiki/H2-invent/jitsi-admin/Outlook-Plugin-API-Doku.md

Getrennte LoginBasePath-/ApiBasePath-Konfiguration, exakte Request-URLs mit
Installationsunterpfad, einheitliche Bearer-Header, fehlender alter token-Header,
Login-Endpointbindung sowie Migration alter Standardkonfigurationen werden
getestet. Benutzer- und Verwaltungsdateien bleiben beim Laden bytegleich;
benutzerdefinierte Präfixe bleiben erhalten, beschädigte Vorgaben bleiben
blockierend. TLS-Korrektur aus 1.3.3 wird weiterverwendet. Ein erfolgreicher
Serverabruf/Erstellung mit den neuen Routen ist noch live zu bestätigen.
