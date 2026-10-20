# Manuelle Abnahme – noch nicht ausgeführt

Testprotokoll je verfügbarer Outlook-Classic-Bitness führen: Version, Windows,
Outlook, WebView2 Runtime, Serverversion, Testkonto, Zeit, Ergebnis und Beleg.
Keine Secrets protokollieren.

1. Passendes x86-/x64-ZIP aus dem PR-Actions-Lauf unter Standardbenutzer
   installieren, wiederholt installieren, Outlook starten. Logo, Kalenderbutton,
   bestehende Diagnosecallbacks und Ribbon in Explorer/Termin/Mail prüfen.
2. Ohne Konfiguration, offline, mit beschädigter Verwaltungsdatei und ohne
   WebView2 Runtime starten: Outlook/Ribbon/Einstellungen bleiben erreichbar;
   nur Browseranmeldung meldet eine fehlende Runtime.
3. Ohne Verwaltungsdatei URL samt Unterpfad in GUI speichern. Danach gültige
   ProgramData-Datei mit Admin-Schreib-/Benutzer-Leserechten verteilen.
   Organisationshinweis/Sperre prüfen, ungültige Datei darf GUI nicht aktivieren.
4. Anmelden: tatsächliches SSO/MFA, Rückkehr zum exakten Login-Endpunkt,
   automatische Tokenübernahme, keine Token in JSON-Konfiguration/Logs/Body.
   Fehlt die Rückkehr, Integrationsgrenze protokollieren. Fremde Origin/Pfade,
   Abbruch, Timeout und mehrfachen Klick prüfen.
5. Outlook neu starten: keine automatische Browseröffnung, Token wird sicher
   wiederverwendet; Server aktualisieren. Serverlisten mit 0/1/2/>2 Einträgen,
   gleichen Namen, Sortierung, ID-Auswahl und Abbruch prüfen.
6. Neues Meetling aus Kalender: nativer Besprechungsentwurf. Termin-Inspector
   sowie Mail-Compose: vorhandenes Element ergänzen. Während Netzwerkaktivität
   aktives Fenster wechseln, schließen, doppelklicken; nur erfasstes Ziel darf
   ergänzt werden, kein zweiter Raum durch Doppelklick. Keine automatische
   Speicherung/Sendung. Eigenen Text, HTML/RTF, Signaturen, Empfänger, Zeiten,
   Anhänge und Betreff bestehender Elemente erhalten. Selbst speichern/senden.
7. Moderator serverseitig mit Tokenbenutzer vergleichen, insbesondere nach
   Outlook-Neustart ohne Browser-Sitzung. Erfolgs-JSON reicht nicht aus.
8. Netzwerk während Create abbrechen: unklaren Ausgang anzeigen, kein Retry.
   Ziel nach erfolgreicher Anlage schließen: manuellen Link anbieten.
9. Abmelden ohne Browser: Warnung für andere Clients, bestätigter Widerruf,
   lokales Löschen und Ablehnung alten Tokens prüfen. Bereits ungültigen Token
   und Netzfehler/erneuten Versuch/„nur lokal“ prüfen. API-Logout beendet keine
   Symfony-Sitzung. Ein neuer Login verwendet separates neues Profil.
10. Während Login/Serverrequest Konfiguration wechseln, Logout auslösen oder
    Outlook schließen: verspätete Antworten dürfen Zustand nicht reaktivieren.
    Erneute Anmeldung für das neue Ziel, kein Token an frühere/fremde Ziele.
11. Für jede tatsächlich verfügbare Outlook-Architektur wiederholen und
    deinstallieren. Nicht vorhandene Architekturen als ungetestet markieren.
