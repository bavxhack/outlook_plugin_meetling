# Manuelle Abnahme der Basisversion

Noch nicht unter Outlook abgenommen. Für x86 und x64 getrennt ausfüllen.

| Prüfmerkmal | Tatsächlicher Wert |
| --- | --- |
| Datum / Tester | offen |
| Windows-Version / Build | offen |
| Outlook Classic-Version / Build | offen |
| Outlook-Bitness | offen |
| Paket / Commit / Assemblyversion | offen |
| CI-Lauf | offen |
| Ergebnis / Diagnosebericht | offen |

1. Passendes vollständiges ZIP bei geschlossenem Outlook installieren.
   Installer muss Registrierung melden; falsches Architekturpaket muss vor
   Kopieren/Registrierung abgewiesen werden. .NET Framework 4.8 voraussetzen.
2. Outlook Classic starten; keine Add-in-Fehlermeldung und keine automatischen
   Dialoge, Anmeldungen oder API-Aufrufe.
3. Im Hauptfenster den eigenen Tab **Meetling**, Gruppe **Meetling**, öffnen.
4. **Plugin testen** klicken. Erwartet: „Meetling wurde geladen. Der Button
   funktioniert.“ und tatsächlich gebaute Version (aktuell 1.2.0.0).
5. **Einstellungen** klicken. Erwartet: „Basisversion: Die Server- und
   API-Konfiguration folgt im nächsten Schritt.“, Version und Diagnosepfad;
   keine Zugangsdaten abfragen.
6. Separaten Termin und separates E-Mail-Verfassen-Fenster öffnen.
   Jeweils denselben Tab, drei Controls und beide lokalen Dialoge prüfen.
7. **Meetling-Konferenz** überall sichtbar, deaktiviert; Tooltip
   „Wird nach erfolgreichem Test der Basisversion aktiviert.“
8. Outlook schließen und erneut starten. Tab und Testbutton erneut prüfen.
9. Bei geschlossenem Outlook vorhandene settings.xml und api-key.bin unter
   %LOCALAPPDATA%\Meetling\OutlookAddIn sicher außerhalb des Ordners sichern.
   Test ohne beide Dateien wiederholen. Anschließend ausschließlich mit
   temporären beschädigten Dateien testen. Danach eigene Sicherungen wiederherstellen.
   Niemals produktive Secrets überschreiben oder in einen Bericht kopieren.
10. Prüfen, dass keine Konferenz erstellt und kein Outlook-Element verändert,
    gespeichert oder versendet wurde. Eigene Testentwürfe verwerfen.
11. Diagnosebericht erstellen. OnConnection (ConnectMode), GetCustomUI mit
    jeweiliger Ribbon-ID, beide Callback-Eintritte und Shutdown prüfen.
    OnStartupComplete ist beim Startladen zu erwarten, nicht zwingend beim
    nachträglichen manuellen Aktivieren. Ein expliziter COM-Test ist kein Outlooklauf.
12. Bei geschlossenem Outlook dasselbe Paket nochmals installieren. Neustart
    und Testbutton prüfen. Deinstallation prüfen: nur eigene Registrierung/
    Laufzeit-DLLs entfernt, Einstellungen und Secrets unverändert erhalten.

Bei Fehlschlag Kontext, Zeitpunkt und tatsächliche Fehlermeldung dokumentieren.
XML-Fehleranzeige und gezielte Reaktivierung siehe README. Keine fremden Add-ins
oder globalen Resiliency-Daten löschen. Erfolg erst nach echtem Lauf eintragen.

## Kalenderbutton und Mock-Terminplaner (1.2.0.0)

Die Installation und die bisherigen Testbuttons wurden vom Nutzer als
funktionierend bestätigt. Exakte Windows-/Outlook-Version und Bitness bitte
weiterhin oben festhalten. Die nachfolgenden neuen UI-Schritte stehen noch aus:

1. Aktualisiertes passendes Paket installieren und Outlook Classic starten.
2. Zur Kalenderübersicht wechseln und den eingebauten Tab **Start** öffnen.
3. Direkt nach der Gruppe **Neu** die Gruppe **Meetling** mit **Neues Meetling**
   prüfen. Andere Add-ins können die endgültige Gruppenreihenfolge beeinflussen.
4. Das Symbol muss ausschließlich den grünen Würfel mit zwei weißen Punkten
   zeigen, ohne Schriftzug und ohne die beiden kleineren Würfel. Ecken transparent;
   klassisches und vereinfachtes Ribbon sowie 100/150/200 % Skalierung prüfen.
5. Einmal klicken. Erwartet: neue native Outlook-Besprechung mit Teilnehmerfeldern,
   Datum/Uhrzeit und verfügbarer Terminplanungs-Assistent-Ansicht.
6. Body enthält „Testlink (Mock – keine echte Konferenz)“ und eine eindeutige
   https://meetling.example.invalid/mock/…-Adresse. Sie startet keine echte Konferenz.
7. Betreff, Teilnehmer und Zeiten sind normal editierbar. Kein automatisches
   Speichern oder Versenden; Testentwurf ohne Speichern schließen.
8. Vorhandene ausgewählte Termine bleiben unverändert. Keine API-/Anmeldeabfrage.
9. Nach Neustart, bei fehlenden/beschädigten Einstellungsdateien und unter x86/x64
   wiederholen. Der Mock-Button darf diese Dateien nicht lesen oder ändern.
10. Diagnose prüfen: CreateMeeting entered (mock); keine Konferenzlinks im Log.

CI ersetzt ausschließlich die native Editor-Grenze durch eine Testfunktion,
öffnet also keinen Outlook-Dialog. Die erfolgreiche CI beweist nicht diese
Outlook-Anzeige oder das Öffnen des echten Besprechungsfensters.
