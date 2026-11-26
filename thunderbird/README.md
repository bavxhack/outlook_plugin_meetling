# Meetling für Thunderbird

Thunderbird-Erweiterung für dieselbe Meetling-API wie das Outlook-Plugin. Zielversion der ersten Ausgabe: **Thunderbird 158**, Windows/Linux/macOS. Ein XPI für alle Architekturen. Andere Hauptversionen sind bis zur Prüfung bewusst nicht freigeschaltet.

## Installation und Verwendung

1. `Meetling-Thunderbird-1.0.1.xpi` im Actions-Artefakt herunterladen.
2. Thunderbird → Add-ons und Themes → Zahnrad → Add-on aus Datei installieren.
3. Meetling-Schaltfläche in der Hauptsymbolleiste öffnen, Einstellungen aufklappen und die HTTPS-Basisadresse eintragen. Bestehende Outlook-JSON-Konfiguration lässt sich unter „Erweiterte API-Konfiguration“ importieren.
4. „Anmelden“ öffnet den Thunderbird-Browser für SSO/OTP. Nach der validierten Token-Antwort schließt sich das Loginfenster. Benutzername und Serverbuttons erscheinen im Meetling-Fenster.
5. Im geöffneten Kalendertermin „Meetling hinzufügen“ wählen und den gewünschten Server anklicken. Alternativ über die Hauptsymbolleiste starten: Beim Serverklick öffnet sich zunächst ein neuer, ungespeicherter Kalendertermin.
6. Teilnahmelink und „Konferenz in Meetling bearbeiten (nur für Moderatoren)“ werden an die Beschreibung angehängt. Die ID stammt aus der Erstellungsantwort, der Link verwendet `baseUrl + room/dashboard?editConference_id=<id>`. Vorhandene Formatierung und Terminangaben bleiben bestehen. Anschließend selbst prüfen, speichern und Einladungen versenden.

Das Fenster nutzt skalierbare HTML-Buttons, native Systemfarben und Hell-/Dunkeldarstellung. Es besitzt keinen kleinen, verschachtelten Scrollbereich für die Serverliste.

## Konfiguration

```json
{
  "baseUrl": "https://jitsiadmin.dev07.h2-invent.com/",
  "apiBasePath": "api/v1/outlook/plugin/",
  "loginBasePath": "room/outlook/plugin/",
  "endpoints": {
    "login": "outlooklogin", "servers": "server",
    "createConference": "createConference", "logout": "logout"
  },
  "timeoutSeconds": 30
}
```

Login: GET `/room/outlook/plugin/outlooklogin`. Serverliste: GET `/api/v1/outlook/plugin/server`. Erstellung: POST `/api/v1/outlook/plugin/createConference`, URL-kodiertes Formular `server_id`. Logout: GET `/api/v1/outlook/plugin/logout`. Alle API-Aktionen verwenden `Authorization: Bearer <token>`, keine Cookies und keine HTTP-Weiterleitungen. Installationsunterpfade bleiben erhalten. Das bekannte alte API-Präfix wird beim Einlesen im Speicher migriert.

## Sicherheit und Berechtigungen

- Token werden installationsgebunden in Thunderbirds verschlüsselter Passwortverwaltung gespeichert. Ein gesetztes Hauptpasswort schützt auch diese Einträge. `storage.local` enthält nur die Konfiguration. Kein Token wird an die Oberfläche, Protokolle oder Synchronisation weitergegeben.
- HTTPS-Hostzugriff ist für frei konfigurierbare Meetling-Installationen nötig. Der HTTP-Zugriff dient ausschließlich dazu, die bekannte fehlerhafte HTTP-Rückleitung des exakten Login-Endpunkts durch HTTPS zu ersetzen. Ein zweiter Downgrade wird blockiert. Keine TLS- oder Zertifikatsausnahmen.
- `webRequest`/`webRequestBlocking` lesen nur die Antwort des konfigurierten Login-Endpunkts im gerade geöffneten Login-Tab. Andere Webseiten liefern keine Tokens. Die Listener werden bei Abschluss/Abbruch entfernt.
- Die kleine mitgelieferte Experiment-API benötigt Thunderbird-Vollzugriff, weil die Standard-MailExtension-APIs den Kalendereditor nicht abdecken. Sie ergänzt den Kalenderbutton, erfasst genau den ursprünglichen Entwurf, fügt HTML am Beschreibungsende ein und nutzt die Passwortverwaltung. Sie speichert oder versendet keine Termine.
- Keine automatischen Wiederholungen der Erstellung. Bei unklarem Netzwerkergebnis zuerst Meetling prüfen. Bei geschlossenem Termin werden die erzeugten Links zur manuellen Übernahme angezeigt; es wird kein Ersatzraum erstellt.
- Fehlende/ungültige Konferenz-ID: Teilnahmelink bleibt erhalten, Hinweis auf fehlenden Bearbeitungslink. Den tatsächlichen Moderatorzugriff prüft Meetling serverseitig.

## Entwicklung und Prüfung

`node --test thunderbird/tests/*.test.cjs`

`python thunderbird/tools/package.py`

Die Actions führen die API-/Zustandstests aus, bauen ein reproduzierbares XPI mit SHA-256-Prüfsumme und prüfen den Kalenderadapter in einem echten Thunderbird 158 unter Xvfb. Der Test benutzt ein frisches Profil und synthetische Daten, keinen produktiven Meetling-Account. SSO/OTP und der vollständige Ablauf gegen eine echte Installation müssen anschließend manuell geprüft werden.

Technische Quellen: [Thunderbird Experiments](https://developer.thunderbird.net/add-ons/mailextensions/experiments), [Thunderbird Kalender-Quellcode](https://github.com/mozilla/releases-comm-central/tree/master/calendar/base/content/item-editing), [WebRequest](https://webextension-api.thunderbird.net/en/esr-mv2/webRequest.html), [Meetling API](https://github.com/H2-invent/jitsi-admin/wiki/Outlook-Plugin-API-Doku).
