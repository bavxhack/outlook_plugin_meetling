# Nach der Abnahme: Meetling-Anbindung

Nur dokumentiert, nicht Teil der aktiven Basisversion. Umsetzung erst in einem
weiteren Auftrag nach erfolgreicher Outlook-Abnahme. Bestehende SettingsStore,
MeetlingApiClient, AppointmentService und SettingsForm bleiben als Ausgangsmaterial
erhalten; ihr heutiges Verhalten definiert keine neuen Anforderungen.

## Serverauswahl und Konfiguration

- Dropdown der für den authentifizierten Nutzer verfügbaren Server.
- Anzeigename und technische Server-ID getrennt behandeln.
- Frei konfigurierbare Base-URL, Serverlisten-URL und Erstellungs-URL.
- JSON-Konfiguration; Auflösung relativer/absoluter URLs dokumentieren.
- Bestehende XML-Einstellungen kontrolliert migrieren; Secrets erhalten,
  niemals unbemerkt verlieren oder in Klartext übernehmen.
- Asynchroner Servercache und Ribbon-Aktualisierung ohne blockierendes Netzwerk.

## Sofortkonferenz

- Auswahl erstellt über die tatsächlich belegte Meetling-Schnittstelle eine
  Sofortkonferenz; ausgewählte Server-ID übermitteln.
- Meetling ordnet den authentifizierten Nutzer als Moderator zu.
- Heutige Endpunkte /api/v1/room und /api/v1/info/{uid} nicht ungeprüft übernehmen.
- API-Vertrag, Authentifizierung und Link-Gültigkeit zunächst verifizieren.
- Keine automatischen Erstellungs-Retries nach unklarem Ergebnis.

## Outlook-Verhalten

- Hauptfenster: neuen nativen Termin öffnen.
- Geöffneter Termin: genau diesen ergänzen.
- E-Mail-Verfassen-Fenster: nur diesen Nachrichtenentwurf ergänzen.
- Zielkontext vor dem Request erfassen und anschließend beibehalten.
- Teilnehmerlink am Ende des vorhandenen Bodys einfügen; eigener Text bleibt
  darüber möglich. Formatierung und native Terminfelder erhalten.
- Keine Host-/Moderatorlinks in den Teilnehmerbody.
- Kein automatisches Save oder Send.
- Keine Synchronisation bei Änderungen, Absagen oder Löschungen.
- Duplikate verhindern. Bei Einfügungsfehlern bereits erstellten Link manuell
  kopierbar anbieten.
- Alle Office-COM-Zugriffe auf dem Outlook-UI-Thread.
