# Roadmap

- Erhalten: klassisches COM-Add-in, .NET Framework 4.8, ProgID/CLSID,
  Default-Callbackinterface/DISPIDs, Logo, Kalenderbutton, x86/x64-Paketierung.
- Implementiert: Outlook-Plugin-API, WebView2-Browserlogin, DPAPI-Tokenbindung,
  verwaltete JSON-Konfiguration und GUI-Fallback, asynchrone Serverauswahl,
  Konferenzerstellung, native Linkeinfügung, Logout und lokale Abmeldung.
- Automatisierte Abnahme: Windows Actions mit HTTP-Mocks, Konfigurationstests,
  DPAPI- und COM-/Ribbon-Smoke unter Standardbenutzer; fertige Architektur-ZIPs.
- Offen: echte SSO-Anmeldung/Rückkehr, eingebettete Browserfreigabe und MFA,
  Moderatoridentität nach Neustart ohne Browsersitzung, native Outlook-Abnahme
  mit beiden verfügbaren Bitnessen. Siehe smoke-test.md und plugin-api.md.
- Nicht im Umfang: Backendänderung, Terminsynchronisierung, Refresh-/Device-Code-
  Endpunkte, externe Browsercallbacks oder Cookie-Workarounds.
