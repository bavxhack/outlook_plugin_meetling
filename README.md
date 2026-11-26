# Meetling für Outlook Classic – Outlook-Plugin-API

Klassisches COM-Add-in für .NET Framework 4.8, Version 1.3.7.0.
ProgID und CLSID sowie Logo und Kalenderbutton bleiben erhalten.
Die aktive Konferenzaktion verwendet die dokumentierte Outlook-Plugin-API.

Über **Einstellungen** eine HTTPS-Installationsadresse eintragen oder eine
administrierte JSON-Konfiguration verteilen. Beim nächsten Outlook-Start öffnet
sich bei fehlender gespeicherter Anmeldung einmalig das Anmeldefenster. Eine
vorhandene Anmeldung wird wiederverwendet. **Anmelden** startet den Ablauf auch
manuell; nach Abbruch erfolgt kein automatischer Wiederholungsversuch in derselben
Outlook-Sitzung. Ohne gültige Konfiguration wird kein Fenster geöffnet. Das sichtbare
WebView2-Fenster verwendet den SSO-Anmeldeweg der Installation. Eine vorhandene
SSO-Sitzung kann die Anmeldung ohne Eingabe abschließen; erforderliche Zugangsdaten,
Zustimmungen oder MFA bestätigt der Nutzer weiterhin. Anschließend
**Server aktualisieren** oder **Neues Meetling** wählen. Bei mehreren Servern
muss die Auswahl bestätigt werden. Outlook-Entwürfe werden ergänzt, weder
automatisch gespeichert noch versendet.

Voraussetzung für Login: Microsoft Edge WebView2 Evergreen Runtime. COM-Laden
und Diagnose funktionieren auch ohne Runtime. Sichere Tokenspeicherung mit DPAPI,
GET-Logout ohne Browser, ProgramData-Konfiguration und bekannte Integrationsgrenzen
sind in [API und Administration](docs/plugin-api.md) beschrieben.

Der Nutzer bestätigt den bisherigen Ablauf als funktionierend. Version 1.3.6
ersetzt die Serverauswahl durch große anklickbare Schaltflächen und ergänzt
unter dem Teilnahmelink einen Bearbeitungslink mit „nur für Moderatoren“ aus der
zurückgegebenen Konferenz-ID. Beide Links sind native Outlook-/Word-Hyperlinks.
Neue Darstellung und Links sind nach Installation manuell zu prüfen; siehe
[Prüfbericht](docs/verification.md) und [manuelle Abnahme](docs/smoke-test.md).

Version 1.3.7 übernimmt auch zurückgegebene HTTP-Konferenzlinks. Damit werden
bei solchen Antworten die Einwahlnummern und der Einwahl-PIN ebenfalls eingefügt.
API-Verbindung und Anmeldung verlangen weiterhin HTTPS.

## Standardweg: MSI per Doppelklick

1. Unter **Datei → Office-Konto → Info zu Outlook** die Outlook-Bitness prüfen.
2. Im [aktuellen Release](https://github.com/bavxhack/outlook_plugin_meetling/releases/tag/continuous)
   `Meetling-OutlookAddIn-x86.msi` für 32-Bit-Outlook oder
   `Meetling-OutlookAddIn-x64.msi` für 64-Bit-Outlook herunterladen.
3. Outlook vollständig schließen und die MSI per Doppelklick öffnen.
   Die Installation erfolgt für den aktuellen Benutzer ohne Administratorrechte
   und ohne PowerShell-Befehle. Danach Outlook starten.

Die MSI-Pakete werden nach erfolgreicher Windows-CI als Actions-Artefakte und
über den bestehenden Release-Schritt veröffentlicht. Im Actions-Lauf stehen unter
**Artifacts** die separaten Downloads `Meetling-OutlookAddIn-MSI-x86` und
`Meetling-OutlookAddIn-MSI-x64` bereit. GitHub verpackt Artefakte beim Download als
ZIP; daraus die MSI entpacken und per Doppelklick starten. Windows kann bei den
unsignierten Paketen eine Sicherheitsbestätigung verlangen; ein signierter,
warnungsfreier Download wird nicht zugesichert. .NET Framework 4.8 ist erforderlich
und wird vor der Installation geprüft. Für die Anmeldung muss zusätzlich die
Microsoft Edge WebView2 Evergreen Runtime vorhanden sein.
Nur Outlook Classic unterstützt dieses COM-Add-in.

Updates: Outlook schließen und die neue MSI derselben Architektur öffnen.
Auch erneute Builds derselben Version ersetzen die vorhandene MSI-Installation.
Eine niedrigere Produktversion wird abgewiesen. Unter **Windows-Einstellungen →
Apps → Installierte Apps → Meetling Outlook Add-in** lässt sich das Add-in
entfernen; Einstellungen, Anmeldung und Logs bleiben erhalten.
Eine vorhandene ZIP-Installation zuerst mit deren `Uninstall.ps1` entfernen.
Beim Wechsel der Outlook-Bitness die bisherige Installation zuerst deinstallieren.
Die MSI installiert nach `%LOCALAPPDATA%\Meetling\OutlookAddIn\<x86|x64>`.

Die CI prüft beide MSIs unter einem temporären Standardbenutzer: Installation,
Reparatur einer gelöschten DLL, COM-/Ribbon-Aktivierung, Deinstallation und Erhalt
der Benutzerdaten. Auf dem kurzlebigen Windows-Server-CI-Runner erlaubt der
Testtreiber dafür vorübergehend benutzerbezogene MSI-Installationen (`DisableMSI=0`)
und stellt die vorherige Richtlinie im `finally`-Block wieder her. Der Testbenutzer
bleibt ohne Administratorrechte; Unternehmensrichtlinien auf Zielgeräten gelten
weiterhin. Die tatsächliche Outlook-Oberfläche bleibt Teil der
[manuellen Abnahme](docs/smoke-test.md).

## Alternativ: ZIP und Windows PowerShell 5.1

1. Outlook-Bitness unter **Datei → Office-Konto → Info zu Outlook** feststellen.
   Die Windows-Bitness allein reicht nicht.
2. Aus den [Actions-Artefakten](https://github.com/bavxhack/outlook_plugin_meetling/actions/workflows/build-outlook-addin.yml)
   das passende `Meetling-OutlookAddIn-x86.zip` oder `-x64.zip` vollständig
   entpacken. Die Pakete entstehen erst nach erfolgreicher Windows-CI.
3. Outlook vollständig schließen. Im entpackten Verzeichnis Windows PowerShell
   5.1 öffnen und ausführen (für 32-Bit-Outlook `x86` einsetzen):

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -OutlookBitness x64
   ```

   Als normaler Benutzer installieren und Outlook ohne Administratorrechte starten. Lokale Unternehmensrichtlinien gelten
   weiterhin. .NET Framework 4.8 muss installiert sein.
4. Registrierung abgeschlossen bedeutet noch keine bestätigte Outlook-Ladung.
   Jetzt [die manuelle Abnahme](docs/smoke-test.md) vollständig durchführen.

Der Installer prüft Paketidentität, explizite PE-/Assemblyarchitektur, Framework,
laufendes Outlook und – soweit verfügbar – die Bitness der registrierten
OUTLOOK.EXE bzw. Outlook-spezifische Registrywerte. Ist Outlook nicht zuverlässig
erkennbar, warnt er; die manuelle Bitnessprüfung bleibt erforderlich.
Laufzeit-DLLs werden vollständig nach
`%LOCALAPPDATA%\Meetling\OutlookAddIn\<x86|x64>\<Version>` kopiert und per Hash
verglichen. CodeBase und Assemblyidentität werden aus diesem Paket abgeleitet;
die Registrierung erfolgt pro Benutzer in der passenden Registryansicht.
Dieselbe Version kann bei geschlossenem Outlook erneut installiert werden.
Ältere Versionsordner bleiben bis zur Deinstallation erhalten. Keine parallelen
MSI- und ZIP-Installationen oder Installation beider Bitnessen vornehmen.

Deinstallation bei geschlossenem Outlook:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Uninstall.ps1 -OutlookBitness x64
```

Entfernt ausschließlich eigene Registrierung und bekannte Laufzeit-DLLs.
Einstellungen, Secrets und Logs bleiben erhalten. Für Updates denselben
Installationsweg verwenden; Einstellungen werden nicht migriert oder gelesen.

## Diagnose

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Diagnose.ps1 -OutlookBitness x64
```

Bericht standardmäßig auf dem Desktop. Das Werkzeug liest eigene Registrierung
in HKCU und HKLM, prüft das registrierte Paket und zeigt letzte Lifecycle-Einträge.
Es liest keine Secrets, Mailtexte, Konferenzlinks oder vollständigen
Windows-Ereignismeldungen. Das Log liegt unter
`%LOCALAPPDATA%\Meetling\OutlookAddIn\diagnose.log`, mit maximal rund 256 KiB
pro Datei und einer rotierenden Sicherung. Bei Schreibfehlern fällt Logging
still aus; ein fehlendes Log allein beweist keinen Ladefehler.

Ein COM-Aktivierungstest ist optional und startet nur ausdrücklich mit
`-TestComActivation`, bei geschlossenem Outlook und unter Windows PowerShell
5.1 mit passender Prozessarchitektur:

- x64: `%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe`
- x86 auf 64-Bit-Windows: `%WINDIR%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`

Dieser Test erzeugt nur die COM-Instanz; er öffnet keine Dialoge und ersetzt
keinen Outlook-Test. Falsche Prozessbitness und PowerShell 7 werden hierfür
abgewiesen. Im Repository heißen die Skripte `Install-PublishedAddIn.ps1`,
`Uninstall-AddIn.ps1` und `Collect-Diagnostics.ps1`.

Bei Deaktivierung: **Datei → Optionen → Add-Ins → Verwalten: Deaktivierte
Elemente → Los** und gezielt Meetling reaktivieren. Anschließend unter
**COM-Add-Ins** ausschließlich Meetling aktivieren. Gegebenenfalls unter
**Langsame und deaktivierte COM-Add-Ins** prüfen. Keine Resiliency-/DisabledItems-
Schlüssel pauschal löschen. Zur XML-Diagnose unter den erweiterten Outlook-
Optionen **Fehler von Add-In-Benutzeroberflächen anzeigen** aktivieren
(Bezeichnung kann je nach Outlook-Version variieren).

## Windows-Build und automatisierte Prüfung

Visual Studio Build Tools 2022 mit .NET Framework 4.8 Targeting Pack und .NET SDK
8 erforderlich. Im Developer PowerShell:

```powershell
msbuild src\Meetling.OutlookAddIn\Meetling.OutlookAddIn.csproj /restore /t:Rebuild /p:Configuration=Release /p:Platform=x86
msbuild src\Meetling.OutlookAddIn\Meetling.OutlookAddIn.csproj /restore /t:Rebuild /p:Configuration=Release /p:Platform=x64
msbuild tests\Meetling.Core.Tests\Meetling.Core.Tests.csproj /restore /t:Rebuild /p:Configuration=Release
& tests\Meetling.Core.Tests\bin\Release\net48\Meetling.Core.Tests.exe
```

Ausgabe explizit `src\Meetling.OutlookAddIn\bin\<Architektur>\Release\net48`.
Für lokale Installation daraus:
`installer\Install-AddIn.ps1 -OutlookBitness x64`.

Der bestehende Windows-Workflow baut beide Architekturen, prüft und verpackt
alle DLLs, installiert das Paket zweimal unter einem temporären Standardbenutzerkonto und führt einen passenden nativen
Prozess mit dem .NET-Framework-Smoke-Harness aus. Dieser aktiviert die
registrierte COM-Klasse, prüft die Hostinterfaces und ruft die Callbacknamen
über echtes Standard-IDispatch (`GetIDsOfNames` und `Invoke`) auf. Nur die
interne Dialogausgabe wird im Test ersetzt; keine CI-Dialoge. Zusätzlich prüft
er XML-Struktur, drei Fensterkontexte, zusätzliche Callbacks, DPAPI und Lifecycle.
Zusätzlich prüft er den Bildcallback als echtes IPictureDisp, die Originalfarben
und Transparenz des Würfels. Die native Outlook-UI wird im CI nicht geöffnet.
HTTP-Tests sind simuliert, keine Live-Integrationstests.

Die Pipeline baut zusätzlich mit WiX 3.14.1 eigenständige MSI-Pakete inklusive
aller sieben Laufzeit-DLLs und COM-Registrierung in der passenden Registryansicht.
Produkt- und Assemblyversion werden aus der gebauten DLL übernommen. ZIPs bleiben
als alternativer Installationsweg verfügbar. Die MSI-SHA256-Prüfsummen werden mit
veröffentlicht. Lokale MSI-Builds benötigen WiX 3 und die MSBuild-Properties
`ProductVersion` (dreiteilig) und `AssemblyVersion` (vierteilig) passend zur DLL.

Der verifizierte Ausgangsstand ist `main` / `0ca8de3d9741af5378dde2989239e78e551b5066`.
Siehe [Prüfbericht](docs/verification.md) und [Roadmap](docs/roadmap.md).
