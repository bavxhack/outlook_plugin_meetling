# Meetling für Outlook Classic – lokale Basisversion

Klassisches COM-Add-in für .NET Framework 4.8, kein VSTO. ProgID
`Meetling.OutlookAddIn`, CLSID `{4BB84EF2-9D83-4710-BDBA-7CBFD6939B48}`.
Version 1.1.0.0. Unterstützt Outlook Classic unter Windows; das neue Outlook
unterstützt dieses COM-Add-in nicht.

Der eigene Tab **Meetling** erscheint im Hauptfenster, im separat geöffneten
Termin und im separaten E-Mail-Verfassen-Fenster. **Plugin testen** zeigt eine
lokale Bestätigung mit Assemblyversion. **Einstellungen** zeigt den Ausbauhinweis,
Version und Diagnosepfad. **Meetling-Konferenz** ist deaktiviert und hat keinen
Callback. Start, Ribbon und Buttons lesen keine Einstellungen oder Secrets,
rufen keine API auf und verändern keine Outlook-Elemente.

## Standardweg: ZIP und Windows PowerShell 5.1

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
er XML-Struktur, drei Fensterkontexte, deaktivierte Konferenz und Lifecycle.
Core-Tests sind reine Logiktests, keine Outlook-Integrationstests.

**WiX/MSI ist ein nicht abgenommener alternativer Weg.** Der vorhandene Entwurf
enthält weiterhin feste Versions-/Assemblyangaben, abweichenden Zielordner,
keine Outlook-Interop-DLL und keine gleichwertigen Vorabprüfungen. Er wird in
diesem Meilenstein weder ausgeliefert noch als Standardweg empfohlen.

## Prüfstatus und weitere Arbeit

Ausgangsstand erneut geprüft: `main` / `c623582f61fcb165f5cb40338642bd5de3eab27d`.
Bestätigte Codebefunde und verbleibende Grenzen:
[docs/verification.md](docs/verification.md).
Einen aktuellen erfolgreichen CI-Lauf und die konkrete Windows-/Outlook-Abnahme
getrennt nachweisen. Eine erfolgreiche Kompilierung oder COM-Aktivierung belegt
keine Ribbon-Anzeige in Outlook.

Meetling-API, Serverauswahl und Bearbeitung von Terminen/Nachrichten bleiben
ausgeschaltet. Der fachliche Ausbau beginnt erst nach Basisabnahme:
[docs/roadmap.md](docs/roadmap.md).
