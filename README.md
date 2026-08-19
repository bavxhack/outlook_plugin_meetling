# Meetling Outlook Add-in

Ein COM-Add-in für **klassisches Microsoft Outlook 2021 und 2024 unter Windows**. In einem geöffneten Termin ergänzt es die Ribbon-Gruppe **Meetling** mit **Konferenz erstellen** und **Einstellungen**. Microsoft 365 und das neue Outlook sind nicht Zielplattformen.

## Voraussetzungen

- Windows 10/11 mit .NET Framework 4.8
- klassisches Outlook 2021 oder Outlook 2024
- Visual Studio 2022 mit „.NET-Desktopentwicklung“ und installierten Office Primary Interop Assemblies
- optional WiX Toolset 3.14 für das MSI
- eine Meetling/Jitsi-Admin-Instanz, API-Schlüssel und Name/ID eines dort hinterlegten Konferenzservers

## Kompilierung

Outlook und Add-in müssen dieselbe Bitness besitzen. In einer „Developer PowerShell for VS 2022“:

```powershell
msbuild .\Meetling.OutlookAddIn.sln /restore /t:Build /p:Configuration=Release /p:Platform=x64
# oder
msbuild .\Meetling.OutlookAddIn.sln /restore /t:Build /p:Configuration=Release /p:Platform=x86
```

Die Bibliotheken liegen danach unter `src\Meetling.OutlookAddIn\bin\Release\<Plattform>\net48`.

## Outlook-Bitness bestimmen

In Outlook **Datei > Office-Konto > Info zu Outlook** öffnen. Am Anfang der Versionszeile steht „32-Bit“ oder „64-Bit“. Alternativ kann `HKLM\Software\Microsoft\Office\ClickToRun\Configuration\Platform` geprüft werden. Nicht die Windows-, sondern die Outlook-Bitness ist maßgeblich.

## Installation und Registrierung

Outlook schließen und PowerShell mit normalen Benutzerrechten öffnen. Die Installation ist benutzerbezogen und benötigt keine Administratorrechte:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\installer\Install-AddIn.ps1 -OutlookBitness x64
# Für 32-Bit-Outlook: -OutlookBitness x86
```

Das Skript kopiert die Release-Dateien nach `%LOCALAPPDATA%\Meetling\OutlookAddIn`, verwendet das zur Outlook-Bitness passende `RegAsm.exe` und registriert das Add-in unter `HKCU`. Danach Outlook neu starten.

### MSI (optional und signierbar)

Nach dem passenden Release-Build mit installiertem WiX Toolset:

```powershell
msbuild .\installer\Meetling.OutlookAddIn.wixproj /p:Configuration=Release /p:Platform=x64
```

Für x86 entsprechend `Platform=x86`. Das resultierende MSI kann anschließend mit dem unternehmenseigenen Code-Signing-Zertifikat und `signtool sign` signiert werden. Für produktive Verteilung wird ein Strong Name beziehungsweise eine organisationsspezifisch signierte Assembly empfohlen.

## Deinstallation

Outlook schließen, dann:

```powershell
.\installer\Uninstall-AddIn.ps1 -OutlookBitness x64
```

Die persönlichen Einstellungen unter `%LOCALAPPDATA%\Meetling\OutlookAddIn` bleiben zur Vermeidung eines unbeabsichtigten Schlüsselverlusts erhalten und können manuell gelöscht werden.

## Konfiguration

1. Einen Termin in einem eigenen Fenster öffnen.
2. **Meetling > Einstellungen** wählen.
3. Basis-URL (ohne API-Pfad), API-Schlüssel und den in Meetling hinterlegten Konferenzserver eintragen.
4. Optional Organisator-E-Mail und Keycloak-ID eintragen.

Die Basis-URL wird validiert und abschließende `/` werden entfernt. Unverschlüsseltes HTTP wird nur bei localhost ohne Warnung akzeptiert. Der API-Schlüssel wird getrennt von den übrigen benutzerbezogenen Einstellungen mit Windows DPAPI (`CurrentUser`) verschlüsselt gespeichert und nie protokolliert.

## Verwendung

Terminbetreff sowie Start und Ende festlegen und **Konferenz erstellen** klicken. Das Add-in ermittelt bevorzugt die SMTP-Adresse des sendenden Kontos und verwendet andernfalls die konfigurierte Organisator-E-Mail. Nach erfolgreicher Anlage wird der Browser-Teilnehmerlink am Anfang des Termintexts ergänzt; vorhandenes HTML und Signaturen bleiben erhalten. UID und Link werden als Outlook-Benutzereigenschaften `MeetlingUid` und `MeetlingParticipantLink` gespeichert. Dadurch wird ein versehentlicher zweiter Aufruf blockiert. Der Termin wird abschließend gespeichert.

## API-Anbindung

Die Implementierung verwendet ausschließlich:

- `POST {Basis-URL}/api/v1/room` mit `application/x-www-form-urlencoded`: `email`, `name`, `duration`, `server`, `start` und optional `keycloakId`.
- `GET {Basis-URL}/api/v1/info/{uid}`.
- `Authorization: Bearer <API-Schlüssel>` bei beiden Aufrufen.
- `uid` aus der POST-Antwort und `room_url` als Browser-Teilnehmerlink aus der Info-Antwort.

Referenz sind die [API-Endpunkte im Jitsi-Admin-Wiki](https://github.com/H2-invent/jitsi-admin/wiki/API-Endpoints) und das [Jitsi-Admin-Quellrepository](https://github.com/H2-invent/jitsi-admin). Vor einem Rollout gegen eine abweichende Serverversion sollte deren API-Vertrag nochmals geprüft werden. Der Client hat ein 30-Sekunden-Timeout, prüft HTTP-Status, JSON, `error`, UID und URL und deaktiviert die TLS-Zertifikatsprüfung nicht.

## Tests

Die Tests sind ein paketfreier .NET-Framework-Konsolen-Testläufer, damit keine Testadapter-Abhängigkeit erforderlich ist:

```powershell
msbuild .\tests\Meetling.Core.Tests\Meetling.Core.Tests.csproj /restore /p:Configuration=Release
.\tests\Meetling.Core.Tests\bin\Release\net48\Meetling.Core.Tests.exe
```

Sie prüfen URL-Regeln, Dauer, Formulardaten, Bearer-Header, erfolgreiche und fehlerhafte Antworten, ungültiges JSON, fehlende UID beziehungsweise URL und Duplikaterkennung. HTTP wird vollständig gemockt. Ribbon-Laden, COM-Lebenszyklus und das tatsächliche Speichern eines Outlook-Termins müssen auf einem Windows-Testrechner mit der jeweiligen 32-/64-Bit-Outlook-Version als Integrationstest geprüft werden.

## Fehlerdiagnose

Kurze Meldungen erscheinen auf Deutsch. Technische Diagnosen stehen in `%LOCALAPPDATA%\Meetling\OutlookAddIn\diagnose.log`. Das Protokoll enthält keine API-Schlüssel und keine vollständigen Termininhalte. Falls Outlook das Add-in deaktiviert: **Datei > Optionen > Add-Ins > Verwalten: Deaktivierte Elemente/COM-Add-Ins** prüfen sowie Bitness und `LoadBehavior=3` kontrollieren.

## Bekannte Einschränkungen

- Nur klassische Windows-Outlook-Desktopversionen; kein neues Outlook, Outlook im Web oder macOS.
- Serien- und Besprechungssemantik folgt dem aktuell geöffneten `AppointmentItem`; Änderungen an einzelnen Vorkommen sollten vor dem Erstellen gespeichert sein.
- Es gibt bewusst keine global deaktivierte Zertifikatsprüfung; private CAs müssen im Windows-Zertifikatsspeicher vertrauenswürdig sein.
- Ein nachträglich manuell entfernter Link bleibt über die benutzerdefinierte UID-Eigenschaft als verknüpft erkannt.
