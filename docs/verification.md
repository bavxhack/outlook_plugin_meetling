# Technischer Prüfbericht

## Erneut geprüfter Ausgangsstand

main: c623582f61fcb165f5cb40338642bd5de3eab27d. Im Repository keine AGENTS.md
gefunden. COM-Identität unverändert. Bestätigte statische Befunde:

- GetCustomUI lieferte nur für Appointment XML; Explorer und Mail.Compose fehlten.
- ClassInterface.None ohne explizites Callback-Defaultinterface.
- Konferenzbutton führte synchron einen HTTP-Request aus; Startklasse besaß SettingsStore.
- IRibbonControl war leer. IDTExtensibility2 war als reines Dispatchinterface
  statt dual deklariert (Microsoft-Typelib-Flag 4160 = FDual | FDispatchable).
- Installer und WiX hatten Assemblyversion 1.0.0.0 fest eingebaut; Installer
  prüfte nicht die tatsächliche Paketarchitektur.
- WiX wich beim Installationspfad ab und enthielt nicht die Outlook-Interop-DLL.

Dies sind Codebefunde, keine nachgewiesene Ursache eines konkreten Outlook-
Ladefehlers. Ein entsprechender Outlooklauf lag nicht vor.

## COM-Verträge

Lokale, schmale Definitionen bleiben erhalten, damit keine zusätzliche, auf dem
Buildhost installierte Office-/Visual-Studio-PIA zur Voraussetzung wird.
Keine neuen GUIDs für Microsoft-Interfaces:

| Interface | IID | Vertrag |
| --- | --- | --- |
| IDTExtensibility2 | B65AD801-ABAF-11D0-BB8B-00A0C90F2744 | dual, DISPIDs 1–5, IDispatch-Objekte, ref SAFEARRAY(VARIANT) |
| IRibbonExtensibility | 000C0396-0000-0000-C000-000000000046 | dual, DISPID 1, BSTR Ein-/Ausgabe |
| IRibbonControl | 000C0395-0000-0000-C000-000000000046 | dual, Id=1/BSTR, Context=2/IDispatch, Tag=3/BSTR |
| IMeetlingRibbonCallbacks | 6D6E490A-EC2B-4F53-9AF3-32690E634D19 | eigenes Dispatch-Defaultinterface, TestPlugin=1, OpenSettings=2 |

Microsoft-Referenzen:
[IDTExtensibility2](https://learn.microsoft.com/en-us/dotnet/api/extensibility.idtextensibility2?view=visualstudiosdk-2022),
[IRibbonExtensibility](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.core.iribbonextensibility?view=office-pia),
[IRibbonControl](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.core.iribboncontrol?view=office-pia),
[Callback-Beispiele](https://learn.microsoft.com/en-us/office/vba/api/office.iribboncontrol).
Methodenreihenfolge und explizites Marshalling sind Bestandteil der lokalen Verträge.
Die Callbackparameter werden als IDispatch/object transportiert; die Basis
greift nicht auf IRibbonControl oder dessen Fensterkontext zu. Dadurch benötigt
die Callback-Typinformation keine registrierte Office-Typbibliothek, auch im
Office-freien Windows-CI-Test.
ClassInterface.None und ComDefaultInterface werden gemeinsam verwendet.
Die Assembly ist standardmäßig COM-unsichtbar; nur explizite Verträge und Klasse
werden exponiert. Keine AutoDual-Klassen.

## Automatisierte und manuelle Grenzen

Die Bearbeitungsumgebung ist Linux ohne .NET-/MSBuild-/PowerShell-Laufzeit,
Windows oder Outlook. Lokal möglich: statische Quellprüfung, XML-Parsing,
Workflow-Struktur, Diff-/Whitespaceprüfung. Nicht lokal ausgeführt:
Release-Builds, Core-Tests, PowerShell-Installer, Registry-/COM-Aktivierung und
Outlook-Abnahme. Der ergänzte Windows-CI-Test ist ausführbarer Prüfcode,
sein bloßes Vorhandensein gilt nicht als bestandenes Ergebnis.

Der Windows-Workflow führt x86/x64-Releasebuilds und paketierte Installationen
einschließlich Wiederholung aus; dann echte COM-Aktivierung, Default-IDispatch-
Namensauflösung und Invoke beider Callbacks mit ersetzter interner Dialogausgabe.
So bleiben CI-Läufe ohne modale Dialoge. Ribbon und Lifecycle werden zusätzlich
geprüft. Kein authentifizierter Nutzer, kein Outlook-Kalender und keine Meetling-
API erforderlich. Das Harness ersetzt keinen Outlook-UI-/Hosttest und ist
nicht Bestandteil der ausgelieferten Installationspakete.

Windows-/Outlook-Versionen und Bitness müssen in docs/smoke-test.md nach
tatsächlicher Ausführung ergänzt werden. WiX bleibt ausdrücklich nicht abgenommen.

## Windows-CI-Kontext

Der gehostete Runner läuft mit Administrator-Token. Die ersten isolierten
Aktivierungen unter diesem Konto meldeten HRESULT 0x80070002 trotz vorhandener
HKCU-Registrierung und lesbarem Paket. Das ist kein erfolgreicher COM-Nachweis.
Der endgültige Testweg führt daher Installation, Wiederinstallation, Diagnose
und Harness unter einem temporären lokalen Standardbenutzer mit geladenem
Benutzerprofil aus. Konto und Testverzeichnis werden danach entfernt. Diese
CI-Hilfsskripte werden nicht an Anwender ausgeliefert. Die Produktinstallation
bleibt ausschließlich unter HKCU; Outlook soll ohne Administratorrechte starten.

[Microsoft: COM und Benutzerregistrierung bei erhöhten Rechten](https://learn.microsoft.com/en-us/windows/win32/com/the-com-elevation-moniker).
