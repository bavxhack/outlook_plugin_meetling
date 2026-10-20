# Outlook-Plugin-API und Administration

Vertragsquelle: https://github.com/H2-invent/jitsi-admin/wiki/Outlook-Plugin-API-Doku
(geprüft am 20.10.2026). Das Wiki beschreibt Controllerverhalten, keine live
getestete Installation. Implementierter Client, simulierte Vertragstests und
reale Abnahme sind getrennt zu betrachten.

## Administrierte Konfiguration

`%PROGRAMDATA%\Meetling\OutlookAddIn\config.json` hat Vorrang vor
`%LOCALAPPDATA%\Meetling\OutlookAddIn\config.json`. Ist die Verwaltungsdatei
vorhanden, aber beschädigt oder nicht lesbar, gilt kein GUI-Fallback. Outlook,
Ribbon, Einstellungen und Diagnose bleiben erreichbar. Die GUI zeigt die
Organisationsvorgabe schreibgeschützt an. Ohne Verwaltungsdatei kann der Nutzer
die HTTPS-Installationsadresse speichern. JSON enthält keine Secrets.

`config.example.json` ist eine inaktive Vorlage. Mit Gruppenrichtlinien-
Einstellungen → Windows-Einstellungen → Dateien die geprüfte JSON-Datei nach
ProgramData kopieren; Ordner bei Bedarf administrativ anlegen. ACLs: SYSTEM und
Administratoren dürfen schreiben, Benutzer dürfen lesen; Benutzer dürfen weder
Datei noch übergeordneten Ordner ersetzen. Der Add-in-Installer legt keine aktive
Verwaltungsdatei an und überschreibt keine vorhandene Konfiguration.

`baseUrl` umfasst den Installationspfad. `https://host/meetling/` führt zu
`https://host/meetling/room/outlook/plugin/server`. API-Pfade sind relativ;
führende Slashes, Traversierung, Query und absolute Fremdadressen werden abgelehnt.
Timeout: 1–300 Sekunden. Änderungen werden über „Konfiguration neu laden“ bzw.
die Einstellungen übernommen. Laufende Antworten werden ungültig, Browserlogin
abgebrochen, Cache und In-Memory-Token entfernt. Für das Ziel erneut anmelden.
Kein automatischer Logout beim früheren Server.

## Anmeldung und Token

„Anmelden“ öffnet ein eigenes sichtbares Edge-WebView2-Fenster erst nach Klick.
Der Standard-Login-Endpunkt ist nun `/room/outlook/plugin/outlooklogin`
(statt `/room/outlook/plugin/login`). Bei bestehenden Konfigurationsdateien
`endpoints.login` auf `outlooklogin` setzen und die Konfiguration neu laden.
Der Schlüssel `login` bleibt unverändert; nur sein Pfadwert ändert sich.
Es navigiert zum konfigurierten Login-Endpunkt, erlaubt HTTPS-Weiterleitungen und
MFA im gleichen Fenster und übernimmt nur dessen exakte JSON-Netzwerkantwort
mittels WebResourceResponseReceived/GetContentAsync. Fremde Origins, abweichende
Pfade, Queries, Fragmente und fremde Antworten liefern kein Token. Abbruch,
Schließen, fünf Minuten Timeout, Reload, Logout und Shutdown machen alte
Antworten ungültig. Höchstens eine Anmeldung ist aktiv.

WebView2 Evergreen Runtime muss administrativ bereitgestellt werden, zum Beispiel
über die Microsoft-Unternehmensverteilung. Das ZIP liefert die verwalteten
WebView2-Assemblies und den zur Outlook-Bitness passenden nativen Loader mit;
es enthält die Browserruntime nicht. Ohne Runtime funktionieren Laden, Ribbon
und Einstellungen weiter; nur Login meldet die fehlende Browserkomponente.
Referenz: https://learn.microsoft.com/microsoft-edge/webview2/concepts/distribution

Jeder Login nutzt einen neuen, ausschließlich für das Add-in bestimmten
WebView2-Profilordner unter `%LOCALAPPDATA%\Meetling\OutlookAddIn\WebView2\`.
Damit werden Cookies einer früheren Add-in-Anmeldung nicht wiederverwendet.
Verwaiste Profilordner können bei geschlossenem Outlook administrativ entfernt
werden. API-Logout beendet keine Symfony-Web-Sitzung; es gibt keinen erfundenen
SSO-Logout. Systembrowserprofile werden weder gelesen noch verändert. SSO-Popups werden
über WebView2 NewWindowRequested/GetDeferral/NewWindow im gleichen isolierten
Browserkontext geöffnet, einschließlich opener-Verbindung. Sie werden bei Ende
des Anmeldevorgangs geschlossen. Nur HTTPS und leere about:blank-Dokumente sind
zulässig; HTTP und fremde Protokolle bleiben gesperrt.

Bei ausbleibender automatischer Rückkehr nach abgeschlossener MFA kann der Nutzer
„Zurück zu Meetling“ wählen. Das navigiert im bestehenden Browserkontext erneut
zum exakt konfigurierten outlooklogin-Endpunkt. Es umgeht weder OTP noch die
Prüfung der Tokenantwort, erfindet keine Callback-URLs und übernimmt keine Cookies
in HttpClient. Vor Abschluss der MFA kann der Button wieder zur Anmeldung führen.

Im Fenster stehen Navigationsfehler und der Anmeldeschritt. diagnose.log enthält
BrowserLogin-Einträge für Start, Weiterleitungskategorie, WebErrorStatus,
HTTP-Fehlerstatus, Popup und validierte Tokenantwort. URLs, Queryparameter,
SSO-Codes, OTP, Zugangsdaten, Cookies und Antwortinhalte werden nicht protokolliert.
Diagnosekategorien sind plugin-login, installation, identity-provider,
empty-document und blocked. Die Diagnose liest keine Formularinhalte.
Referenz: https://learn.microsoft.com/dotnet/api/microsoft.web.webview2.core.corewebview2.newwindowrequested

Token und Anzeigename werden gemeinsam per Windows DPAPI CurrentUser geschützt
und atomar in `plugin-token-<SHA256-Installation>.bin` gespeichert. Installation
(einschließlich Unterpfad) ist im verschlüsselten Datensatz und in der DPAPI-
Entropie gebunden. Kein Klartextfallback, kein Import von `api-key.bin`, kein
JWT-/Längen-/Ablaufzeitmodell, kein Refresh. Der Anzeigename ist keine Benutzer-ID.
Ein gespeichertes Token wird beim Outlook-Start geladen, erst beim nächsten
API-Aufruf geprüft und öffnet keinen Browser. Bei ungültigem Token erneut anmelden.

## Antwortübernahme nach MFA (Version 1.3.1)

HTTP 200 bestätigt noch keine erfolgreiche Tokenübernahme. Der Browser kann
bei abgebrochener Navigation oder Download keinen Antwortstream liefern.
Fehlender Stream und Laufzeitfehler schließen das Anmeldefenster nicht mehr;
die bestehende Sitzung bleibt für „Zurück zu Meetling“ erhalten. Es erfolgen
keine automatischen Wiederholungen. Nur die validierte JSON-Antwort vom exakten
HTTPS-Endpunkt kann eine Anmeldung abschließen. UI-Änderungen nach asynchronem
Lesen laufen über den UI-Dispatcher. Downloads im Anmeldefenster werden
abgebrochen, damit eine Tokenantwort nicht als unverschlüsselte Datei landet.

BrowserLogin protokolliert jetzt zusätzlich die feste Fehlerphase,
Exceptiontyp/HRESULT und einen festen Sperrgrund (z. B. http, data, blob,
userinfo). Es werden weiterhin keine Adressen oder Antwortinhalte ausgegeben.

Ein eng begrenzter Rückleitungsfall wird behandelt: Leitet die Installation
auf eine HTTP-Adresse mit exakt demselben Host und demselben Login-Pfad wie die
konfigurierte HTTPS-Adresse zurück, ohne Query, Fragment, Zugangsdaten oder
abweichenden Port, wird HTTP weiterhin blockiert. Stattdessen wird einmalig
der bereits konfigurierte HTTPS-Login-Endpunkt im selben Browser aufgerufen.
Parameter werden nicht übernommen. Andere HTTP-Rückleitungen bleiben blockiert;
es gibt keine pauschale HTTP-Freigabe und keine Endlosschleife. Dies ist eine
abgesicherte Behandlung eines möglichen Proxy-/Scheme-Problems, kein Nachweis,
dass genau dieses Problem beim gemeldeten Live-Versuch vorlag.

## API-Transport nach erfolgreicher Anmeldung (Version 1.3.2)

WebView2 und HttpClient verwenden unterschiedliche Netzwerkstacks. Ein Erfolg
im Anmeldebrowser beweist deshalb noch keine funktionierende API-Verbindung.
Ab Version 1.3.3 bietet ausschließlich der API-HttpClient TLS 1.2 und TLS 1.3
über HttpClientHandler.SslProtocols an. Damit begrenzen möglicherweise alte
TLS-Vorgaben des Outlook-/CLR-Hosts diesen Client nicht mehr auf TLS 1.0.
Windows handelt eine unterstützte Version aus; SSL3/TLS1.0/TLS1.1 werden nicht
angeboten. Betriebssystemrichtlinien und Zertifikatsprüfung bleiben wirksam.
Globale ServicePointManager-Vorgaben des Outlook-Prozesses werden nicht geändert.
Es werden keine eigenen Zertifikatsprüfungen eingesetzt. Systemproxy, Redirect-Sperre und Verzicht auf
Cookies bleiben erhalten. Keine automatische Wiederholung.
Referenzen: https://learn.microsoft.com/dotnet/api/system.net.http.httpclienthandler.sslprotocols
und https://learn.microsoft.com/dotnet/framework/network-programming/tls
(TLS-1.3-Konstante 12288 für ältere Referenzassemblies).

ApiRequest-Einträge melden die feste Aktion (servers/create/logout), Start,
HTTP-Status und bei Serverlisten die Anzahl. Transportfehler liefern eine
begrenzte innere Exceptionkette mit Typ, HRESULT, WebExceptionStatus und
SocketErrorCode. Keine Exceptiontexte, URLs, Header oder Antwortinhalte.
DNS-/Proxy-/TLS-/Timeoutfehler werden verständlich angezeigt. Ein Transportfehler
löscht kein gespeichertes Token; ein erneuter manueller Serverabruf genügt für
einen weiteren Versuch. Ein Zertifikats- oder Proxyproblem muss administrativ
behoben werden, nicht durch Abschalten der Zertifikatsprüfung.

## HTTP und Konferenz

Konvention: GET Login (Browser), GET Server und Logout (HttpClient), POST
createConference. Wiki-Beispiele nennen bei Login/Logout auch POST, die Übersicht
empfiehlt GET; wir folgen der Übersicht und wechseln Methoden niemals automatisch.
Aktualisierung laut Nutzer vom 20.10.2026: Server/Logout benötigen jetzt
`Authorization: Bearer <token>`. Erstellung verwendet `token: <token>` und genau
ein Formularfeld `server_id`. Die Bearer-Ergänzung betrifft die Authorization-
Header. Kein JSON-Body oder Outlook-Inhalt. Gespeicherte Tokens bleiben opak
und können weiterverwendet werden; der Client ergänzt das Präfix pro Request. HttpClient
verwendet keine Cookies und keine Redirects; Header gelten nur für den Request.
TLS-Zertifikatsprüfungen bleiben aktiv.

Servercache maximal zwei Minuten, verworfen bei Anmeldung, Logout und Reload.
Null Server verhindern Erstellung; ein Server wird direkt verwendet, ab zwei
Servern ist eine bestätigte Auswahl erforderlich. Namen und numerische IDs sind
getrennt, IDs erscheinen zusätzlich im Dialog. Große Ganzzahlen bleiben exakt;
String-IDs werden bewusst abgewiesen. Abbruch sendet keinen Create-Request.

Konferenzantwort muss genau ein Objekt mit HTTPS-URL und Titel enthalten.
Die URL wird unverändert eingefügt; es gibt keine erfundene UID. Vor Netzwerk
wird der Zielentwurf erfasst. Im Explorer wird ein nativer Besprechungsentwurf
geöffnet, im bearbeitbaren Termin-/Mailfenster das vorhandene Element ergänzt.
Word-Range-Einfügung hängt den Linkblock hinten an und erhält Formatierung und
native Felder. Kein Save, Send oder spätere Synchronisation. Der Titel wird nur
bei neu angelegten Entwürfen gesetzt. Empfangene Besprechungen und geschützte
Editoren werden abgewiesen.

Globaler Beschäftigtzustand verhindert Doppelklicks. Nicht geheime
`MeetlingCreationState`-Metadaten blockieren erneute Erstellung im gleichen
Element, sobald ein Request begonnen wurde; `MeetlingParticipantLink` hält den
zurückgegebenen Link. Die API hat keine Idempotenz: Bei unklarem Ausgang erfolgt
kein Retry. Der konservative Pending-Marker bleibt auch nach fehlgeschlagener
Erstellung erhalten; vor einem bewusst neuen Entwurf den Server prüfen. Nach
Erstellung und gescheiterter Einfügung wird ein kopierbarer Link angezeigt.

„Abmelden“ erklärt den Widerruf für alle Clients mit identischem Token.
GET /logout öffnet keinen Browser. Nur `token:null` bestätigt Widerruf;
`No user loged found` entfernt die lokale Kopie ohne neuen Widerruf zu behaupten.
Bei Netzwerkfehler kann erneut versucht oder ausdrücklich nur lokal abgemeldet
werden. Lokales Löschen bedeutet keinen serverseitigen Widerruf.

## Externe Integrationsgrenzen – offen bis Live-Abnahme

Die Installation muss einen belegten SSO-Anmeldeweg mit Rückkehr zum exakten
Login-Endpunkt liefern. Dieser führt selbst keine Anmeldung durch. Liefert er
`No user loged in`, wird dies erklärt; es wird kein SSO-Endpunkt oder Callback
erfunden. Eingebettete Browser müssen vom tatsächlichen SSO unterstützt werden.

Die Moderatoridentität muss serverseitig dem Tokenbenutzer entsprechen. Der
Controller setzt den Moderator über Symfony getUser(), während Serverrechte über
das Token bestimmt werden. Ein Tokenheader etabliert allein keinen Security-
Benutzer. Erforderlich ist zum Beispiel ein Token-Authenticator oder die direkte
Moderatorzuordnung aus dem Tokenbenutzer. Keine Backendänderung, keine Cookie-
Übernahme im Plugin. JSON-Erfolg beweist diese Voraussetzung nicht. Besonders
nach Outlook-Neustart ohne Browsersitzung muss die Moderatoridentität geprüft
werden. Aktuell nicht live bestätigt.
