using System;
using System.Net;
using System.Net.Sockets;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;
namespace Meetling.Core;
public static class TransportFailure
{
    public static string Category(Exception error)
    {
        for (Exception? current = error; current != null; current = current.InnerException)
        {
            if (current is AuthenticationException) return "tls";
            if (current is WebException web)
            {
                switch (web.Status)
                {
                    case WebExceptionStatus.SecureChannelFailure:
                    case WebExceptionStatus.TrustFailure: return "tls";
                    case WebExceptionStatus.NameResolutionFailure: return "dns";
                    case WebExceptionStatus.ProxyNameResolutionFailure: return "proxy";
                    case WebExceptionStatus.Timeout: return "timeout";
                }
            }
            if (current is SocketException socket)
            {
                if (socket.SocketErrorCode == SocketError.HostNotFound || socket.SocketErrorCode == SocketError.NoData) return "dns";
                if (socket.SocketErrorCode == SocketError.TimedOut) return "timeout";
            }
            if (current is TaskCanceledException) return "timeout";
        }
        return "network";
    }
    public static string UserMessage(Exception error)
    {
        switch (Category(error))
        {
            case "tls": return "Die sichere Verbindung ist fehlgeschlagen. Bitte Windows-Zertifikatsvertrauen, TLS-Richtlinien und gegebenenfalls den Unternehmensproxy prüfen.";
            case "dns": return "Der Servername konnte nicht aufgelöst werden. Bitte Netzwerk, DNS und Installationsadresse prüfen.";
            case "proxy": return "Der konfigurierte Netzwerkproxy ist nicht erreichbar oder konnte nicht aufgelöst werden.";
            case "timeout": return "Zeitüberschreitung beim API-Aufruf. Es erfolgt keine automatische Wiederholung.";
            default: return "Die API-Verbindung ist fehlgeschlagen. Bitte die ApiRequest-Einträge im Diagnose-Log prüfen.";
        }
    }
    public static string SafeDetails(Exception error)
    {
        var result = new StringBuilder("Category=" + Category(error));
        int depth = 0;
        for (Exception? current = error; current != null && depth < 8; current = current.InnerException, depth++)
        {
            result.Append(" Cause").Append(depth).Append('=').Append(current.GetType().FullName)
                .Append(" HRESULT=0x").Append(current.HResult.ToString("X8"));
            if (current is WebException web) result.Append(" WebStatus=").Append(web.Status);
            if (current is SocketException socket) result.Append(" SocketCode=").Append(socket.SocketErrorCode);
        }
        // No exception messages, URLs, headers, certificates or response bodies.
        return result.ToString();
    }
}
