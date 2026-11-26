/* Meetling API contract shared by the Thunderbird UI and tests. No Thunderbird dependencies. */
var Meetling = (() => {
  class ApiError extends Error {
    constructor(message, invalidToken = false) { super(message); this.invalidToken = invalidToken; }
  }
  const defaults = { baseUrl: "", apiBasePath: "api/v1/outlook/plugin/", loginBasePath: "room/outlook/plugin/",
    endpoints: { login: "outlooklogin", servers: "server", createConference: "createConference", logout: "logout" }, timeoutSeconds: 30 };
  function https(value) {
    const url = new URL(value);
    if (url.protocol !== "https:" || url.username || url.password) throw new Error("Bitte eine HTTPS-Adresse ohne Zugangsdaten verwenden.");
    return url;
  }
  function relative(value) {
    if (typeof value !== "string" || !value || /[^a-zA-Z0-9_\-/]/.test(value) || value.startsWith("/") || value.includes("//")) throw new Error("Ungültiger relativer API-Pfad.");
    return value.replace(/\/$/, "");
  }
  function settings(input) {
    const s = { ...defaults, ...input, endpoints: { ...defaults.endpoints, ...input?.endpoints } };
    const base = https(s.baseUrl.trim());
    if (base.search || base.hash) throw new Error("Die Basisadresse darf keine Parameter enthalten.");
    s.baseUrl = base.href.replace(/\/*$/, "/");
    if (["room/outlook/plugin", "room/outlook/plugin/"].includes(s.apiBasePath)) {
      s.apiBasePath = defaults.apiBasePath;
      if (s.endpoints.login === "login") s.endpoints.login = "outlooklogin";
    }
    relative(s.apiBasePath); relative(s.loginBasePath);
    Object.values(s.endpoints).forEach(relative);
    if (!Number.isInteger(s.timeoutSeconds) || s.timeoutSeconds < 1 || s.timeoutSeconds > 300) throw new Error("Timeout muss zwischen 1 und 300 Sekunden liegen.");
    return s;
  }
  function endpoint(s, action) {
    const prefix = action === "login" ? s.loginBasePath : s.apiBasePath;
    const key = { login: "login", servers: "servers", create: "createConference", logout: "logout" }[action];
    if (!key) throw new Error("Unbekannte API-Aktion.");
    return s.baseUrl + relative(prefix) + "/" + relative(s.endpoints[key]);
  }
  function token(value) {
    if (typeof value !== "string" || !value || value.length > 16384 || /[\s\x00-\x1f\x7f]/.test(value)) throw new Error("Ungültige Anmeldeantwort.");
    return value;
  }
  function parse(body) {
    if (typeof body !== "string" || body.length > 1048576) throw new ApiError("API-Antwort zu groß oder ungültig.");
    let result;
    try {
      // Gecko 135+: retain the exact integer spelling before IEEE-754 conversion loses precision.
      result = JSON.parse(body, (key, value, context) => key === "id" && typeof value === "number" ? context.source : value);
    } catch { throw new ApiError("Meetling liefert kein gültiges API-JSON."); }
    if (result?.error === true) throw new ApiError("Meetling hat die Anfrage abgelehnt.", ["No user loged in", "No user loged found", "User not found"].includes(result.reason));
    return result;
  }
  function required(value) {
    if (typeof value !== "string" || !value.trim()) throw new ApiError("Ein Pflichtfeld in der Meetling-Antwort fehlt.");
    return value;
  }
  function login(body) { const j = parse(body); return { token: token(j?.token), name: required(j?.name) }; }
  function servers(j) {
    if (!Array.isArray(j)) throw new ApiError("Ungültige Serverliste.");
    const seen = new Set();
    return j.map(s => {
      if (!/^-?\d+$/.test(s?.id) || seen.has(String(s.id))) throw new ApiError("Ungültige oder doppelte Server-ID.");
      seen.add(String(s.id)); return { id: String(s.id), name: required(s.name) };
    });
  }
  function escape(value) { return String(value).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]); }
  function conference(s, j) {
    if (!Array.isArray(j) || j.length !== 1) throw new ApiError("Erwartet wird genau eine Konferenz.");
    const c = j[0], url = https(required(c?.url)).href, title = required(c?.title);
    const id = typeof c.id === "string" && /^\d+$/.test(c.id) && /[1-9]/.test(c.id) ? c.id : null;
    const editUrl = id ? s.baseUrl + "room/dashboard?editConference_id=" + encodeURIComponent(id) : null;
    const text = "Meetling-Konferenz\nTeilnehmen: " + url + (editUrl ? "\n\nKonferenz in Meetling bearbeiten (nur für Moderatoren):\n" + editUrl : "");
    const html = '<p>Meetling-Konferenz<br>Teilnehmen: <a href="' + escape(url) + '">' + escape(url) + '</a></p>' +
      (editUrl ? '<p>Konferenz in Meetling bearbeiten (nur für Moderatoren):<br><a href="' + escape(editUrl) + '">' + escape(editUrl) + '</a></p>' : "");
    return { url, title, id, editUrl, text, html };
  }
  async function request(s, action, secret, id, transport = fetch) {
    const controller = new AbortController(), timer = setTimeout(() => controller.abort(), s.timeoutSeconds * 1000);
    try {
      const init = { method: action === "create" ? "POST" : "GET", headers: { Authorization: "Bearer " + token(secret), Accept: "application/json" },
        redirect: "error", credentials: "omit", cache: "no-store", signal: controller.signal };
      if (action === "create") {
        if (!/^-?\d+$/.test(id)) throw new Error("Server-ID fehlt.");
        init.headers["Content-Type"] = "application/x-www-form-urlencoded";
        init.body = new URLSearchParams({ server_id: id }).toString();
      }
      const r = await transport(endpoint(s, action), init);
      if (!r.ok) throw new ApiError("Meetling HTTP " + r.status + ".", [401, 403].includes(r.status));
      const json = parse(await r.text());
      if (action === "servers") return servers(json);
      if (action === "create") return conference(s, json);
      if (json?.token !== null) throw new ApiError("Abmeldung wurde vom Server nicht bestätigt.");
      return null;
    } catch (error) {
      if (error instanceof ApiError) throw error;
      throw new ApiError(action === "create" ? "Ergebnis der Erstellung unbekannt. Bitte zuerst in Meetling prüfen, bevor du erneut erstellst." : "Meetling konnte nicht erreicht werden. HTTPS-Adresse, Netzwerk und Routen prüfen.");
    } finally { clearTimeout(timer); }
  }
  return { defaults, settings, endpoint, token, parse, login, servers, conference, request, escape, ApiError };
})();
if (typeof module !== "undefined") module.exports = Meetling;
