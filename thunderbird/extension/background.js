/* global Meetling, browser */
(() => {
  const host = browser.meetlingHost;
  let config = null, managed = false, configError = false, busy = false, auth = null, epoch = 0, loginSession = null, cache = null;
  const results = new Map();
  const initialized = (async () => {
    let policy = {};
    try { policy = await browser.storage.managed.get("configuration"); } catch { /* No managed storage configured. */ }
    const local = await browser.storage.local.get("configuration");
    managed = typeof policy.configuration === "string";
    const raw = managed ? (() => { try { return JSON.parse(policy.configuration); } catch { configError = true; return null; } })() : local.configuration;
    if (raw) { try { config = Meetling.settings(raw); } catch { configError = true; } }
    if (config) { const entry = await host.vaultGet(config.baseUrl); auth = entry ? { name: entry.name } : null; }
  })();
  async function openPanel(target = "") {
    await browser.windows.create({ type: "popup", url: browser.runtime.getURL("panel.html") + "?target=" + encodeURIComponent(target), width: 640, height: 740 });
  }
  browser.browserAction.onClicked.addListener(() => openPanel());
  host.onCommand.addListener(target => openPanel(target));
  const safeError = error => error instanceof Meetling.ApiError ? error.message : "Die Aktion konnte nicht abgeschlossen werden. Konfiguration, Anmeldung und geöffneten Termin prüfen.";
  const cancelLogin = () => {
    const session = loginSession; loginSession = null;
    if (session) { session.cleanup(); browser.windows.remove(session.windowId).catch(() => {}); }
  };
  async function startLogin() {
    if (loginSession) { await browser.windows.update(loginSession.windowId, { focused: true }); return; }
    if (!config) throw new Meetling.ApiError("Bitte zuerst die Meetling-Adresse einstellen.");
    const s = config, version = epoch, url = Meetling.endpoint(s, "login");
    const httpUrl = url.replace(/^https:/, "http:");
    const session = { windowId: -1, tabId: -1, upgraded: false, filters: new Set(), message: "Anmeldung läuft …" };
    loginSession = session;
    const headers = new Map();
    const valid = details => loginSession === session && details.tabId === session.tabId && details.frameId === 0;
    function beforeRequest(details) {
      if (!valid(details)) return {};
      if (details.url === httpUrl) {
        if (!session.upgraded) { session.upgraded = true; return { redirectUrl: url }; }
        session.message = "Die Anmeldung leitet wiederholt auf HTTP um. Bitte Serverkonfiguration prüfen.";
        return { cancel: true };
      }
      if (details.url !== url) return {};
      const filter = browser.webRequest.filterResponseData(details.requestId);
      session.filters.add(filter);
      let chunks = [], size = 0;
      filter.ondata = event => {
        filter.write(event.data);
        size += event.data.byteLength;
        if (size <= 1048576) chunks.push(new Uint8Array(event.data)); else chunks = [];
      };
      filter.onerror = () => { session.filters.delete(filter); headers.delete(details.requestId); };
      filter.onstop = async () => {
        filter.close(); session.filters.delete(filter);
        const status = headers.get(details.requestId); headers.delete(details.requestId);
        if (status !== 200 || size > 1048576 || epoch !== version || loginSession !== session) return;
        try {
          const decoder = new TextDecoder();
          const body = chunks.map(c => decoder.decode(c, { stream: true })).join("") + decoder.decode();
          // A login page is not a failed token response. Keep SSO/OTP interaction open.
          if (!body.trimStart().startsWith("{")) return;
          const entry = Meetling.login(body);
          if (epoch !== version || loginSession !== session) return;
          await host.vaultSet(s.baseUrl, entry.token, entry.name);
          if (epoch !== version || loginSession !== session) return;
          auth = { name: entry.name }; cache = null; cancelLogin();
        } catch { session.message = "Die Login-Antwort konnte nicht übernommen werden. Bitte Anmeldung erneut versuchen."; }
      };
      return {};
    }
    function received(details) { if (valid(details) && details.url === url) headers.set(details.requestId, details.statusCode); }
    function removed(id) { if (id === session.windowId && loginSession === session) { loginSession = null; session.cleanup(); } }
    const filter = { urls: [url + "*", httpUrl + "*"], types: ["main_frame"] };
    let timer;
    session.cleanup = () => {
      clearTimeout(timer);
      browser.webRequest.onBeforeRequest.removeListener(beforeRequest);
      browser.webRequest.onHeadersReceived.removeListener(received);
      browser.windows.onRemoved.removeListener(removed);
      for (const stream of session.filters) { try { stream.disconnect(); } catch { /* Already closed. */ } }
      session.filters.clear(); headers.clear();
    };
    browser.webRequest.onBeforeRequest.addListener(beforeRequest, filter, ["blocking"]);
    browser.webRequest.onHeadersReceived.addListener(received, filter);
    browser.windows.onRemoved.addListener(removed);
    try {
      // Open blank first so the exact login tab is known before its first request.
      const win = await browser.windows.create({ type: "popup", url: "about:blank", width: 1000, height: 800 });
      session.windowId = win.id; session.tabId = win.tabs?.[0]?.id ?? (await browser.tabs.query({ windowId: win.id }))[0].id;
      await browser.tabs.update(session.tabId, { url });
      timer = setTimeout(cancelLogin, 10 * 60 * 1000);
    } catch (error) { cancelLogin(); throw error; }
  }
  async function api(action, id) {
    if (!config) throw new Meetling.ApiError("Bitte zuerst die Meetling-Adresse einstellen.");
    const entry = await host.vaultGet(config.baseUrl);
    if (!entry) { auth = null; throw new Meetling.ApiError("Bitte zuerst anmelden."); }
    try { return await Meetling.request(config, action, entry.token, id); }
    catch (error) {
      if (error.invalidToken) { await host.vaultDelete(config.baseUrl); auth = null; cache = null; }
      throw error;
    }
  }
  async function act(message) {
    await initialized;
    if (message.action === "state") return { config, managed, configError, name: auth?.name ?? null, busy,
      login: loginSession?.message ?? null, result: results.get(message.target || message.operationId)?.conference ?? null };
    if (message.action === "cancelLogin") { epoch++; cancelLogin(); return {}; }
    if (busy) throw new Meetling.ApiError("Eine Aktion läuft bereits. Bitte kurz warten.");
    busy = true;
    try {
      switch (message.action) {
        case "save": {
          if (managed) throw new Meetling.ApiError("Die Konfiguration wird zentral verwaltet.");
          const next = Meetling.settings(message.config);
          epoch++; cancelLogin();
          await browser.storage.local.set({ configuration: next });
          config = next; configError = false; cache = null;
          const entry = await host.vaultGet(next.baseUrl); auth = entry ? { name: entry.name } : null;
          return {};
        }
        case "login": await startLogin(); return {};
        case "logout":
          epoch++; cancelLogin();
          try { if (auth) await api("logout"); }
          finally { if (config) await host.vaultDelete(config.baseUrl); auth = null; cache = null; }
          return {};
        case "servers":
          if (!cache || message.refresh || Date.now() - cache.time > 60000) cache = { servers: await api("servers"), time: Date.now() };
          return { servers: cache.servers };
        case "create": {
          if (results.has(message.target || message.operationId)) throw new Meetling.ApiError("Für diesen Vorgang wurde bereits eine Konferenz erstellt. Bitte den vorhandenen Link verwenden.");
          const list = await api("servers");
          if (!list.some(s => s.id === message.serverId)) throw new Meetling.ApiError("Der gewählte Server ist nicht mehr verfügbar.");
          const target = message.target || await host.openDraft();
          await host.checkTarget(target);
          const conference = await api("create", message.serverId);
          const key = message.target || message.operationId; results.set(key, { conference, target });
          try { await host.insert(target, conference.html, conference.title); return { conference, inserted: true }; }
          catch { return { conference, inserted: false }; }
        }
        default: throw new Meetling.ApiError("Unbekannte Aktion.");
      }
    } finally { busy = false; }
  }
  browser.runtime.onMessage.addListener(async (message, sender) => {
    if (sender.id !== browser.runtime.id || !sender.url?.startsWith(browser.runtime.getURL(""))) return;
    try { return { ok: true, ...await act(message) }; }
    catch (error) { return { ok: false, error: safeError(error) }; }
  });
})();
