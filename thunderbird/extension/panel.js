/* global Meetling, browser */
(() => {
  const $ = id => document.getElementById(id);
  const target = new URL(location.href).searchParams.get("target") || "";
  const operationId = sessionStorage.getItem("meetling-operation") || crypto.randomUUID();
  sessionStorage.setItem("meetling-operation", operationId);
  let state, running = false, loaded = false;
  const message = async (action, values = {}) => {
    const result = await browser.runtime.sendMessage({ action, target, operationId, ...values });
    if (!result?.ok) throw new Error(result?.error || "Meetling ist nicht erreichbar.");
    return result;
  };
  const status = (text, error = false) => { $("status").textContent = text; $("status").className = error ? "error" : ""; };
  const result = (conference, inserted) => {
    $("selection").hidden = true; $("result").hidden = false;
    $("insert-status").textContent = inserted ? "In den Termin eingefügt. Bitte in Thunderbird prüfen und speichern." : "Die Konferenz wurde erstellt. Du kannst die Links hier übernehmen; bitte keine zweite Konferenz erstellen.";
    $("links").replaceChildren();
    for (const [label, url] of [["Teilnehmen", conference.url], ["Konferenz in Meetling bearbeiten (nur für Moderatoren)", conference.editUrl]]) {
      if (!url) continue;
      const p = document.createElement("p"), a = document.createElement("a"); a.textContent = label; a.href = url; a.target = "_blank"; a.rel = "noopener noreferrer"; p.append(a); $("links").append(p);
    }
    $("result-text").value = conference.text;
    if (!conference.editUrl) status("Der Server hat keine gültige Konferenz-ID geliefert. Der Bearbeitungslink fehlt.", true);
  };
  async function action(fn) {
    if (running) return;
    running = true; document.querySelectorAll("button").forEach(b => b.disabled = true);
    try { await fn(); } catch (error) { status(error.message, true); }
    finally { running = false; document.querySelectorAll("button").forEach(b => b.disabled = false); if (state?.managed) $("save").disabled = true; }
  }
  async function loadServers(refresh = false) {
    const response = await message("servers", { refresh });
    $("servers").replaceChildren(); $("selection").hidden = false;
    const names = new Map(); response.servers.forEach(s => names.set(s.name, (names.get(s.name) || 0) + 1));
    response.servers.sort((a, b) => a.name.localeCompare(b.name) || a.id.localeCompare(b.id));
    for (const server of response.servers) {
      const button = document.createElement("button"); button.className = "server";
      const icon = document.createElement("img"); icon.src = "icon.svg"; icon.alt = "";
      const text = document.createElement("span"); text.textContent = server.name;
      if (names.get(server.name) > 1) { const id = document.createElement("small"); id.textContent = "Server-ID: " + server.id; text.append(id); }
      button.append(icon, text); button.title = server.name + " · ID " + server.id;
      button.addEventListener("click", () => action(async () => {
        status("Konferenz wird erstellt …"); const r = await message("create", { serverId: server.id }); status(""); result(r.conference, r.inserted);
      }));
      $("servers").append(button);
    }
    if (!response.servers.length) status("Für dein Konto sind keine Konferenzserver verfügbar.");
    loaded = true;
  }
  async function update() {
    state = await message("state");
    $("name").textContent = state.name ? "Angemeldet als " + state.name : "Noch nicht angemeldet";
    $("login").hidden = !!state.name || !!state.login; $("logout").hidden = !state.name; $("cancel-login").hidden = !state.login;
    if (state.name && !loaded && !running && !state.result) { status(""); await loadServers(); }
    if (!state.name) { $("selection").hidden = true; loaded = false; }
    if (state.login) status(state.login);
    if (state.result && $("result").hidden) result(state.result, false);
  }
  $("config-form").addEventListener("submit", event => { event.preventDefault(); action(async () => {
    const config = JSON.parse($("config-json").value); config.baseUrl = $("base-url").value;
    await message("save", { config }); loaded = false; status("Einstellungen gespeichert."); await update();
  }); });
  $("import").addEventListener("change", () => action(async () => {
    const file = $("import").files[0]; if (!file) return; if (file.size > 65536) throw new Error("Konfigurationsdatei zu groß.");
    const s = Meetling.settings(JSON.parse(await file.text())); $("base-url").value = s.baseUrl; $("config-json").value = JSON.stringify(s, null, 2); status("Konfiguration eingelesen. Zum Übernehmen speichern.");
  }));
  $("login").addEventListener("click", () => action(async () => { await message("login"); await update(); }));
  $("cancel-login").addEventListener("click", () => action(async () => { await message("cancelLogin"); status("Anmeldung abgebrochen."); await update(); }));
  $("logout").addEventListener("click", () => action(async () => { try { await message("logout"); status("Abgemeldet."); } finally { await update(); } }));
  $("refresh").addEventListener("click", () => action(() => loadServers(true)));
  (async () => {
    await update();
    $("base-url").value = state.config?.baseUrl || ""; $("config-json").value = JSON.stringify(state.config || Meetling.defaults, null, 2);
    $("settings").open = !state.config || state.configError || new URL(location.href).searchParams.has("settings");
    $("managed").textContent = state.managed ? "Die Konfiguration wird zentral verwaltet." : "";
    if (state.managed) $("config-form").querySelectorAll("input, textarea, button").forEach(e => e.disabled = true);
    if (state.configError) status("Die gespeicherte oder zentral vorgegebene Konfiguration ist ungültig.", true);
    setInterval(() => { if (!running) update().catch(() => {}); }, 1500);
  })().catch(error => status(error.message, true));
})();
