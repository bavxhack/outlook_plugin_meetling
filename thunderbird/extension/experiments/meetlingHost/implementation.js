/* Minimal privileged bridge: draft editors and Thunderbird's encrypted login manager. */
var { ExtensionCommon } = ChromeUtils.importESModule("resource://gre/modules/ExtensionCommon.sys.mjs");
var { ExtensionSupport } = ChromeUtils.importESModule("resource:///modules/ExtensionSupport.sys.mjs");
var meetlingHost = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    const targets = new Map(), documents = new Set(), topWindows = new Map();
    const listenerId = context.extension.id + ":meetling-drafts";
    const origin = "chrome://meetling-thunderbird";
    let fireCommand = null;
    const isEditor = win => win?.location?.href === "chrome://calendar/content/calendar-item-iframe.xhtml";
    const ready = win => !win.closed && win.onLoad?.hasLoaded && win.calendarItem?.isEvent() && !win.gIsReadOnly &&
      !win.calendarItem.calendar?.readOnly && win.document.getElementById("item-description")?.getHTMLEditor;
    const capture = win => {
      if (!ready(win)) throw new Error("Bitte einen bearbeitbaren Kalendertermin öffnen.");
      for (const [id, target] of targets) if (target.win === win && target.item === win.calendarItem) return id;
      const id = Services.uuid.generateUUID().toString();
      targets.set(id, { win, item: win.calendarItem, title: win.document.getElementById("item-title").value, isNew: win.mode === "new" });
      return id;
    };
    const check = id => {
      const target = targets.get(id);
      if (!target || !ready(target.win) || target.win.calendarItem !== target.item)
        throw new Error("Der ursprüngliche Termin ist geschlossen oder nicht mehr bearbeitbar.");
      return target;
    };
    const attach = win => {
      if (!isEditor(win) || documents.has(win.document)) return;
      const grid = win.document.getElementById("event-grid");
      if (!grid) return;
      const box = win.document.createXULElement("hbox"); box.id = "meetling-event-actions";
      box.setAttribute("style", "padding: 8px; gap: 8px; align-items: center;");
      const button = win.document.createXULElement("button");
      button.setAttribute("label", "Meetling hinzufügen"); button.setAttribute("image", context.extension.rootURI.resolve("icon.svg"));
      button.setAttribute("style", "list-style-image: url('" + context.extension.rootURI.resolve("icon.svg") + "');");
      button.addEventListener("command", () => {
        try { fireCommand?.(capture(win)); }
        catch { Services.prompt.alert(win, "Meetling", "Bitte einen bearbeitbaren Kalendertermin öffnen."); }
      });
      box.append(button); grid.before(box); documents.add(win.document);
    };
    const scan = win => {
      try {
        attach(win);
        for (const frame of win.document.querySelectorAll("iframe")) {
          try { attach(frame.contentWindow); } catch { /* Other add-ons or web content are not editors. */ }
        }
      } catch { /* A window may close during scanning. */ }
    };
    ExtensionSupport.registerWindowListener(listenerId, {
      chromeURLs: ["chrome://messenger/content/messenger.xhtml", "chrome://calendar/content/calendar-event-dialog.xhtml"],
      onLoadWindow(win) {
        const handler = () => scan(win); topWindows.set(win, handler);
        win.addEventListener("load", handler, true); handler();
      },
      onUnloadWindow(win) {
        win.removeEventListener("load", topWindows.get(win), true); topWindows.delete(win);
        for (const [id, t] of targets) if (t.win.closed || t.win.top === win) targets.delete(id);
      }
    });
    const allEditors = () => {
      for (const win of Services.wm.getEnumerator(null)) scan(win);
      return [...documents].map(doc => doc.defaultView).filter(win => win && ready(win));
    };
    const installation = value => {
      const u = Services.io.newURI(value).QueryInterface(Ci.nsIURL);
      if (!u.schemeIs("https") || u.userPass || u.query || u.ref) throw new Error("Ungültige Installation.");
      return u.spec;
    };
    const logins = value => Services.logins.searchLoginsAsync({ origin, httpRealm: installation(value) });
    context.callOnClose({ close() {
      ExtensionSupport.unregisterWindowListener(listenerId);
      for (const [win, handler] of topWindows) win.removeEventListener("load", handler, true);
      for (const doc of documents) doc.getElementById("meetling-event-actions")?.remove();
      targets.clear(); documents.clear(); topWindows.clear();
    } });
    return { meetlingHost: {
      onCommand: new ExtensionCommon.EventManager({ context, name: "meetlingHost.onCommand", register(fire) {
        fireCommand = id => fire.async(id); return () => { fireCommand = null; };
      } }).api(),
      async openDraft() {
        const win = Services.wm.getMostRecentWindow("mail:3pane");
        if (!win?.createEventWithDialog) throw new Error("Bitte das Thunderbird-Hauptfenster öffnen.");
        const { CalEvent } = ChromeUtils.importESModule("resource:///modules/CalEvent.sys.mjs");
        const event = new CalEvent(); win.setDefaultItemValues(event);
        const marker = Services.uuid.generateUUID().toString(); event.setProperty("X-MEETLING-DRAFT", marker);
        win.createEventWithDialog(null, null, null, null, event);
        for (let attempt = 0; attempt < 200; attempt++) {
          const editor = allEditors().find(w => w.calendarItem.getProperty("X-MEETLING-DRAFT") === marker);
          if (editor) { editor.calendarItem.deleteProperty("X-MEETLING-DRAFT"); return capture(editor); }
          await new Promise(resolve => win.setTimeout(resolve, 50));
        }
        throw new Error("Der neue Kalendertermin konnte nicht geöffnet werden. Bitte einen Kalender einrichten.");
      },
      async checkTarget(id) { check(id); return true; },
      async insert(id, html, title) {
        const target = check(id), win = target.win;
        const element = win.document.getElementById("item-description");
        const editor = element.getHTMLEditor(element.contentWindow);
        editor.beginTransaction();
        try { editor.endOfDocument(); editor.insertHTML("<br><br>" + html); }
        finally { editor.endTransaction(); }
        const input = win.document.getElementById("item-title");
        if (target.isNew && input.value === target.title) {
          input.value = title; input.dispatchEvent(new win.Event("input", { bubbles: true }));
        }
        targets.delete(id);
        win.updateParentSaveControls?.();
        return true;
      },
      async vaultGet(value) {
        const found = (await logins(value))[0]; return found ? { token: found.password, name: found.username } : null;
      },
      async vaultSet(value, secret, name) {
        if (!secret || /[\s\x00-\x1f\x7f]/.test(secret) || secret.length > 16384) throw new Error("Ungültiges Token.");
        const realm = installation(value);
        const info = Cc["@mozilla.org/login-manager/loginInfo;1"].createInstance(Ci.nsILoginInfo);
        info.init(origin, null, realm, name, secret, "", "");
        const existing = (await logins(value))[0];
        if (existing) Services.logins.modifyLogin(existing, info); else await Services.logins.addLoginAsync(info);
      },
      async vaultDelete(value) { for (const entry of await logins(value)) Services.logins.removeLogin(entry); }
    } };
  }
};
