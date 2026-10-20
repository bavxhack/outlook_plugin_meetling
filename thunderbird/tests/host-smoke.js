/* Test-only privileged harness. Never included in the production XPI. */
var { ExtensionCommon } = ChromeUtils.importESModule("resource://gre/modules/ExtensionCommon.sys.mjs");
var { setTimeout } = ChromeUtils.importESModule("resource://gre/modules/Timer.sys.mjs");
var smoke = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    let calendar, editorWindow, start;
    return { smoke: {
      async setup() {
        const wait=ms=>new Promise(resolve=>setTimeout(resolve,ms));
        let win;
        for(let i=0;i<600;i++) {
          win=Services.wm.getMostRecentWindow("mail:3pane");
          if(win?.setDefaultItemValues && win.document.readyState==="complete")break;
          await wait(100);
        }
        if(!win?.setDefaultItemValues)throw Error("Main window not ready");
        const {AddonManager}=ChromeUtils.importESModule("resource://gre/modules/AddonManager.sys.mjs");
        const addon=await AddonManager.getAddonByID("meetling-thunderbird@h2-invent.com");
        if(!addon?.isActive)throw Error("Packaged production XPI is not active");
        const {cal}=ChromeUtils.importESModule("resource:///modules/calendar/calUtils.sys.mjs");
        calendar=cal.manager.createCalendar("memory",Services.io.newURI("moz-memory-calendar://"));
        calendar.id=Services.uuid.generateUUID().toString();calendar.name="Meetling CI";
        cal.manager.registerCalendar(calendar);win.getSelectedCalendar=()=>calendar;
      },
      async seed() {
        for(const top of Services.wm.getEnumerator(null)) {
          for(const frame of top.document.querySelectorAll("iframe")) {
            if(frame.contentWindow?.calendarItem?.calendar?.id===calendar.id)editorWindow=frame.contentWindow;
          }
        }
        if(!editorWindow)throw Error("Native calendar editor not found");
        const doc=editorWindow.document;
        if(!doc.getElementById("meetling-event-actions"))throw Error("Native Meetling button missing");
        const element=doc.getElementById("item-description"),editor=element.getHTMLEditor(element.contentWindow);
        editor.insertHTML('<p><b>Existing rich text</b></p>');
        doc.getElementById("item-location").value="Room 42";start=editorWindow.calendarItem.startDate.icalString;
      },
      async verify() {
        const doc=editorWindow.document,element=doc.getElementById("item-description"),editor=element.getHTMLEditor(element.contentWindow);
        if(editor.rootElement.querySelector('b')?.textContent!=="Existing rich text")throw Error("Formatting lost");
        if(editor.rootElement.querySelectorAll('a').length!==2)throw Error("Two clickable links not inserted");
        if(doc.getElementById("item-title").value!=="Meetling title")throw Error("New title missing");
        if(doc.getElementById("item-location").value!=="Room 42"||editorWindow.calendarItem.startDate.icalString!==start)throw Error("Appointment fields changed");
        const entries=await calendar.getItemsAsArray(Ci.calICalendar.ITEM_FILTER_TYPE_EVENT,0,null,null);
        if(entries.length)throw Error("Draft was saved automatically");
      },
      async finish(report) {
        const outcome={...JSON.parse(report),version:Services.appinfo.version};
        if (!outcome.success) outcome.console=Services.console.getMessageArray().map(m=>m.message).slice(-30);
        await IOUtils.writeUTF8(Services.env.get("MEETLING_SMOKE_RESULT"),JSON.stringify(outcome,null,2));
        Services.startup.quit(Ci.nsIAppStartup.eForceQuit);
      }
    }};
  }
};
