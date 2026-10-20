/* Test-only privileged harness. Never included in the production XPI. */
var { ExtensionCommon } = ChromeUtils.importESModule("resource://gre/modules/ExtensionCommon.sys.mjs");
var { setTimeout } = ChromeUtils.importESModule("resource://gre/modules/Timer.sys.mjs");
var smoke = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    return { smoke: { async run() {
      let outcome;
      try {
        const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
        let win;
        for (let i=0;i<600;i++) {
          win=Services.wm.getMostRecentWindow("mail:3pane");
          if(win?.setDefaultItemValues && win.document.readyState==="complete") break;
          await wait(100);
        }
        if (!win?.setDefaultItemValues) throw Error("Main window not ready");
        const { cal } = ChromeUtils.importESModule("resource:///modules/calendar/calUtils.sys.mjs");
        const manager=cal.manager.getCalendarManager();
        const calendar=manager.createCalendar("memory",Services.io.newURI("moz-memory-calendar://"));
        calendar.id=Services.uuid.generateUUID().toString();calendar.name="Meetling CI";
        manager.registerCalendar(calendar);
        win.getSelectedCalendar=()=>calendar;
        const scope={ChromeUtils,Services,Cc,Ci,URL};
        Services.scriptloader.loadSubScript(context.extension.rootURI.resolve("production.js"),scope);
        const implementation=new scope.meetlingHost(context.extension);
        const api=implementation.getAPI(context).meetlingHost;
        const target=await api.openDraft();
        const docs=[];
        for(const top of Services.wm.getEnumerator(null)) {
          for(const frame of top.document.querySelectorAll("iframe")) {
            if(frame.contentWindow?.calendarItem?.isEvent()) docs.push(frame.contentWindow);
          }
        }
        const editorWindow=docs.find(w=>w.calendarItem.calendar.id===calendar.id);
        if(!editorWindow) throw Error("Native calendar editor not found");
        const doc=editorWindow.document;
        if(!doc.getElementById("meetling-event-actions")) throw Error("Native Meetling button missing");
        const element=doc.getElementById("item-description"),editor=element.getHTMLEditor(element.contentWindow);
        editor.insertHTML('<p><b>Existing rich text</b></p>');
        doc.getElementById("item-location").value="Room 42";
        const start=editorWindow.calendarItem.startDate.icalString;
        await api.insert(target,'<p>Meetling<br><a href="https://call.example/r">Join</a></p><p>Nur für Moderatoren: <a href="https://host.example/room/dashboard?editConference_id=77">Edit</a></p>',"Meetling title");
        const content=editor.rootElement;
        if(content.querySelector('b')?.textContent!=="Existing rich text")throw Error("Formatting lost");
        if(content.querySelectorAll('a').length!==2)throw Error("Two clickable links not inserted");
        if(doc.getElementById("item-title").value!=="Meetling title")throw Error("New title missing");
        if(doc.getElementById("item-location").value!=="Room 42"||editorWindow.calendarItem.startDate.icalString!==start)throw Error("Appointment fields changed");
        let refused=false;try{await api.insert(target,"<p>duplicate</p>","Wrong");}catch{refused=true;}
        if(!refused)throw Error("Consumed target accepted twice");
        const entries=await calendar.getItemsAsArray(Ci.calICalendar.ITEM_FILTER_TYPE_EVENT,0,null,null);
        if(entries.length)throw Error("Draft was saved automatically");
        await api.vaultSet("https://host.example/meetling/","synthetic-token","CI user");
        if((await api.vaultGet("https://host.example/meetling/")).token!=="synthetic-token")throw Error("Vault roundtrip failed");
        if(await api.vaultGet("https://other.example/"))throw Error("Vault installation binding failed");
        await api.vaultDelete("https://host.example/meetling/");
        if(await api.vaultGet("https://host.example/meetling/"))throw Error("Vault delete failed");
        outcome={success:true,version:Services.appinfo.version,checks:["native draft opened","native button injected","rich text preserved","two clickable links","title set","other fields preserved","no automatic save","duplicate insertion blocked","encrypted vault roundtrip and binding"]};
      } catch(error) { outcome={success:false,error:String(error),stack:error.stack}; }
      await IOUtils.writeUTF8(Services.env.get("MEETLING_SMOKE_RESULT"),JSON.stringify(outcome,null,2));
      Services.startup.quit(Ci.nsIAppStartup.eForceQuit);
    } } };
  }
};
