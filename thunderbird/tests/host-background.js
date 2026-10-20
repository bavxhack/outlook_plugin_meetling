(async()=>{
  let stage="setup";
  try{
    await browser.smoke.setup();
    stage="normal-new-event";await browser.smoke.naturalDraft();
    stage="openDraft";
    const target=await browser.meetlingHost.openDraft();
    stage="seed";await browser.smoke.seed();
    stage="insert";await browser.meetlingHost.insert(target,'<p>Meetling<br><a href="https://call.example/r">Join</a></p><p>Nur für Moderatoren: <a href="https://host.example/room/dashboard?editConference_id=77">Edit</a></p>',"Meetling title");
    stage="verify";await browser.smoke.verify();
    let refused=false;try{await browser.meetlingHost.insert(target,"<p>duplicate</p>","Wrong");}catch{refused=true;}
    if(!refused)throw Error("Consumed target accepted twice");
    stage="vault";await browser.meetlingHost.vaultSet("https://host.example/meetling/","synthetic-token","CI user");
    if((await browser.meetlingHost.vaultGet("https://host.example/meetling/")).token!=="synthetic-token")throw Error("Vault roundtrip failed");
    await browser.meetlingHost.vaultSet("https://host.example/meetling/","replacement-token","Updated CI user");
    if((await browser.meetlingHost.vaultGet("https://host.example/meetling/")).token!=="replacement-token")throw Error("Vault replacement failed");
    if(await browser.meetlingHost.vaultGet("https://other.example/"))throw Error("Vault installation binding failed");
    await browser.meetlingHost.vaultDelete("https://host.example/meetling/");
    if(await browser.meetlingHost.vaultGet("https://host.example/meetling/"))throw Error("Vault delete failed");
    await browser.smoke.finish(JSON.stringify({success:true,checks:["normal new-event window shows button without forced API scanning","native draft opened","native button injected","rich text preserved","two clickable links","title set","other fields preserved","no automatic save","duplicate insertion blocked","encrypted vault roundtrip, replacement, deletion and binding","packaged XPI installed and active"]}));
  }catch(error){await browser.smoke.finish(JSON.stringify({success:false,error:String(error),stage,stack:error.stack}));}
})();
