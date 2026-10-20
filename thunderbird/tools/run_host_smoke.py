"""Run a packaged test harness in a disposable, real Thunderbird profile."""
from pathlib import Path
import argparse
import json
import os
import subprocess
import tempfile
import zipfile
p=argparse.ArgumentParser()
p.add_argument('--binary',required=True)
p.add_argument('--result',required=True)
a=p.parse_args()
root=Path(__file__).resolve().parents[1]
result=Path(a.result).resolve()
result.parent.mkdir(parents=True,exist_ok=True)
if result.exists(): result.unlink()
with tempfile.TemporaryDirectory(prefix='meetling-thunderbird-') as tmp:
    profile=Path(tmp)
    extensions=profile/'extensions';extensions.mkdir()
    manifest={"manifest_version":2,"name":"Meetling CI harness","version":"1.0","browser_specific_settings":{"gecko":{"id":"meetling-ci@example.invalid","strict_min_version":"140.0"}},"background":{"scripts":["background.js"]},"experiment_apis":{"smoke":{"schema":"schema.json","parent":{"scopes":["addon_parent"],"paths":[["smoke"]],"script":"smoke.js"}}}}
    with zipfile.ZipFile(extensions/'meetling-ci@example.invalid.xpi','w') as z:
        z.writestr('manifest.json',json.dumps(manifest))
        z.writestr('schema.json',json.dumps([{"namespace":"smoke","functions":[{"name":"run","type":"function","async":True,"parameters":[]}]}]))
        z.writestr('background.js','browser.smoke.run();')
        z.writestr('icon.svg',(root/'extension/icon.svg').read_bytes())
        z.writestr('production.js',(root/'extension/experiments/meetlingHost/implementation.js').read_bytes())
        z.writestr('smoke.js',(root/'tests/host-smoke.js').read_bytes())
    prefs={"extensions.autoDisableScopes":0,"extensions.enabledScopes":15,"extensions.startupScanScopes":15,"xpinstall.signatures.required":False,"mail.shell.checkDefaultClient":False,"mail.provider.enabled":False,"mailnews.start_page.enabled":False,"mail.accountmanager.accounts":"account1","mail.accountmanager.defaultaccount":"account1","mail.account.account1.server":"server1","mail.accountmanager.localfoldersserver":"server1","mail.server.server1.type":"none","mail.server.server1.hostname":"Local Folders","mail.server.server1.name":"Local Folders","mail.server.server1.userName":"nobody","calendar.item.editInTab":True,"calendar.timezone.local":"Europe/Berlin","calendar.alarms.show":False,"datareporting.policy.dataSubmissionPolicyBypassNotification":True,"toolkit.telemetry.reportingpolicy.firstRun":False}
    (profile/'user.js').write_text('\n'.join(f'user_pref({json.dumps(k)}, {json.dumps(v)});'for k,v in prefs.items()))
    env=dict(os.environ,MEETLING_SMOKE_RESULT=str(result),MOZ_CRASHREPORTER_DISABLE='1')
    command=[a.binary,'-no-remote','-profile',str(profile)]
    try:
        process=subprocess.run(command,env=env,timeout=150,capture_output=True,text=True)
        print(process.stdout[-12000:]);print(process.stderr[-12000:])
    except subprocess.TimeoutExpired as error:
        print((error.stdout or b'')[-12000:]);print((error.stderr or b'')[-12000:])
        raise SystemExit('Thunderbird smoke timed out')
    if not result.exists(): raise SystemExit('Thunderbird did not write a smoke result')
    report=json.loads(result.read_text());print(json.dumps(report,indent=2))
    if not report.get('success'): raise SystemExit(1)
