"""Build a deterministic, dependency-free Thunderbird XPI and verify its entries."""
import hashlib
import json
from pathlib import Path
import sys
import zipfile
root = Path(__file__).resolve().parents[1]
source = root / 'extension'
manifest = json.loads((source / 'manifest.json').read_text())
out = Path(sys.argv[1]) if len(sys.argv)>1 else root / 'dist'
out.mkdir(parents=True, exist_ok=True)
artifact = out / f'Meetling-Thunderbird-{manifest["version"]}.xpi'
files = sorted(p for p in source.rglob('*') if p.is_file())
with zipfile.ZipFile(artifact, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
    for file in files:
        name = file.relative_to(source).as_posix()
        assert '..' not in name.split('/')
        info = zipfile.ZipInfo(name, (2026,1,1,0,0,0))
        info.compress_type = zipfile.ZIP_DEFLATED
        archive.writestr(info, file.read_bytes())
with zipfile.ZipFile(artifact) as archive:
    assert archive.testzip() is None
    names=set(archive.namelist())
    for name in ['manifest.json','core.js','background.js','panel.html','panel.js','panel.css','icon.svg','experiments/meetlingHost/schema.json','experiments/meetlingHost/implementation.js']:
        assert name in names, name
(out/'SHA256SUMS').write_text(hashlib.sha256(artifact.read_bytes()).hexdigest()+'  '+artifact.name+'\n')
(out/'README.md').write_text((root/'README.md').read_text())
print(artifact)
