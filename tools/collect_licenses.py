"""Collect exact installed-package notices and bundled binary dependency sources.

Only license/notice text and SBOMs are copied; no downloaded code is executed.
Network archives are kept in memory. Failure prevents an apparently complete inventory.
"""
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
import base64
import importlib.metadata as metadata
import json
from pathlib import Path
import sys
import tarfile
import urllib.request

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'release' / 'licenses'


def fetch(url):
    request=urllib.request.Request(url,headers={'User-Agent':'Excel-Privacy-license-inventory'})
    with urllib.request.urlopen(request,timeout=35) as response:return response.read()


def archive_notices(data, destination):
    found=[]
    with tarfile.open(fileobj=BytesIO(data),mode='r:*') as archive:
        for member in archive.getmembers():
            basename=Path(member.name).name
            if member.isfile() and any(basename.upper().startswith(s) for s in ['LICENSE','LICENCE','COPYING','NOTICE','COPYRIGHT']):
                if member.size>500000:continue
                destination.mkdir(parents=True,exist_ok=True)
                content=archive.extractfile(member).read()
                (destination/basename).write_bytes(content);found.append(basename)
    if not found:raise RuntimeError('No license text found for '+destination.name)
    return found


def from_pypi(name, version):
    data=json.loads(fetch(f'https://pypi.org/pypi/{name}/{version}/json'))
    item=next(item for item in data['urls'] if item['packagetype']=='sdist')
    return archive_notices(fetch(item['url']),OUT/name)


def crate_notices(component):
    name, version=component['name'],component['version']
    destination=OUT/'rust'/f'{name}-{version}'
    existing=list(destination.glob('*')) if destination.exists() else []
    if existing:
        return name,[p.name for p in existing]
    return name,archive_notices(fetch(f'https://static.crates.io/crates/{name}/{name}-{version}.crate'),destination)


def main():
    OUT.mkdir(parents=True,exist_ok=True)
    names=['openpyxl','et-xmlfile','cryptography','cffi','pycparser','pyinstaller','pyinstaller-hooks-contrib','altgraph','packaging']
    names+=['macholib'] if sys.platform=='darwin' else ['pefile','pywin32-ctypes']
    inventory=[]
    for name in names:
        dist=metadata.distribution(name);folder=OUT/name;folder.mkdir(exist_ok=True);texts=[]
        for entry in dist.files or []:
            relative=str(entry)
            if '.dist-info/' not in relative:continue
            basename=Path(relative).name
            if any(basename.upper().startswith(p) for p in ['LICENSE','COPYING','NOTICE']):
                (folder/basename).write_bytes(dist.locate_file(entry).read_bytes());texts.append(basename)
            if '/sboms/' in relative and relative.endswith('.json'):
                (folder/basename).write_bytes(dist.locate_file(entry).read_bytes())
        if not texts:
            texts=[p.name for p in folder.iterdir() if p.is_file() and any(
                p.name.upper().startswith(prefix) for prefix in ['LICENSE','LICENCE','COPYING','NOTICE'])]
        if not texts:texts=from_pypi(name,dist.version)
        inventory.append({'name':name,'version':dist.version,
            'license':dist.metadata.get('License-Expression') or dist.metadata.get('License'), 'notices':texts})
    python_license=Path(sys.base_prefix)/'lib'/'python3.12'/'LICENSE.txt'
    if not python_license.exists():python_license=Path(sys.base_prefix)/'LICENSE.txt'
    if not python_license.exists():raise RuntimeError('Python runtime license not found')
    (OUT/'Python-LICENSE.txt').write_bytes(python_license.read_bytes())
    import tkinter
    interpreter=tkinter.Tcl()
    tcl_library=Path(interpreter.eval('info library'))
    candidates=[tcl_library/'license.terms',tcl_library.parent/'license.terms',Path('/opt/homebrew/opt/tcl-tk/license.terms')]
    native=next((p for p in candidates if p.is_file()),None)
    if native:
        (OUT/'Tcl-Tk-license.terms').write_bytes(native.read_bytes())
    elif not (OUT/'Tcl-Tk-license.terms').exists():
        raise RuntimeError('Tcl/Tk license terms missing')
    rust=json.loads((OUT/'cryptography'/'cryptography-rust.cyclonedx.json').read_text())
    crates=[c for c in rust['components'] if not c['name'].startswith('cryptography-') and c.get('scope')!='excluded']
    with ThreadPoolExecutor(max_workers=6) as executor:
        for name, texts in executor.map(crate_notices,crates):print('crate notices',name,len(texts),flush=True)
    # The cryptography wheel includes a static OpenSSL. Include upstream notices
    # from the exact source archive referenced by its own wheel SBOM.
    ssl=json.loads((OUT/'cryptography'/'sbom.json').read_text())['components'][0]
    ssl_out=OUT/'openssl';ssl_out.mkdir(exist_ok=True)
    for filename in ['LICENSE.txt', 'NOTICE']:
        try:
            document=json.loads(fetch(f'https://api.github.com/repos/openssl/openssl/contents/{filename}?ref=openssl-{ssl["version"]}'))
        except urllib.error.HTTPError as error:
            if error.code==404 and filename=='NOTICE':continue
            raise
        (ssl_out/filename).write_bytes(base64.b64decode(document['content']))
    (OUT/'inventory.json').write_text(json.dumps({'python':sys.version.split()[0],
        'platform':sys.platform,'tcl':interpreter.eval('info patchlevel'),
        'packages':inventory,'openssl':ssl['version'],'rust_components':len(crates)},indent=2)+'\n')
    print('License inventory complete',flush=True)


if __name__=='__main__':main()
