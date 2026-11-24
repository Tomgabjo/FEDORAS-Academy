"""Read a publish-approved Nextcloud share; never write to Nextcloud."""
import base64, concurrent.futures, hashlib, json, os, sys
from pathlib import Path
from urllib.parse import urlsplit, unquote, quote
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET

HOST = 'https://labs.fedoras-academy.eu'
SHARE = os.environ.get('FEDORAS_APPROVED_SHARE', 'yqWgxFNtA2QEFr6')
NS = {'d': 'DAV:'}
DAV = HOST + '/public.php/webdav/'

def open_url(url, method='GET', headers=None):
    result = urlopen(Request(url, method=method, headers=headers or {}), timeout=45)
    if not result.geturl().startswith(HOST + '/public.php/'):
        result.close(); raise ValueError('Unexpected source redirect')
    return result

def listing(folder='', seen=None):
    if seen is None: seen = set()
    if folder in seen or len(seen) > 50: raise ValueError('Unexpected source folder tree')
    seen.add(folder)
    prefix = '/public.php/webdav/' + folder
    auth = base64.b64encode((SHARE + ':').encode()).decode()
    with open_url(DAV + quote(folder), 'PROPFIND', {'Depth':'1', 'Authorization':'Basic ' + auth}) as res:
        if res.status != 207: raise ValueError('Expected WebDAV listing')
        xml = ET.fromstring(res.read())
    files = {}
    for row in xml.findall('d:response', NS):
        href = unquote(urlsplit(row.findtext('d:href', namespaces=NS)).path)
        if href.rstrip('/') == prefix.rstrip('/'): continue
        if not href.startswith(prefix): raise ValueError('Unexpected source path')
        name = href[len(prefix):].rstrip('/')
        if not name or name in ('.','..') or '/' in name or '\\' in name: raise ValueError('Unsafe source path')
        props = next((p.find('d:prop', NS) for p in row.findall('d:propstat', NS) if ' 200 ' in p.findtext('d:status', '', NS)), None)
        if props is None: raise ValueError('Missing source listing properties')
        path = folder + name
        if props.find('d:resourcetype/d:collection', NS) is not None:
            for key, val in listing(path + '/', seen).items():
                if key in files: raise ValueError('Duplicate source listing path')
                files[key] = val
        else:
            if path in files: raise ValueError('Duplicate source listing path')
            files[path] = {'size':int(props.findtext('d:getcontentlength', namespaces=NS)), 'etag':props.findtext('d:getetag', namespaces=NS)}
    return files

def snapshot(destination):
    if not SHARE or any(c not in 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-' for c in SHARE):
        raise SystemExit('Set FEDORAS_APPROVED_SHARE to the approved read-only share identifier.')
    before = listing()
    if not any('/' not in k and k.lower().endswith('.docx') for k in before): raise ValueError('No activity documents in source')
    raw = destination / 'raw'; raw.mkdir(parents=True, exist_ok=False)
    def fetch(entry):
        name, file = entry
        if file['size'] < 1 or file['size'] > 250_000_000: raise ValueError('Invalid source size')
        href = HOST + '/public.php/dav/files/' + quote(SHARE) + '/' + quote(name)
        with open_url(href) as res:
            if res.status != 200: raise ValueError('Source download failed')
            data = res.read(file['size'] + 1)
        if len(data) != file['size']: raise ValueError('Source download size mismatch: ' + name)
        path = raw / name; path.parent.mkdir(exist_ok=True, parents=True); path.write_bytes(data)
        return dict(file, path=name, sha256=hashlib.sha256(data).hexdigest(), url=href)
    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as pool:
        inventory = list(pool.map(fetch, sorted(before.items())))
    if before != listing(): raise ValueError('Nextcloud changed during snapshot; retry')
    (destination/'inventory.json').write_text(json.dumps(inventory, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Snapshot: {len(inventory)} files; {sum(f["size"] for f in inventory):,} bytes')

if __name__ == '__main__': snapshot(Path(sys.argv[1]).absolute())
