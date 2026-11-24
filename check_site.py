"""Check a generated site against its Word inputs and exercise it in Chromium."""
import hashlib
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from functools import partial
from collections import Counter
from pathlib import Path
from threading import Thread
from urllib.parse import urlsplit, unquote
import json
import sys
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright, expect
import docx
from docx.oxml.ns import qn

site, snapshot, report = map(Path, sys.argv[1:4]); site=site.absolute()
results=json.loads(report.read_text()); inventory=json.loads((snapshot/'inventory.json').read_text())
by_file={e['path']:e for e in inventory}
ids=[r['id'] for r in results]
pages=[site/'index.html',site/'about'/'index.html']+[site/i/'index.html' for i in ids]
assert sorted(p.parent.name for p in site.glob('*/index.html'))==sorted(ids+['about'])
assert len(ids)==len({i for i in ids})

def dom(p): return BeautifulSoup(p.read_text(encoding='utf-8'),'html.parser')
for p in pages:
    d=dom(p)
    assert len(d.select('h1'))==len(d.select('footer.credit'))==1
    assert d.select_one('main > div').find_all(recursive=False)[-1].name=='footer'
    assert len(d.select('img.fund-mark'))==1 and 'Grant Agreement n° 101196137.' in d.select_one('footer').text
    assert 'Background reading' not in d.get_text() and 'To be reviewed by WP2' not in d.get_text()
    assert not d.find('details') and not d.find('table')
    assert len(d.select('[id]'))==len({n['id'] for n in d.select('[id]')})
    for n in d.select('img[src], a[href], script[src], link[href]'):
        value=n.get('src',n.get('href'))
        u=urlsplit(value)
        if u.netloc or u.scheme: continue
        dest=(p.parent/unquote(u.path)) if u.path else p
        if dest.is_dir(): dest/= 'index.html'
        assert dest.is_file(), (p,value)
        if u.fragment: assert dom(dest).find(id=u.fragment), (p,value)
for r in results:
    d=dom(site/r['id']/'index.html')
    source=snapshot/'raw'/r['file']; doc=docx.Document(source)
    assert len(d.select('.reading math'))==sum(1 for _ in doc.element.body.iter(qn('m:oMath')))
    assert d.select_one('.orig .download-tile')['href']=='https://labs.fedoras-academy.eu/index.php/s/yqWgxFNtA2QEFr6?dir=/&openfile=true'
    assert d.select_one('.orig .download-caption').get_text()==r['file']
    assert d.select_one('.orig .download-help').get_text().endswith('choose Download.')
    assert d.select_one('.orig .file-mark').get_text(strip=True)=='DOCX'
    assert d.select_one('time')['datetime']==r['date']
    assert d.select_one('.reading section h2').get_text()=='The activity'
    assert not any(len(x)>140 for x in r['omitted_placeholders'])
    assert all(img['src']!='../assets/img/5e4c5a3f2f11.png' for img in d.select('.reading img'))
comics='comic-creation-species-climate-adaptation'
comic=dom(site/comics/'index.html'); assert comic.h1.text=='From dystopia to desirable futures'
assert comic.select_one('.subtitle').text=='Creating comics to learn about threatened species and their adaptations to climate zones'
assert len(comic.select('.reading ol > li'))==7 and comic.find('figcaption')
assert 'At the knowledge level, the activity familiarizes students' in dom(site/'building-fyoutures-through-scenario-making'/'index.html').get_text(' ')
assert 'CASE 1' in dom(site/'osn-discussing-tensions-in-socioscientific-teaching'/'index.html').find(id='worksheets').parent.get_text()
assert 'Komorek' in dom(site/'chaotic-pendulum'/'index.html').find(id='references').parent.get_text()
for asset, basename in [('fedoras-introduction.pdf','FEDORAS Introduction to Materials.pdf'),('fedoras-slidebook.pdf','FEDORAS Materials - Slidebook.pdf')]:
    assert hashlib.sha256((site/'documents'/asset).read_bytes()).hexdigest()==next(e['sha256'] for e in inventory if Path(e['path']).name==basename)

class Quiet(SimpleHTTPRequestHandler):
    def log_message(self,*args): pass
server=ThreadingHTTPServer(('127.0.0.1',0),partial(Quiet,directory=str(site)))
Thread(target=server.serve_forever,daemon=True).start()
base='http://127.0.0.1:'+str(server.server_port)
errors=[]; layout_issues=[]
with sync_playwright() as pw:
    browser=pw.chromium.launch(args=['--no-sandbox'])
    for w,h in [(390,844),(1280,900)]:
        context=browser.new_context(viewport={'width':w,'height':h},is_mobile=w==390,has_touch=w==390,accept_downloads=True)
        page=context.new_page(); page.on('pageerror',lambda exc:errors.append(str(exc)))
        # Rendering tests exercise the self-contained site; remote fonts have fallback stacks.
        page.route('https://fonts.**/*',lambda route:route.abort())
        requested=[];page.on('request',lambda req:requested.append(req.url))
        for p in pages:
            path='/' if p.parent==site else '/'+p.parent.name+'/'
            assert page.goto(base+path,wait_until='load').ok
            overflow=page.evaluate('document.documentElement.scrollWidth - document.documentElement.clientWidth')
            if overflow>1: layout_issues.append((w,path,overflow))
            assert page.locator('.fund-mark').evaluate('(el) => el.complete && el.naturalWidth > 0 && el.getBoundingClientRect().height >= 45')
        assert not any('search-index' in x for x in requested), requested
        page.goto(base+'/')
        expect(page.locator('[data-rail="start"]')).to_have_attribute('aria-current','location')
        page.locator('[data-rail="activities"]').click()
        expect(page.locator('[data-rail="activities"]')).to_have_attribute('aria-current','location')
        expect(page.locator('[data-rail="start"]')).not_to_have_attribute('aria-current','location')
        page.locator('[data-rail="start"]').click()
        expect(page.locator('[data-rail="start"]')).to_have_attribute('aria-current','location')
        page.locator('[data-rail="about"]').click()
        expect(page.locator('[data-rail="about"]')).to_have_attribute('aria-current','page')
        expect(page.locator('h1')).to_have_text('About FEDORAS')
        assert page.locator('.credit-row').evaluate("e=>getComputedStyle(e).flexDirection")==('column' if w<560 else 'row')
        assert page.locator('.credit').evaluate("e=>Math.abs(e.clientWidth-e.parentElement.clientWidth)<1")
        page.goto(base+'/'+comics+'/'); page.locator('.tags a',has_text='Biology').click()
        expect(page.get_by_role('button',name='Biology',exact=True)).to_have_attribute('aria-pressed','true')
        page.locator('#query').fill('storyboard'); expect(page.locator('#count')).to_have_text('1 activity')
        page.get_by_role('link',name='From dystopia to desirable futures',exact=False).first.click()
        assert comics in page.url
        page.locator('.toc a[href="#learning-objectives"]').click(); assert page.url.endswith('#learning-objectives')
        page.locator('[data-back]').click(); expect(page.locator('#query')).to_have_value('storyboard')
        expect(page.locator('#count')).to_have_text('1 activity')
        page.get_by_role('button',name='Clear filters').click();expect(page.locator('#count')).to_have_text(f'{len(ids)} activities')
        for n in ['fedoras-introduction.pdf','fedoras-slidebook.pdf']:
            with page.expect_download() as dl: page.locator(f'#start a[href$="{n}"]').click()
            assert dl.value.suggested_filename==n and dl.value.failure() is None
        if len(sys.argv)>4:
            page.goto(base+'/',wait_until='load'); page.screenshot(path=str(Path(sys.argv[4])/f'home-{w}.png'))
            page.goto(base+'/'+comics+'/',wait_until='load'); page.screenshot(path=str(Path(sys.argv[4])/f'comics-{w}.png'))
        context.close()
assert not layout_issues, layout_issues
assert not errors, errors
print(f'{len(pages)} pages checked at 390 and 1280 CSS pixels; tag, search, return, jumps, and both PDF downloads passed.')
