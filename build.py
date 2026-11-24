"""Build the FEDORAS reading site from a verified snapshot of the approved share."""
import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import shutil
from urllib.parse import urlencode, urlsplit
from bs4 import BeautifulSoup
import docx
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
import parse
from reader import Reader, cover, esc, space, is_cover_background
from facets import SUBJECT_MAP

ROOT=Path(__file__).resolve().parent
COPY=json.loads((ROOT/'site-copy.json').read_text(encoding='utf-8'))
OFFICIAL_DISCLAIMER='Funded by the European Union. Views and opinions expressed are however those of the author(s) only and do not necessarily reflect those of the European Union or the European Education and Culture Executive Agency (EACEA). Neither the European Union nor EACEA can be held responsible for them.'
PDFS={
 'Introduction to the Materials - Start Here!/FEDORAS Introduction to Materials.pdf': 'fedoras-introduction.pdf',
 'Introduction to the Materials - Start Here!/FEDORAS Materials - Slidebook.pdf': 'fedoras-slidebook.pdf',
}
FACETS={'subject':'Subject','time':'Time needed','country':'Country','pillar':'FEDORAS pillar','type':'Activity type'}


def write(path,text): path.write_text(text,encoding='utf-8')


def credit(prefix,edited=None):
    edited_html=f'<p class="edited">Last edited <time datetime="{edited.isoformat()}">{edited.day} {edited:%B %Y}</time></p>' if edited else ''
    return f'''<footer class="credit" aria-label="European Union funding">{edited_html}<div class="credit-row"><img class="fund-mark" src="{prefix}assets/eu-funding.png" width="4125" height="920" alt="Funded by the European Union"><div class="credit-copy"><p>{esc(COPY['grant'])}</p><p>{esc(OFFICIAL_DISCLAIMER)}</p></div></div></footer>'''


def shell(title,body,prefix='./',index=False,about=False,edited=None):
    links=[('#start','Start here','start'),('#activities','Browse everything','activities'),('about/','About FEDORAS','about')]
    nav=[]
    for href,label,area in links:
        current=' aria-current="page"' if area=='about' and about else ' aria-current="location"' if index and area=='start' else ''
        source=href if index and href.startswith('#') else prefix+href
        nav.append(f'<li><a href="{source}" data-rail="{area}"{current}>{label}</a></li>')
    nav=''.join(nav)
    return f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{esc(title)} · FEDORAS</title><link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&amp;family=Source+Serif+4:ital,opsz,wght@0,8..60,400;0,8..60,600;1,8..60,400&amp;display=swap" rel="stylesheet"><link rel="stylesheet" href="{prefix}site.css"><script src="{prefix}site.js" defer></script></head><body><a class="skip" href="#main">Skip to content</a><div class="shell"><nav class="rail" aria-label="Main"><a class="brand" href="{prefix}"><b>FEDORAS</b><span>Teaching materials</span></a><ul class="nav">{nav}</ul></nav><main id="main"><div class="{'wrap' if index or about else 'act'}">{body}{credit(prefix,edited)}</div></main></div></body></html>'''



def download(href,title,kind,caption='',local=False):
    attr=' download' if local else ''
    subtitle=f'<span class="download-caption">{esc(caption)}</span>' if caption else ''
    icon=f'''<span class="file-mark" aria-hidden="true"><svg viewBox="0 0 40 44" fill="none" focusable="false"><path d="M10 6H29V36H10Z" fill="var(--panel)" stroke="currentColor" stroke-width="1.25"/><path d="M14 9H27L34 16V40H14Z" fill="var(--card)" stroke="currentColor" stroke-width="1.25" stroke-linejoin="round"/><path d="M27 9V16H34" stroke="currentColor" stroke-width="1.25" stroke-linejoin="round"/></svg><span>{kind}</span></span>'''
    action=f'Download {kind}' if local else 'Open in Nextcloud'
    arrow='↓' if local else '→' 
    return f'''<a href="{esc(href)}" class="download-tile"{attr}>{icon}<span class="download-title"><strong>{esc(title)}</strong>{subtitle}</span><span class="download-badge"><span aria-hidden="true">{arrow}</span> {action}</span></a>'''


def about_page():
    pairs=lambda group: ''.join(f'<div><dt>{esc(term)}</dt><dd>{esc(meaning)}</dd></div>' for term,meaning in group.items())
    refs=''.join(f'<li>{esc(t)}</li>' for t in COPY['references'])
    links=[('About the project',COPY['project_url']),('Teaching Resources',COPY['resources_url']), (COPY['source_title'],COPY['source']),('WP2: FEDORAS Materials',COPY['repository_url'])]
    sources=''.join(f'<li><a href="{esc(href)}">{esc(label)}</a></li>' for label,href in links)
    body=f'''<article class="about-project"><h1>About FEDORAS</h1><p class="lede">{esc(COPY['project_intro'])}</p><div class="prose"><p>{esc(COPY['project_resources'])}</p><h2>Concepts and terms</h2><dl class="gl">{pairs(COPY['pillars'])}{pairs(COPY['concepts'])}<div><dt>GreenComp</dt><dd>{esc(COPY['greencomp'])}</dd></div></dl><h2>References and background</h2><ul class="references">{refs}</ul><ul class="docs">{sources}</ul></div></article>'''
    return shell('About FEDORAS',body,'../',about=True)

def find_payload(dom):
    payload=Counter(('link',a['href']) for a in dom.select('a[href]'))
    payload.update(('image',i['src']) for i in dom.select('img[src]'))
    payload[('math','')]=len(dom.find_all('math'))
    return payload


def source_payload(d,r):
    """Check independently against the XML, including hyperlinks in merged worksheets."""
    body=d.element.body
    excluded=set(d.tables[i]._tbl for i in range(4))
    for n in body:
        if n.tag == qn('w:tbl'): break
        if n.tag == qn('w:p') and is_cover_background(Paragraph(n,d),r.imgmap): excluded.add(n)
    keep=lambda x: not any(p in excluded for p in x.iterancestors())
    out=Counter()
    for h in body.iter(qn('w:hyperlink')):
        if not keep(h): continue
        rel=d.part.rels.get(h.get(qn('r:id')))
        if rel is not None and rel.is_external and re.match(r'^(https?|mailto):',rel.target_ref,re.I): out['link',rel.target_ref]+=1
        else: raise ValueError('Unsupported narrative hyperlink')
    for b in body.iter(qn('a:blip')):
        if keep(b): out['image',r.imgmap[b.get(qn('r:embed'))]]+=1
    out['math','']=sum(keep(x) for x in body.iter(qn('m:oMath')))
    return out


def when(s):
    m=re.fullmatch(r'([0-9]+(?:[.,][0-9]+)?)\s*(minutes?|mins?|m|hours?|hrs?|h)',s.strip(),re.I)
    if not m: return []
    mins=float(m[1].replace(',','.'))*(1 if m[2].lower().startswith('m') else 60)
    return ['Up to 1 hour' if mins<=60 else 'Over 1 hour']


def tag(key,label):
    return f'<a class="chip" href="../?{esc(urlencode({key:label}))}#activities">{esc(label)}</a>'


def activity(d,a,rr,url):
    title,sub,typ,origin,edited=cover(d)
    meta=a['raw_meta']
    subjects=set()
    for raw in a['disciplines_list']:
        for s in re.split(r',\s*(?![^()]*\))',raw):
            if s.strip(): subjects.update(SUBJECT_MAP.get(s.strip().casefold(),[s.strip()]))
    facets=dict(subject=sorted(subjects),time=when(a['duration']),country=a['hubs'],pillar=a['pillars'],type=[typ.removeprefix('FEDORAS ')])
    sections=rr.sections(d)
    if not sections or sections[0][0]!='The activity': raise ValueError('Missing lesson: '+title)
    details=[('Written by',a['authors']),('Hub',', '.join(a['hubs'])),('Designed at',a['designed_at']),('Tried out at',a['implemented_at']),('Disciplines',', '.join(a['disciplines_list'])),('STEAM topics',a['steam']),('Keywords',a['keywords']),('FEDORAS pillars',', '.join(a['pillars'])),('GreenComp areas',', '.join(a['greencomp_raw']))]
    # Keep any author-provided metadata not in the main template field list.
    details += [(label,val) for label,val in meta.items() if label.casefold() not in parse.FIELD_OF and not label.casefold().startswith('steam topics')]
    if origin: details.append(('Development',origin))
    dl=''.join(f'<div><dt>{esc(k)}</dt><dd>{esc(v).replace(chr(10),"<br>")}</dd></div>' for k,v in details if v.strip() and v.strip() not in ('-','–','—'))
    if dl: sections.append(('About this activity',f'<dl class="about">{dl}</dl>'))
    ids=[]
    body=[]
    for k,v in sections:
        ident=re.sub(r'[^a-z0-9]+','-',k.lower()).strip('-')
        if ident in ids: ident+=f'-{len(ids)+1}'
        ids.append(ident)
        body.append(f'<section aria-labelledby="{ident}"><h2 id="{ident}">{esc(k)}</h2>{v}</section>')
    body='\n'.join(body)
    dom=BeautifulSoup(body,'html.parser')
    if find_payload(dom)!=source_payload(d,rr):
        raise ValueError(f'Links, images or equations differ for {title}: expected {source_payload(d,rr)-find_payload(dom)} extra {find_payload(dom)-source_payload(d,rr)}')
    toc=''.join(f'<a href="#{i}">{esc(k)}</a>' for i,(k,_) in zip(ids,sections))
    chips=''.join(tag(key,v) for key in FACETS for v in facets[key])
    facts=' · '.join(space(v) for v in (a['duration'],a['target'],', '.join(a['hubs'])) if space(v) and space(v) not in ('-','–','—'))
    kicker=esc(typ.removeprefix('FEDORAS '))
    description=a['description']
    original=download(COPY['repository_url']+'?dir=/&openfile=true','Original Word document','DOCX', a['file'])
    html=f'''<article><a class="back" href="../#activities" data-back>← All activities</a><p class="kicker">{kicker}</p><h1>{esc(title)}</h1>{f'<p class="subtitle">{esc(sub)}</p>' if sub else ''}{f'<p class="facts">{esc(facts)}</p>' if facts else ''}{f'<p class="lede">{esc(description)}</p>' if description.strip() else ''}<div class="chips tags">{chips}</div><div class="orig">{original}<p class="download-help">In Nextcloud, use the file’s <b>Actions</b> menu (⋯), then choose <b>Download</b>.</p></div><nav class="toc" aria-label="On this page"><span>On this page</span>{toc}</nav><div class="reading">{body.replace('src="assets/img/','src="../assets/img/')}</div></article>'''
    card=dict(id=a['slug'],title=title,subtitle=sub,summary=space(a['description']),facets=facets,facts=facts)
    search=space(f'{title} {sub} {facts} {a["description"]} '+dom.get_text(' ',strip=True))
    report=dict(file=a['file'],id=a['slug'],title=title,subtitle=sub,date=str(edited),sections=list(zip(ids,[k for k,_ in sections])),
                links=sum(v for (k,_),v in rr.expected.items() if k=='link'),images=sum(v for (k,_),v in rr.expected.items() if k=='image'),equations=rr.expected['math',''],
                omitted_placeholders=dict(rr.exclusions),source_metadata_warnings=a['warnings'])
    return card,shell(title,html,'../',edited=edited),search,report


def home(cards):
    filters=[]
    for key,label in FACETS.items():
        vals=sorted({v for a in cards for v in a['facets'][key]}, key=str.casefold)
        if vals: filters.append(f'<div class="facet" role="group" aria-label="{label}"><span>{label}</span>'+''.join(f'<button class="chip" type="button" data-facet="{key}" value="{esc(v)}" aria-pressed="false">{esc(v)}</button>' for v in vals)+'</div>')
    rows=[]
    for a in sorted(cards,key=lambda c:c['title'].casefold()):
        summary=a['summary']; limit=240
        if len(summary)>limit: summary=summary[:limit].rsplit(' ',1)[0]+'…'
        attr=' '.join(f'data-{k}="{esc(json.dumps(v,ensure_ascii=False))}"' for k,v in a['facets'].items())
        rows.append(f'''<li data-activity="{a['id']}" {attr}><a class="card" href="./{a['id']}/"><span class="ct">{esc(a['title'])}</span>{f'<span class="cs">{esc(a["subtitle"])}</span>' if a['subtitle'] else ''}{f'<p class="co">{esc(summary)}</p>' if summary else ''}<span class="card-meta">{esc(a['facts'])}</span></a></li>''')
    intro=download('./documents/fedoras-introduction.pdf','FEDORAS Introduction to Materials','PDF','',local=True)
    slide=download('./documents/fedoras-slidebook.pdf',COPY['slidebook'],'PDF','',local=True)
    start=f'''<section id="start"><h1>Teaching materials</h1><p class="lede">{esc(COPY['overview'])}</p><div class="downloads">{intro}{slide}</div></section>'''
    browse=f'''<section id="activities"><h2>Browse everything</h2><form class="facets" hidden><label for="query">Search activities</label><input id="query" class="search" type="search" name="q" placeholder="Search titles and activity text" autocomplete="off">{''.join(filters)}<button type="reset" class="clear">Clear filters</button></form><p id="count" class="count" role="status">{len(cards)} activities</p><ul class="cards" id="cards">{''.join(rows)}</ul></section>'''
    return shell('Teaching materials',start+browse,index=True)


def build(snapshot,output,report,baseline):
    inputs=json.loads((snapshot/'inventory.json').read_text(encoding='utf-8'))
    found={e['path']:e for e in inputs}
    if len(found)!=len(inputs) or not inputs: raise ValueError('Empty or duplicate inventory')
    if baseline:
        missing=set(json.loads(baseline.read_text(encoding='utf-8')))-set(found)
        if missing: raise ValueError('Maintainer review required for removed or renamed files: '+', '.join(sorted(missing)))
    output.mkdir(parents=True,exist_ok=False)
    parse.IMG=str(output/'assets'/'img'); Path(parse.IMG).mkdir(parents=True)
    for f in ('site.css','site.js','_headers','assets/eu-funding.png'): shutil.copyfile(ROOT/f,output/f)
    cards=[]; search={}; reports=[]; used=set()
    share=urlsplit(COPY['repository_url']).path.rstrip('/').rsplit('/',1)[-1]
    allowed_prefix=f'https://labs.fedoras-academy.eu/public.php/dav/files/{share}/'
    for e in inputs:
        fn=e['path']; p=Path(fn); src=snapshot/'raw'/p; url=e['url']
        if p.is_absolute() or '..' in p.parts or not url.startswith(allowed_prefix):
            raise ValueError('Unexpected source path or URL')
        if hashlib.sha256(src.read_bytes()).hexdigest()!=e['sha256']: raise ValueError('Input checksum mismatch: '+fn)
        if '/' in fn or p.suffix.lower()!='.docx': continue
        a=parse.parse_activity(str(src)); slug=a['slug']
        if slug in search or not slug: raise ValueError('Duplicate or empty activity ID: '+fn)
        d=docx.Document(src); r=Reader(d,slug)
        card,page,txt,verification=activity(d,a,r,url)
        folder=output/slug; folder.mkdir(); write(folder/'index.html',page)
        cards.append(card); search[slug]=txt; reports.append(verification)
        used.update(k[1] for k in r.expected if k[0]=='image')
    if not cards: raise ValueError('No activity documents in build')
    for asset in Path(parse.IMG).iterdir():
        if 'assets/img/'+asset.name not in used: asset.unlink()
    (output/'documents').mkdir()
    for name,alias in PDFS.items():
        if name not in found: raise ValueError('Missing featured PDF: '+name)
        src=snapshot/'raw'/name
        with src.open('rb') as f:
            if f.read(5)!=b'%PDF-': raise ValueError('Expected PDF: '+name)
        shutil.copyfile(src,output/'documents'/alias)
    write(output/'index.html',home(cards))
    (output/'about').mkdir(); write(output/'about'/'index.html',about_page())
    write(output/'search-index.json',json.dumps(search,ensure_ascii=False,separators=(',',':')))
    write(report,json.dumps(reports,ensure_ascii=False,indent=2)+'\n')
    print(f'{len(cards)} activity pages, {sum(x["links"] for x in reports)} links, {sum(x["images"] for x in reports)} body images, {sum(x["equations"] for x in reports)} equations.')


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('snapshot',type=Path);p.add_argument('output',type=Path);p.add_argument('report',type=Path);p.add_argument('--baseline',type=Path,default=ROOT/'approved-file-list.json'); a=p.parse_args()
    build(a.snapshot,a.output,a.report,a.baseline)
