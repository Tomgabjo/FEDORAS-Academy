"""Turn the accepted WP2 Word templates into a continuous reading page."""
from collections import Counter
from datetime import date
import html
import re
from bs4 import BeautifulSoup
import docx
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
import parse

esc = html.escape
space = lambda s: ' '.join(s.split())
SEPARATOR = re.compile(r'^[\s\-–—●•]+$')
COVER_BACKGROUND = 'assets/img/5e4c5a3f2f11.png'  # shared, non-instructional Word cover template
CAP = re.compile(r'^(?:fig(?:ure)?[.\s]+\d|(?:εικ(?:όνα)?)[.\s]+\d)', re.I)
PROMPTS = re.compile(r'^(?:\(?provide a description of the materials needed\)?[.:]?|please provide any worksheets or other kind of supporting material developed[.:]?|(?:please,? describe )?how the activity contributes to pursuing (?:embodying sustainability values|embracing complexity|envisioning sustainab(?:ility|le) futures|acting for sustainability)(?: (?:and )?objectives)?[.:]?|how does this reconsider and reformulate the aims of science and sustainability education[?]?|other[:. ]*)$', re.I)


def unique_cells(row): return list({c._tc: c for c in row.cells}.values())


def is_cover_background(p,imgmap):
    blips=list(p._p.iter(qn('a:blip')))
    return not p.text.strip() and len(blips)==1 and imgmap.get(blips[0].get(qn('r:embed')))==COVER_BACKGROUND and not p._p.findall('.//'+qn('w:hyperlink'))


def cover(d):
    first = d.tables[0]
    ps = first.cell(0,0).paragraphs
    lines = [line.strip() for p in ps for line in p.text.splitlines() if line.strip()]
    title, subtitle = lines[0], ' '.join(lines[1:])
    if not subtitle:
        runs = [(r.font.size.pt if r.font.size else 30, r.text.strip()) for p in ps for r in p.runs if r.text.strip()]
        if runs and max(s for s,_ in runs) >= 30 and any(s <= 20 for s,_ in runs) and runs[0][0] >= 30:
            cut = next(i for i,(s,_) in enumerate(runs) if s<=20)
            title, subtitle = ' '.join(t for _,t in runs[:cut]), ' '.join(t for _,t in runs[cut:])
    if title.lower() == 'laboratory activity:' and subtitle:
        title, subtitle = title + ' ' + subtitle, ''
    typ, *origin = [line.strip() for line in first.cell(1,0).text.splitlines() if line.strip()]
    match = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{4})',first.cell(2,0).text)
    edited = date(int(match[3]),int(match[2]),int(match[1])) if match else None
    return title, subtitle, typ, ' '.join(origin), edited


def list_info(p):
    props = p._p.pPr
    np = props.numPr if props is not None else None
    if np is None or np.numId is None:
        style = p.style
        while style is not None and (np is None or np.numId is None):
            pr = style.element.pPr
            np = pr.numPr if pr is not None else None
            style = style.base_style
    if np is None or np.numId is None or not np.numId.val: return None
    numid = np.numId.val
    level = np.ilvl.val if np.ilvl is not None else 0
    root = p.part.numbering_part.element
    n = root.find(f'./{qn("w:num")}[@{qn("w:numId")}="{numid}"]')
    if n is None: raise ValueError('Unresolved list numbering')
    aid = n.find(qn('w:abstractNumId')).get(qn('w:val'))
    a = root.find(f'./{qn("w:abstractNum")}[@{qn("w:abstractNumId")}="{aid}"]')
    lvl = a.find(f'./{qn("w:lvl")}[@{qn("w:ilvl")}="{level}"]')
    ov = n.find(f'./{qn("w:lvlOverride")}[@{qn("w:ilvl")}="{level}"]')
    if ov is not None and ov.find(qn('w:lvl')) is not None: lvl=ov.find(qn('w:lvl'))
    def val(parent,key,default):
        x = parent.find(qn('w:'+key)) if parent is not None else None
        return x.get(qn('w:val'),default) if x is not None else default
    fmt = val(lvl,'numFmt','bullet')
    start = int(val(ov,'startOverride',val(lvl,'start','1')))
    return numid,level,fmt,start


class Reader:
    def __init__(self,d,slug):
        self.imgmap = parse.save_images(d,slug)
        self.expected = Counter()
        self.exclusions = Counter()
        self.labels=[]

    def paragraph(self,p,upright=False):
        h = parse.para_html(p,self.imgmap,upright)
        if not h: return '', BeautifulSoup('', 'html.parser')
        soup = BeautifulSoup(h,'html.parser')
        return soup.contents[0].decode_contents(),soup

    def flow(self, paragraphs, *, headings=True):
        paragraphs = list(paragraphs)
        runs = [r for p in paragraphs for r in p.runs if r.text.strip()]
        # The pilot's upright reading treatment also applies to a fully italic
        # paragraph in a mixed section; isolated emphasis stays in the HTML.
        upright = bool(runs) and sum(len(r.text) for r in runs if r.italic) > .8*sum(len(r.text) for r in runs)
        out, lists, count = [], [], Counter()
        pending = None
        depth = 3
        def close(keep=0):
            while len(lists)>keep:
                _,_,_,tag = lists.pop()
                out.append(f'</li></{tag}>')
        def flush_caption():
            nonlocal pending
            if pending is not None: out.append(f'<p>{pending}</p>'); pending=None
        for p in paragraphs:
            raw = p.text
            lines = raw.splitlines()
            t = space(raw)
            rich = any(x.tag in (qn('w:hyperlink'),qn('w:drawing'),qn('m:oMath'),qn('m:oMathPara')) for x in p._p.iter())
            if not rich and (not t or SEPARATOR.fullmatch(t) or PROMPTS.fullmatch(t)):
                if t: self.exclusions[t]+=1
                continue
            r = [r for r in p.runs if r.text.strip()]
            inner,soup = self.paragraph(p, upright or bool(r and all(rn.italic for rn in r)))
            if not space(soup.get_text()) and not soup.find(['img','a','math']): continue
            if len(lines)>1 and PROMPTS.fullmatch(space(lines[0])):
                prefix, sep, remainder = inner.partition('<br/>')
                prefix_dom = BeautifulSoup(prefix,'html.parser')
                if sep and not prefix_dom.find(['img','a','math']) and space(prefix_dom.get_text())==space(lines[0]):
                    self.exclusions[space(lines[0])]+=1
                    inner=re.sub(r'^(?:\s*<br/>)+', '', remainder)
                    t=space(' '.join(lines[1:]))
            for x in soup.select('a[href]'): self.expected[('link',x['href'])]+=1
            for x in soup.select('img[src]'): self.expected[('image',x['src'])]+=1
            self.expected[('math','')]+=len(soup.find_all('math'))
            imgs=soup.find_all('img')
            text=space(soup.get_text())
            if imgs and not text and not soup.find('a'):
                close()
                cap=f'<figcaption>{pending}</figcaption>' if pending else ''
                out.append('<figure>'+inner+cap+'</figure>');pending=None;continue
            flush_caption()
            if CAP.match(t) and not imgs:
                close(); pending=inner;continue
            n = list_info(p)
            if not n and re.match(r'^[●•]\s+',p.text.strip()): n=(-1,0,'bullet',1)
            if n:
                numid,level,fmt,start=n
                tag = 'ul' if fmt in ('bullet','none') else 'ol'
                typ = {'lowerLetter':'a','upperLetter':'A','lowerRoman':'i','upperRoman':'I'}.get(fmt)
                if numid==-1:
                    inner=re.sub(r'^((?:<(?:em|strong|u)>)*)(?:[●•])[\s\t]*',r'\1',inner)
                while lists and (lists[-1][1]>level or (lists[-1][1]==level and (lists[-1][0],lists[-1][2])!=(numid,fmt))): close(len(lists)-1)
                if lists and lists[-1][1]==level:
                    out.append('</li><li>'+inner)
                else:
                    number = start + count[numid,level]
                    attr=(f' start="{number}"' if tag=='ol' and number!=1 else '')+(f' type="{typ}"' if typ else '')
                    out.append(f'<{tag}{attr}><li>'+inner); lists.append((numid,level,fmt,tag))
                count[numid,level]+=1
                for k in list(count):
                    if k[0]==numid and k[1]>level: count[k]=0
            else:
                close()
                styled = (p.style.name or '').lower().startswith('heading')
                if headings and not rich and t and (styled or (len(t)<90 and r and all(x.bold for x in r))):
                    if t.lower().startswith('realisation'): depth=4; level=3
                    elif t.rstrip(':').lower() in ('material','materials','procedure'): level=depth
                    else: depth=3; level=3
                    out.append(f'<h{level}>{inner}</h{level}>')
                else: out.append(f'<p>{inner}</p>')
        close(); flush_caption()
        return '\n'.join(out)

    def labelled(self,table,assessment=False):
        out=[]
        for i,row in enumerate(table.rows):
            cells=unique_cells(row)
            if i==0:
                # Header is the first paragraph only; any subsequent content is real.
                out.append(self.flow(cells[0].paragraphs[1:]));continue
            if len(cells)==1:
                out.append(self.flow(cells[0].paragraphs));continue
            if len(cells)!=2: raise ValueError('Unexpected label/value table structure')
            label=space(cells[0].text)
            p=cells[1].paragraphs
            lower=label.casefold()
            if not assessment and ('objectives covered' in lower):
                label=next((val for key,val in [('knowledge','Knowledge'),('skills','Skills'),('attitudinal','Attitudes')] if key in lower),label)
                # In some documents each selected objective is a plain paragraph.
                if not any(list_info(x) or re.match(r'^[●•]\s*',x.text.strip()) for x in p):
                    frag=self.flow(p,headings=False)
                    dom = BeautifulSoup(frag,'html.parser')
                    if dom.contents and all(x.name=='p' for x in dom.find_all(recursive=False)):
                        frag='<ul>'+''.join(f'<li>{x.decode_contents()}</li>' for x in dom.find_all('p',recursive=False))+'</ul>'
                else: frag=self.flow(p,headings=False)
            else:
                frag=self.flow(p,headings=False)
                if assessment:
                    label=re.sub(r'^(?:Please describe )?how (?:the |a )?-?\s*','',label,flags=re.I)
                    label=re.sub(r'\s+(?:learning )?outcomes are assessed[.:]?$','',label,flags=re.I)
            if frag:
                show_label=bool(label) and lower!='comments'
                level=3 if assessment else 4
                out.append((f'<h{level}>{esc(label)}</h{level}>' if show_label else '')+frag)
        return ''.join(out)

    def sections(self,d):
        lesson, works, objectives, assessment, other = [],[],[],[],[]
        refs, loose = [],[]
        seen_tables=0
        for b in parse.iter_block_items(d):
            if isinstance(b,Paragraph):
                if b.text.strip().casefold()=='references': refs.append([])
                elif refs: refs[-1].append(b)
                elif seen_tables == 0 and is_cover_background(b, self.imgmap): continue
                else: loose.append(b)
                continue
            seen_tables+=1
            if seen_tables<=4: continue
            first=b.cell(0,0).paragraphs[0]
            head=space(first.text); key=head.casefold()
            if key.endswith('objectives'):
                frag=self.labelled(b)
                if frag: objectives.append('<h3>'+esc(head.removesuffix(' Objectives'))+'</h3>'+frag)
                continue
            if key.startswith('assessment of learning outcomes'):
                frag=self.labelled(b,True)
                if frag: assessment.append(frag)
                continue
            # Worksheets are often a single merged row: do not drop content after the heading.
            pars=[]; seen=set()
            for ri,row in enumerate(b.rows):
                for ci,cell in enumerate(unique_cells(row)):
                    if cell._tc in seen: continue
                    seen.add(cell._tc)
                    pars.extend(cell.paragraphs[1 if ri==ci==0 else 0:])
            frag=self.flow(pars)
            if not frag: continue
            if key.startswith('detailed description'): lesson.append(frag)
            elif key=='worksheets': works.append(frag)
            else: other.append((head,frag))
        pieces=[(label,'\n'.join(chunks)) for label,chunks in [('The activity',lesson),('Worksheets',works),('Learning objectives',objectives),('Assessment',assessment)] if chunks]
        pieces.extend(other)
        if refs:
            frag=self.flow([p for group in refs for p in group],headings=False)
            if frag: pieces.append(('References',frag))
        if frag:=self.flow(loose,headings=False): pieces.append(('Additional material',frag))
        return pieces
