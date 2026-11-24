#!/usr/bin/env python3
"""Parse FEDORAS WP2 materials into a structured JSON index + HTML body fragments + images."""
import os, re, hashlib, html
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.oxml.ns import qn

RAW = 'raw'
OUT = 'build'
IMG = os.path.join(OUT, 'assets', 'img')


def para_selected(p):
    runs = [r for r in p.runs if r.text.strip()]
    return bool(runs) and all(bool(r.bold) for r in runs)

def cell_list(cell):
    """Return [(text, selected)] for each non-empty paragraph in a cell."""
    out = []
    for p in cell.paragraphs:
        t = p.text.strip()
        if t:
            out.append((t, para_selected(p)))
    return out

def cell_text(cell):
    return '\n'.join(p.text.strip() for p in cell.paragraphs if p.text.strip())

def iter_block_items(parent):
    body = parent.element.body if hasattr(parent, 'element') else parent._tc
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)

def save_images(doc, slug):
    """Map rel-id -> saved web path."""
    m = {}
    os.makedirs(IMG, exist_ok=True)
    for rid, rel in doc.part.rels.items():
        if "image" in rel.reltype:
            blob = rel.target_part.blob
            ext = os.path.splitext(rel.target_part.partname)[1] or '.png'
            h = hashlib.sha1(blob).hexdigest()[:12]
            name = f"{h}{ext}"
            path = os.path.join(IMG, name)
            if not os.path.exists(path):
                with open(path, 'wb') as f:
                    f.write(blob)
            m[rid] = f"assets/img/{name}"
    return m

def para_html(p, imgmap, upright=False):
    """Render a paragraph to HTML, including inline images."""
    parts = []
    def render_run(element):
        run = Run(element, p)
        rendered = []
        for blip in run.element.findall('.//' + qn('a:blip')):
            rid = blip.get(qn('r:embed'))
            if rid in imgmap:
                drawing = next((node for node in blip.iterancestors() if node.tag == qn('w:drawing')), None)
                description = ''
                if drawing is not None:
                    props = drawing.find('.//' + qn('wp:docPr'))
                    if props is not None: description = props.get('descr') or props.get('title') or ''
                if not description and imgmap[rid].endswith('/aa16ab3349fa.png'):
                    description = 'Didactical model for comic creation about species and their adaptation to a changing climate: a spiral moving between story development and investigation'
                rendered.append(f'<img src="{imgmap[rid]}" alt="{html.escape(description, quote=True)}" loading="lazy">')
        t = html.escape(run.text).replace("\n", "<br>")
        if run.underline and t.strip():
            t = f'<u>{t}</u>'
        if t:
            if run.bold:
                t = f'<strong>{t}</strong>'
            if run.italic and not upright:
                t = f'<em>{t}</em>'
            rendered.append(t)
        return ''.join(rendered)

    # Runs and hyperlinks are siblings: their sequence is the reading order.
    for child in p._p:
        if child.tag == qn('w:r'):
            parts.append(render_run(child))
        elif child.tag in (qn('m:oMath'), qn('m:oMathPara')):
            from equations import math_html
            parts.append(math_html(child))
        elif child.tag == qn('w:hyperlink'):
            inner = ''.join(render_run(r) for r in child.findall(qn('w:r')))
            rid = child.get(qn('r:id'))
            rel = p.part.rels.get(rid) if rid else None
            href = rel.target_ref if rel is not None and rel.is_external else ''
            if href and href.split(':', 1)[0].lower() in ('https', 'http', 'mailto'):
                parts.append(f'<a href="{html.escape(href, quote=True)}" target="_blank" rel="noopener noreferrer">{inner}</a>')
            else:
                # Keep labels in place even for absent or unsupported targets.
                parts.append(inner)
    inner = ''.join(parts).strip()
    if not inner:
        return ''
    style = (p.style.name or '').lower()
    islist = p._p.find('.//' + qn('w:numPr')) is not None
    if islist:
        return f'<li>{inner}</li>'
    if 'heading' in style:
        lvl = re.search(r'(\d)', style)
        n = min(int(lvl.group(1)) + 2, 5) if lvl else 3
        return f'<h{n}>{inner}</h{n}>'
    return f'<p>{inner}</p>'


def blocks_html(blocks, imgmap):
    out, inlist = [], False
    for b in blocks:
        if isinstance(b, Paragraph):
            frag = para_html(b, imgmap)
            if not frag:
                continue
            if frag.startswith('<li>'):
                if not inlist:
                    out.append('<ul>'); inlist = True
            elif inlist:
                out.append('</ul>'); inlist = False
            out.append(frag)
        else:  # nested table
            if inlist:
                out.append('</ul>'); inlist = False
            rows = []
            for r in b.rows:
                cells = ''.join(f'<td>{html.escape(cell_text(c)).replace(chr(10), "<br>")}</td>' for c in r.cells)
                rows.append(f'<tr>{cells}</tr>')
            out.append('<table class="t">' + ''.join(rows) + '</table>')
    if inlist:
        out.append('</ul>')
    return '\n'.join(out)

def cell_blocks_html(cell, imgmap):
    return blocks_html(list(iter_block_items(cell)), imgmap)

NORM_HUB = {'finnish': 'Finland', 'greek': 'Greece', 'italian': 'Italy',
            'norwegian': 'Norway', 'spanish': 'Spain'}

# Some labels use Greek/Cyrillic lookalike capitals (e.g. "Εstimated duration").
LOOKALIKE = {'Ε': 'E', 'А': 'A', 'С': 'C', 'О': 'O',
             'Т': 'T', 'Р': 'P', 'Ο': 'O', 'Η': 'H',
             'Ι': 'I', 'Κ': 'K', 'Μ': 'M', 'Ν': 'N',
             'Τ': 'T', 'Ρ': 'P', 'Β': 'B', 'Ζ': 'Z'}

def norm_label(s):
    s = ''.join(LOOKALIKE.get(ch, ch) for ch in s)
    s = s.replace('’', "'").replace('\xa0', ' ')
    return re.sub(r'\s+', ' ', s).strip().rstrip(':').strip()

# canonical field name -> label variants across the two template versions
ALIASES = {
    'title':        ['title of activity'],
    'duration':     ['estimated duration'],
    'target':       ['target group', 'grade', "teachers' or schools/contexts where the activity was designed for and/or implemented"],
    'disciplines':  ['related disciplines'],
    'authors':      ['author(s) of the activity'],
    'designed_at':  ['schools/contexts where the activity was designed (if applicable)',
                     'osn context where the activity was designed'],
    'implemented_at': ['schools/contexts where the activity was implemented (if applicable)'],
    'keywords':     ['other keywords', 'key-words related to fedoras themes'],
    'hubs':         ['reference to the hub'],
    'pillars':      ['fedora pillars', 'links to fedoras pillars'],
    'greencomp':    ['greencomp competence areas'],
    'description':  ['a brief description of the activity', 'objectives'],
}
FIELD_OF = {v: k for k, vs in ALIASES.items() for v in vs}

# GreenComp has four official areas; documents word them loosely.
GC_CANON = [
    ('embodying sustainability values', ['embodying', 'values']),
    ('embracing complexity in sustainability', ['embracing complexity', 'complexity']),
    ('envisioning sustainable futures', ['envisioning', 'futures', 'adaptability', 'exploratory thinking']),
    ('acting for sustainability', ['acting']),
]

CANON_OPTIONS = {
    'pillars': {'futures thinking', 'futures literacy', 'interdisciplinarity',
                'new languages', 'open schooling', 'sustainability competences'},
    'greencomp': {'embodying sustainability values', 'embracing complexity in sustainability',
                  'envisioning sustainable futures', 'acting for sustainability'},
    'hubs': {'finnish', 'greek', 'italian', 'norwegian', 'spanish'},
}

def is_canonical_option(key, val):
    opts = CANON_OPTIONS.get(key)
    return bool(opts) and val.strip().lower() in opts

def canon_greencomp(s):
    t = s.strip().lower()
    for canon, keys in GC_CANON:
        if any(k in t for k in keys):
            return canon
    return None

def parse_activity(path):
    fn = os.path.basename(path)
    slug = re.sub(r'[^a-z0-9]+', '-', os.path.splitext(fn)[0].lower()).strip('-')
    slug = slug.replace('fedoras-', '', 1)
    # Source rename confirmed 16 Oct 2026: extracted content is unchanged.
    # Keep published page IDs, editorial metadata and sequence links stable.
    if slug == 'osn-introduction-to-futures-studies':
        slug = 'osn-introductory-activity-to-futures-studies'
    if slug == "osn-the-story-of-my-town-s-stuff-exploring-product-life-cycles":
        slug += "-sustainability-and-futures"
    d = docx.Document(path)
    imgmap = save_images(d, slug)

    rec = {'slug': slug, 'file': fn, 'size': os.path.getsize(path),
           'meta': {}, 'raw_meta': {}, 'lists': {}, 'objectives': [], 'assessment': {},
           'sections': [], 'warnings': [], 'unspecified': []}

    outside, outside_title = [], 'Original document cover'

    def flush_outside():
        nonlocal outside
        frag = blocks_html(outside, imgmap)
        if frag.strip():
            rec['sections'].append({'title': outside_title, 'html': frag})
        outside = []

    # Traverse the body, not just its tables: references may follow the tables.
    for block in iter_block_items(d):
        if isinstance(block, Paragraph):
            text = ''.join(n.text or '' for n in block._p.findall('.//' + qn('w:t'))).strip()
            if text.casefold() in ('references', 'disclaimer'):
                flush_outside()
                outside_title = 'References' if text.casefold() == 'references' else 'Source disclaimer'
            else:
                outside.append(block)
            continue
        flush_outside()
        outside_title = 'Additional source text'
        t = block
        head = cell_text(t.rows[0].cells[0]).strip()
        hl = head.lower()
        if hl.startswith(('general information', 'contextual information', 'topic of activity')):
            for r in t.rows[1:]:
                if len(r.cells) < 2:
                    continue
                lab = norm_label(cell_text(r.cells[0]).split('\n')[0])
                if not lab:
                    continue
                field = FIELD_OF.get(lab.lower())
                items = cell_list(r.cells[1])
                picked = [i[0] for i in items if i[1]]
                # a cell is a pick-list when several options are offered and some are bold
                is_list = field in ('hubs', 'pillars', 'greencomp', 'disciplines') and len(items) > 1 and picked
                if field in ('hubs', 'pillars', 'greencomp', 'disciplines') or is_list:
                    vals = picked if picked else [i[0] for i in items]
                    key = field or lab
                    # No bold marking + the cell still holds the untouched option
                    # list => the author never made a choice. Don't read that as
                    # "all of them apply".
                    if not picked and len(items) >= 3 and all(
                            is_canonical_option(key, v) for v in vals):
                        rec['unspecified'].append(key)
                        rec['warnings'].append(f'"{lab}" left unfilled — every option still listed')
                        vals = []
                    rec['lists'][key] = vals
                if field:
                    rec['meta'][field] = cell_text(r.cells[1]) if not is_list else '; '.join(picked)
                rec['raw_meta'][lab] = cell_text(r.cells[1])
        elif ('objectives' in hl or hl.startswith('assessment')) and len(t.rows) > 1:
            # Keep the source labels and rich cell contents, including figures.
            table_rows = []
            for row in t.rows[1:]:
                unique = {c._tc: c for c in row.cells}
                cells = ''.join('<td>' + cell_blocks_html(c, imgmap) + '</td>' for c in unique.values())
                table_rows.append('<tr>' + cells + '</tr>')
            rec['sections'].append({'title': head, 'html': '<table>' + ''.join(table_rows) + '</table>'})
        else:
            # narrative tables: header row = section name, rest = content
            body = t.rows[1:] if len(t.rows) > 1 else t.rows
            frag = '\n'.join(cell_blocks_html(c, imgmap) for r in body for c in r.cells)
            title = head if len(t.rows) > 1 else ''
            if not title:
                first = cell_text(t.rows[0].cells[0]).split('\n')[0]
                title = first[:60]
                frag = cell_blocks_html(t.rows[0].cells[0], imgmap)
            if frag.strip():
                rec['sections'].append({'title': title, 'html': frag})

    flush_outside()

    # ---- normalise facets -------------------------------------------------
    hubs = rec['lists'].get('hubs', [])
    rec['hubs'] = sorted({NORM_HUB.get(h.lower().strip(), h.strip()) for h in hubs})
    rec['pillars'] = [p.strip() for p in rec['lists'].get('pillars', [])]
    gc_raw = rec['lists'].get('greencomp', [])
    rec['greencomp'] = sorted({c for c in (canon_greencomp(g) for g in gc_raw) if c})
    rec['greencomp_raw'] = gc_raw
    rec['disciplines_list'] = [re.sub(r'^Other\s*:?\s*', '', d).strip()
                               for d in rec['lists'].get('disciplines', []) if d.strip()]

    for k in ('title', 'duration', 'target', 'authors', 'designed_at',
              'implemented_at', 'keywords', 'description'):
        rec[k] = rec['meta'].get(k, '')
    if not rec['title']:
        rec['title'] = os.path.splitext(fn)[0].replace('FEDORAS ', '')
    rec['disciplines'] = rec['meta'].get('disciplines', '')
    rec['steam'] = next((v for k, v in rec['raw_meta'].items()
                         if k.lower().startswith('steam topics')), '')
    rec['is_osn'] = 'OSN' in fn or 'osn context' in ' '.join(rec['raw_meta']).lower()
    rec['images'] = len(imgmap)

    for k in ('duration', 'target', 'description', 'authors'):
        if not rec.get(k):
            rec['warnings'].append(f'no "{k}" given in the document')
    if not rec['hubs']:
        rec['warnings'].append('no hub country marked')
    if not rec['pillars']:
        rec['warnings'].append('no FEDORA pillar marked')
    return rec

