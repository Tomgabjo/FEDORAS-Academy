"""Convert the Word equation structures used in the teaching resources to native MathML."""
import html
import re
from docx.oxml.ns import qn


def math_html(node):
    local = lambda n: n.tag.rsplit('}', 1)[-1]
    def render(n):
        tag = local(n)
        def lookup(key):
            found = n.find(qn('m:' + key))
            return render(found) if found is not None else '<mrow/>'
        if tag == 'oMath':
            return '<span class="formula"><math><mrow>' + ''.join(map(render, n)) + '</mrow></math></span>'
        if tag == 'oMathPara':
            return ''.join(render(c) for c in n if local(c) == 'oMath')
        if tag in {'e','num','den','sub','sup','deg'}:
            return '<mrow>' + ''.join(map(render,n)) + '</mrow>'
        if tag == 'r':
            text = ''.join(x.text or '' for x in n.findall(qn('m:t')))
            parts = re.findall(r'\d+(?:\.\d+)?|[^\W\d_]+|\s+|.', text)
            return ''.join(f'<{k}>{html.escape(t)}</{k}>' for t in parts
                for k in ['mn' if t[0].isdigit() else 'mi' if t.isalpha() else 'mtext' if t.isspace() else 'mo'])
        if tag == 'f': return '<mfrac>'+lookup('num')+lookup('den')+'</mfrac>'
        if tag == 'sSub': return '<msub>'+lookup('e')+lookup('sub')+'</msub>'
        if tag == 'sSup': return '<msup>'+lookup('e')+lookup('sup')+'</msup>'
        if tag == 'rad':
            deg = n.find(qn('m:deg'))
            if deg is None or not ''.join(deg.itertext()).strip(): return '<msqrt>'+lookup('e')+'</msqrt>'
            return '<mroot>'+lookup('e')+lookup('deg')+'</mroot>'
        if tag == 'bar':
            pos = n.find('./'+qn('m:barPr')+'/'+qn('m:pos'))
            above = pos is None or pos.get(qn('m:val'))!='bot'
            name = 'mover' if above else 'munder'
            return f'<{name}>'+lookup('e')+f'<mo stretchy="true">¯</mo></{name}>'
        if tag == 'd':
            def delim(key, default):
                v = n.find('./'+qn('m:dPr')+'/'+qn('m:'+key))
                return html.escape(v.get(qn('m:val'),default) if v is not None else default)
            return '<mrow><mo>'+delim('begChr','(')+'</mo>'+lookup('e')+'<mo>'+delim('endChr',')')+'</mo></mrow>'
        if tag.endswith('Pr'): return ''
        raise ValueError('New Word equation construct: '+tag)
    return render(node)
