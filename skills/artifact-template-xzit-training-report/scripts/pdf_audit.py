"""Compare cached Word TOC entries against text and printed PDF page numbers."""
import re
from document_roles import paragraphs, role

ROMAN = re.compile(r'M{0,3}(?:CM|CD|D?C{0,3})(?:XC|XL|L?X{0,3})(?:IX|IV|V?I{0,3})')


def label(token):
    token = token.strip().upper()
    if re.fullmatch(r'[1-9]\d*',token):
        return 'decimal',int(token)
    if token and ROMAN.fullmatch(token):
        values = dict(I=1,V=5,X=10,L=50,C=100,D=500,M=1000)
        return 'roman',sum(-values[ch] if i+1 < len(token) and values[ch]<values[token[i+1]]
                           else values[ch] for i,ch in enumerate(token))
    return None


def compact(text):
    return re.sub(r'\s+','',text).casefold()


def read_pages(path):
    from pypdf import PdfReader
    pages = []
    for page in PdfReader(path).pages:
        candidates = set()
        def visitor(text, cm, tm, font, size):
            # Word exports upright pages. Include text lying in the footer band.
            y = tm[4]*cm[1] + tm[5]*cm[3] + cm[5]
            if 8 < y < 65 and label(text):
                candidates.add(label(text))
        text = page.extract_text(visitor_text=visitor) or ''
        resources = page.get('/Resources',{}).get_object() if page.get('/Resources') else {}
        fonts = {}
        for ref in (resources.get('/Font') or {}).get_object().values() if resources.get('/Font') else ():
            font = ref.get_object()
            descendants = font.get('/DescendantFonts') or [font]
            def embedded(item):
                item = item.get_object()
                descriptor = item.get('/FontDescriptor')
                return item.get('/Subtype') == '/Type3' or bool(descriptor and any(
                    key in descriptor.get_object() for key in ('/FontFile','/FontFile2','/FontFile3')))
            fonts[str(font.get('/BaseFont',font.get('/Name','unknown')))] = all(embedded(item) for item in descendants)
        pages.append({'lines':text.splitlines(),'printed':next(iter(candidates)) if len(candidates)==1 else None,
                      'fonts':fonts})
    return pages


def inspect_pdf(doc, path):
    try:
        pages = read_pages(path)
    except ImportError:
        return {'status':'pending','reason':'pypdf is not installed'}
    entries = []
    for p in paragraphs(doc):
        if role(p) != 'TOC':
            continue
        match = re.fullmatch(r'(.+?)\t+\s*(\S+)\s*',p.text)
        if match and label(match[2]):
            entries.append((match[1],label(match[2])))
        elif p.text.strip():
            entries.append((p.text.strip(),None))
    issues, pending, matches = [], [], []
    prior = None
    for index,page in enumerate(pages,1):
        current = page['printed']
        if current is None:
            if index > 1:pending.append(f'Printed label not reliably extracted on PDF page {index}')
            prior = None
            continue
        if prior is not None:
            expected = (prior[0],prior[1]+1) if prior[0]==current[0] else ('decimal',1) if prior[0]=='roman' else None
            if current != expected:
                issues.append(f'Unexpected printed page sequence on PDF page {index}: {current}; expected {expected}')
        prior = current
    if not any(page['printed'] == ('decimal',1) for page in pages):
        pending.append('First Arabic body page label (1) not confirmed')
    if not entries:
        pending.append('No readable Word TOC entries')
    for title, expected in entries:
        key = compact(title)
        hits = []
        for index,page in enumerate(pages,1):
            lines = [compact(line) for line in page['lines'] if line.strip()]
            matches_title = any(''.join(lines[i:i+length]) == key
                                for i in range(len(lines)) for length in (1,2,3))
            if matches_title:
                hits.append((index,page['printed']))
        if expected is None or len(hits)!=1 or hits[0][1] is None:
            pending.append(f'TOC entry needs manual verification: {title}')
        elif hits[0][1] != expected:
            issues.append(f'TOC entry {title}: cached {expected}, printed {hits[0][1]} on PDF page {hits[0][0]}')
        else:
            matches.append({'title':title,'physical_page':hits[0][0],'printed':expected})
    if pages and pages[0]['printed'] is not None:
        issues.append('Cover displays a printed page number')
    if len(pages)>1 and pages[1]['printed'] not in (None,('roman',1)):
        issues.append('Physical page 2 must display Roman I')
    font_resources = {font:{'embedded_on_all_listed_pages':all(page['fonts'][font]
                      for page in pages if font in page.get('fonts',{}))}
                      for font in sorted({font for page in pages for font in page.get('fonts',{})})}
    return {'status':'fail' if issues else 'pending' if pending else 'pass',
            'page_count':len(pages),'issues':issues,'pending':pending,'matched_entries':matches,
            'font_resources':font_resources}
