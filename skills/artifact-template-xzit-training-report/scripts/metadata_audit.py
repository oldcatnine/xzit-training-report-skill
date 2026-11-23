"""Opt-in Crossref metadata comparison. Only DOI identifiers are transmitted."""
import json
import re
import unicodedata
from urllib.parse import quote
from urllib.request import Request, urlopen
from urllib.error import HTTPError


def fetch_doi(doi):
    request = Request('https://api.crossref.org/works/'+quote(doi,safe=''),
                      headers={'User-Agent':'xzit-training-report-skill/doi-check','Accept':'application/json'})
    with urlopen(request,timeout=8) as response:
        return json.loads(response.read(2_000_000))['message']


def comparable(text):
    return ''.join(ch for ch in unicodedata.normalize('NFKC',text).casefold() if ch.isalnum())


def audit_metadata(refs, fetch=fetch_doi):
    items, cache = [], {}
    for number, paragraph in refs.items():
        text = paragraph if isinstance(paragraph,str) else paragraph.text
        match = re.search(r'10\.\d{4,9}/[^\s，。；]+',text,re.I)
        item = {'reference':number,'status':'pending'}
        items.append(item)
        if not match:
            item['reason'] = 'No DOI found'
            continue
        doi = match[0].rstrip('.,;')
        item['doi'] = doi
        if doi.casefold() not in cache:
            try:
                cache[doi.casefold()] = fetch(doi)
            except HTTPError as error:
                cache[doi.casefold()] = f'Crossref HTTP {error.code}'
            except (OSError,KeyError,ValueError) as error:
                cache[doi.casefold()] = type(error).__name__
        entry = cache[doi.casefold()]
        if isinstance(entry,str):
            item['reason'] = entry
            continue
        title = next(iter(entry.get('title') or []),'')
        authors = [a.get('family') or a.get('name') for a in entry.get('author',[])]
        authors = [a for a in authors if a]
        years = sorted({parts[0][0] for key in ('published','published-print','published-online','issued')
                        if (parts := (entry.get(key) or {}).get('date-parts')) and parts[0]})
        raw = comparable(text[:match.start()])
        fields = {'title':bool(comparable(title) and comparable(title) in raw) if title else None,
                  'first_author':comparable(authors[0]) in raw if authors else None,
                  'year':any(re.search(r'(?<!\d)'+str(year)+r'(?!\d)',text[:match.start()]) for year in years) if years else None}
        item.update(title=title,first_author=authors[0] if authors else None,years=years,checks=fields,
                    status='matched' if all(value is True for value in fields.values()) else
                    ('review' if False in fields.values() else 'pending'))
    return {'status':'pass' if items and all(item['status']=='matched' for item in items) else 'pending',
            'scope':'Bibliographic metadata only; claim support requires source review','items':items}
