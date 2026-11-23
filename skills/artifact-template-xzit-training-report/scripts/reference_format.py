"""Document-only bibliography presentation rules."""
import re

DOI = re.compile(r'10\.\d{4,9}/[^\s，。；]+', re.I)
DOI_SPAN = re.compile(r'[ \t]*(?<![\w/])(?:DOI[ \t]*[:：]?[ \t]*)?'
                      r'(?:(?:https?://)?(?:dx\.)?doi\.org/)?' + DOI.pattern + r'[。]?', re.I)
DOI_URL = re.compile(r'^https?://(?:dx\.)?doi\.org/', re.I)
# The author block precedes the first full stop followed by a title. Initials
# written with periods and followed by another initial or comma stay in it.
AUTHOR_LINE = re.compile(r'^\[\d+\]\s*(?P<authors>(?:\b[^\W\d_]\.(?=\s*[,，]|\s+[^\W\d_]\.)|[^.])+?)\.(?=\s+\S)',re.UNICODE)
NAME_WORD = re.compile(r'[^\W\d_]+',re.UNICODE)
ORG = re.compile(r'\b(agency|association|commission|committee|council|department|institute|laboratory|organization|university|society|international|corporation)\b',re.I)


def author_spans(text):
    match = AUTHOR_LINE.match(text)
    if not match:
        return []
    edits = []
    offset = match.start('authors')
    fields = re.split(r'([,，])',match['authors'])
    for index,field in enumerate(fields):
        words = list(NAME_WORD.finditer(field))
        comma_initials = index+2 < len(fields) and bool(re.fullmatch(
            r'\s*[A-Z]\.?(?:\s+[A-Z]\.?)?\s*',fields[index+2]))
        # Personal names have a surname and an initial or given name.
        if (len(words)>=2 or len(words)==1 and comma_initials) and re.fullmatch(
            r"[^\W\d_]+(?:[\s.'’\-]+[^\W\d_]+)*[\s.'’\-]*",field.strip()
        ) and not ORG.search(field):
            for word in words:
                token = word[0]
                if len(token)>1 and token.isupper() and token.lower()!=token:
                    edits.append((offset+word.start(),offset+word.end(),token.capitalize()))
        offset += len(field)
    return edits


def styled_reference(text):
    out = DOI_SPAN.sub('',text).rstrip()
    for start,end,value in reversed(author_spans(out)):
        out = out[:start] + value + out[end:]
    return out
