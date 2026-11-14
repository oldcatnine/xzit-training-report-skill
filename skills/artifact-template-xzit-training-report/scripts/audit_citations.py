"""Audit numeric citations with shared, legacy-compatible body discovery."""
import argparse, json, re
from collections import Counter
from pathlib import Path
from docx import Document
from document_roles import discover, CITE, numbers


def audit(path, minimum=0, role_map=None):
    if minimum < 0: raise ValueError('Minimum cannot be negative')
    body, refs, fail = discover(Document(path), role_map)
    cited, locations = [], {}
    for pos, p in enumerate(body, 1):
        for match in CITE.finditer(p.text):
            try: values = numbers(match[0])
            except ValueError as error:
                fail.append(str(error)); continue
            for n in values:
                cited.append(n)
                locations.setdefault(n, []).append({'body_paragraph': pos, 'context': p.text[:90]})
    nums, used = set(refs), set(cited)
    if len(nums) < minimum: fail.append(f'References {len(nums)} below task minimum {minimum}')
    if nums != set(range(1, len(nums) + 1)): fail.append('Bibliography is not continuous from 1')
    if used - nums: fail.append('Missing bibliography for ' + str(sorted(used - nums)))
    if nums - used: fail.append('Uncited bibliography ' + str(sorted(nums - used)))
    first = list(dict.fromkeys(cited))
    if first != sorted(first): fail.append('First citation order differs from numbering')
    values = [m[0].rstrip('.,;，。').lower() for p in refs.values() if (m := re.search(r'10\.\d{4,9}/\S+', p.text, re.I))]
    if len(values) != len(set(values)): fail.append('Duplicate DOI entries')
    return {'status': 'fail' if fail else 'pass', 'task_minimum': minimum,
            'reference_count': len(nums), 'distinct_cited_count': len(used),
            'first_occurrence_order': first, 'citation_counts': dict(Counter(cited)),
            'failures': fail, 'locations': locations,
            'scope': 'Numbering and correspondence only; source existence and factual support require review.'}


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('docx', type=Path); ap.add_argument('--min-references', type=int, default=0)
    ap.add_argument('--output', type=Path); ap.add_argument('--role-map', type=Path)
    a = ap.parse_args()
    if a.min_references < 0: ap.error('Minimum cannot be negative')
    if a.output and a.output.exists(): ap.error('Refusing to overwrite output')
    result = audit(a.docx, a.min_references, json.loads(a.role_map.read_text(encoding='utf-8-sig')) if a.role_map else None)
    if a.output: a.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: v for k, v in result.items() if k != 'locations'}, ensure_ascii=False, indent=2))
    raise SystemExit(bool(result['failures']))
