"""Build/repair, refresh in Word, and audit into a new output directory.

Machine checks never certify factual support or visual readiness.
"""
import argparse, copy, hashlib, json, shutil, subprocess, sys
from datetime import datetime, timezone
from pathlib import Path
from docx import Document
from build_report import build, ROOT
from audit_docx import audit
from audit_citations import audit as audit_citations
from scholarly_format import format_document, audit_links
from normalize_report import normalize
from document_roles import discover, role


def write_json(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def continuation_plan(data, doc, layout):
    """Add row breaks only where Word measured an actual page boundary.

    Rebuilds are bounded by the CLI. A header alone or a row taller
    than a page requires manual intervention rather than content deletion.
    """
    result = copy.deepcopy(data)
    positions = {x['table_index']: x['rows'] for x in layout['tables']}
    generated = [(i, t) for i, t in enumerate(doc.tables, 1)
                 if any(role(p)=='Table' for row in t.rows for cell in row.cells for p in cell.paragraphs)]
    cursor, changed, issues = 0, False, []
    for block in result['blocks']:
        if block['type'] != 'table': continue
        breaks = block.get('break_before_rows', [])
        starts, ends = [1] + breaks, breaks + [len(block['rows'])]
        new_breaks = set(breaks)
        for start, end in zip(starts, ends):
            if cursor >= len(generated):raise ValueError('Word table layout does not match input')
            index, table = generated[cursor]; cursor += 1
            rows = positions.get(index, [])
            if len(rows) != len(table.rows):raise ValueError('Word row layout count mismatch')
            header_page = rows[0]['end_page']
            overflow = False
            for local, item in enumerate(rows[1:], 1):
                original = start + local - 1
                if item['start_page'] != item['end_page']:
                    issues.append(f"{block['caption']}: row {original} spans pages; shorten or restructure the row")
                if item['start_page'] > header_page:
                    if local == 1:
                        issues.append(f"{block['caption']}: header orphaned; inspect table height")
                    else:
                        new_breaks.add(original)
                        overflow = True
                    header_page = item['start_page']
            # A repeated header can move the last row of a previously sized
            # chunk. Move the next boundary earlier instead of keeping both
            # boundaries and creating a nearly empty one-row continuation.
            if overflow and end < len(block['rows']):new_breaks.discard(end)
        if sorted(new_breaks) != breaks:
            block['break_before_rows'] = sorted(new_breaks); changed = True
    if cursor != len(generated):raise ValueError('Unmatched generated tables')
    return result, changed, sorted(set(issues))


def evidence_report(doc, evidence, role_map=None):
    body, refs, _ = discover(doc,role_map)
    records = evidence if isinstance(evidence, list) else []
    checks = []
    for n, p in refs.items():
        items = [x for x in records if x.get('reference') == n]
        checks.append({'reference': n, 'entry': p.text,
                       'claims_in_report': [p.text for p in body if f'[{n}]' in p.text],
                       'evidence': items,
                       'record_status': 'recorded' if items and all(x.get('claim') and x.get('source') and x.get('locator') and x.get('scope') for x in items) else 'pending',
                       'source_authenticity': 'pending', 'claim_support': 'pending'})
    return {'status': 'pending', 'references': checks,
            'note': 'Records are user-supplied leads. Verify publisher/full text and each claim; no network upload or automatic factual approval.'}


def run(input_path, output_dir, *, minimum=0, word=False, auto_continue=False,
        max_passes=5, role_map=None, normalize_styles=False, render=False):
    input_path, output_dir = Path(input_path).resolve(), Path(output_dir).resolve()
    if minimum < 0 or not 1 <= max_passes <= 8: raise ValueError('Invalid minimum or max_passes (1..8)')
    if auto_continue and not word:raise ValueError('--auto-continue requires --word')
    if render and not word:raise ValueError('--render requires --word')
    if normalize_styles and not role_map:raise ValueError('--normalize requires an explicit --role-map')
    if output_dir.exists():raise FileExistsError('Use a new output directory: '+str(output_dir))
    if input_path.suffix.lower() not in ('.json','.docx'):raise ValueError('Input must be JSON content or DOCX')
    if auto_continue and input_path.suffix.lower()!='.json':raise ValueError('Automatic continuation rebuild requires JSON input; inspect existing DOCX tables manually')
    data = json.loads(input_path.read_text(encoding='utf-8-sig')) if input_path.suffix.lower()=='.json' else None
    if data is not None:
        # Relative figure paths are resolved against the input JSON file.
        for block in data.get('blocks',[]):
            if block.get('type')=='figure':block['path']=str((input_path.parent / block['path']).resolve())
    output_dir.mkdir(parents=True)
    record = {'started_utc': datetime.now(timezone.utc).isoformat(),
              'input': str(input_path), 'input_sha256': digest(input_path),
              'reference_sha256': digest(ROOT/'assets/reference.docx'),
              'rules_sha256': digest(ROOT/'references/format-requirements.md'),
              'machine_status': 'pending', 'ready_for_submission': False,
              'manual_review': {'sources': 'pending', 'claims': 'pending', 'all_pages': 'pending'},
              'iterations': [], 'errors': []}
    write_json(output_dir/'run.json',record)
    try:
        pagination_issues=[]
        for pass_index in range(1,max_passes+1):
            raw = output_dir/f'pass-{pass_index}.docx'
            if data is not None:build(data,raw)
            else:
                doc=Document(input_path)
                if normalize_styles:record['normalization']=normalize(doc,role_map)
                record['formatting']=format_document(doc,role_map);doc.save(raw)
            current=raw
            layout=None
            if word:
                current=output_dir/f'pass-{pass_index}-updated.docx'
                pdf=output_dir/f'pass-{pass_index}.pdf'
                layout_path=output_dir/f'pass-{pass_index}-layout.json'
                command=['powershell.exe','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(ROOT/'scripts/refresh_fields.ps1'),
                         '-InputDocx',str(raw),'-OutputDocx',str(current),'-OutputPdf',str(pdf),'-OutputLayout',str(layout_path)]
                subprocess.run(command,check=True,timeout=180,capture_output=True,text=True)
                layout=json.loads(layout_path.read_text(encoding='utf-8-sig'))
            record['iterations'].append({'pass':pass_index,'docx':current.name,'pages':layout.get('page_count') if layout else None})
            if auto_continue:
                data,changed,pagination_issues=continuation_plan(data,Document(current),layout)
                if changed and pass_index<max_passes:continue
                if changed:pagination_issues.append('Pagination did not stabilize within max_passes')
            break
        shutil.copy2(current,output_dir/'report.docx')
        if word:shutil.copy2(pdf,output_dir/'report.pdf')
        if data is not None:write_json(output_dir/'resolved-input.json',data)
        doc=Document(output_dir/'report.docx')
        results={'format':audit(output_dir/'report.docx',role_map),
                 'citations':audit_citations(output_dir/'report.docx',minimum,role_map),
                 'links':audit_links(doc,role_map),
                 'pagination':{'status':'fail' if pagination_issues else ('pass' if auto_continue else 'pending'),'issues':pagination_issues}}
        if data is not None:
            expected=sum(b['type']=='equation' for b in data['blocks'])
            actual=len(doc.element.xpath('.//m:oMath'))
            results['equations']={'expected':expected,'actual':actual,'status':'pass' if actual==expected else 'fail'}
            write_json(output_dir/'evidence.json',evidence_report(doc,data.get('evidence',[])))
        else:write_json(output_dir/'evidence.json',evidence_report(doc,[],role_map))
        write_json(output_dir/'checks.json',results)
        if render:
            renderer=shutil.which('pdftoppm')
            if not renderer:raise RuntimeError('pdftoppm is unavailable; PDF saved, visual review remains pending')
            pages=output_dir/'pages';pages.mkdir()
            subprocess.run([renderer,'-scale-to','1600','-png',str(output_dir/'report.pdf'),str(pages/'page')],check=True,timeout=180,capture_output=True)
            record['rendered_pages']=len(list(pages.glob('page-*.png')))
            if record['rendered_pages']!=layout['page_count']:raise RuntimeError('Rendered page count disagrees with Word')
        record['checks']='checks.json'
        record['machine_status']='fail' if any(x.get('status',x.get('machine_status'))=='fail' for x in results.values()) else 'pass'
        if results['format']['manual_review_required']:record['manual_review']['format_details']='pending'
        record['fields_updated']=word
        record['artifacts']={p.name:digest(p) for p in (output_dir/'report.docx',output_dir/'report.pdf') if p.exists()}
    except Exception as error:
        record['machine_status']='fail';record['errors'].append(f'{type(error).__name__}: {error}')
        if isinstance(error,subprocess.CalledProcessError):record['errors'].append(str(error.stderr))
        write_json(output_dir/'run.json',record)
        raise
    write_json(output_dir/'run.json',record)
    return record


if __name__=='__main__':
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('input',type=Path);ap.add_argument('output_dir',type=Path)
    ap.add_argument('--min-references',type=int,default=0)
    ap.add_argument('--word',action='store_true');ap.add_argument('--auto-continue',action='store_true')
    ap.add_argument('--max-passes',type=int,default=5);ap.add_argument('--render',action='store_true')
    ap.add_argument('--role-map',type=Path);ap.add_argument('--normalize',action='store_true')
    a=ap.parse_args()
    try:
        result=run(a.input,a.output_dir,minimum=a.min_references,word=a.word,auto_continue=a.auto_continue,
                   max_passes=a.max_passes,role_map=json.loads(a.role_map.read_text(encoding='utf-8-sig')) if a.role_map else None,
                   normalize_styles=a.normalize,render=a.render)
        print(json.dumps(result,ensure_ascii=False,indent=2))
        raise SystemExit(result['machine_status']=='fail')
    except (ValueError,FileExistsError,RuntimeError,subprocess.SubprocessError) as error:
        ap.exit(1,str(error)+'\n')
