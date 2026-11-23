"""Read-only report inspection and a reviewable repair plan."""
import hashlib
import json
from docx import Document
from audit_docx import audit
from audit_citations import audit as audit_citations
from scholarly_format import audit_links
from document_roles import suggest_roles
from document_roles import discover
from metadata_audit import audit_metadata


def inspect(path, output, minimum=0, mapping=None, overrides=None, crossref=False):
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    doc = Document(path)
    results = {'format':audit(path,mapping,overrides),
               'citations':audit_citations(path,minimum,mapping),
               'links':audit_links(doc,mapping)}
    suggestions = suggest_roles(doc,mapping)
    if crossref:
        _, refs, _ = discover(doc,mapping)
        results['reference_metadata'] = audit_metadata(refs)
    after = hashlib.sha256(path.read_bytes()).hexdigest()
    if before != after:
        raise RuntimeError('Input changed while the read-only inspection was running')
    status = 'fail' if any(r.get('status',r.get('machine_status')) == 'fail' for r in results.values()) else 'pass'
    summary = {'mode':'check','machine_status':status,'ready_for_submission':False,'input_sha256':before,
               'input_sha256_after':after,'checks':results,'role_suggestions':suggestions,
               'course_rules':overrides or {}}
    output.mkdir(parents=True)
    (output/'inspection.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    lines = ['# 格式检查与修改建议', '', '检查方式：只读', '', '## 已发现的问题','']
    for item in results['format']['failures']:
        lines.append(f"- {item['location']}｜{item['rule']}｜{item['detail']}｜依据：{item['source']}"
                     f"｜实测：{item['actual'] if item['actual'] is not None else '未提取'}"
                     f"｜要求：{item['expected'] if item['expected'] is not None else '见规则'}"
                     f"｜建议：{'确认角色后按规则修正' if item['fixable'] else '对照规则人工复核'}")
    for name in ('citations','links'):
        for item in results[name]['failures']:
            lines.append(f'- {name}｜{item}')
    if not any(r.get('failures') for r in results.values()):
        lines.append('- 自动检查未发现错误。')
    if crossref:
        lines.extend(['','## DOI 书目信息核对',''])
        for item in results['reference_metadata']['items']:
            lines.append(f"- 文献 [{item['reference']}]｜{item.get('doi','无 DOI')}｜{item['status']}｜"
                         f"{item.get('reason',item.get('checks',''))}")
    lines.extend(['','## 待确认的样式与人工检查',''])
    for name,items in suggestions['unmapped_styles'].items():
        for item in items:
            lines.append(f"- 段落 {item['paragraph']}｜{name}｜{item['preview']}｜"
                         f"{item['role'] or '待识别'}｜{item['reason']}")
    for item in results['format']['manual_review_required']:
        if isinstance(item, str):
            lines.append('- '+item)
    (output/'inspection.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    return {'mode':'check','machine_status':status,'ready_for_submission':False,
            'inspection':'inspection.json','plan':'inspection.md','input_sha256':before}
