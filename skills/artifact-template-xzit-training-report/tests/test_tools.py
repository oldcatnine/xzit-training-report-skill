import sys,tempfile,unittest,hashlib
from pathlib import Path
from docx import Document
from docx.shared import Pt
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from build_report import build
from audit_docx import audit
from report_format import child

def fixture():
    return dict(title_zh='排版工具测试',title_en='Format Tool Test',student='测试姓名',student_id='0000',**{'class':'测试班级'},supervisor='测试教师',major='测试专业',date='2026年11月13日',abstract_zh='仅用于验证排版工具，不构成实训研究。',abstract_en='This is a formatting test fixture only.',keywords_zh=['排版','测试'],keywords_en=['formatting','test'],blocks=[{'type':'heading','level':1,'text':'1 测试章节'},{'type':'paragraph','text':'仅验证排版，不作研究结论[1]。'},{'type':'heading','level':2,'text':'1.1 测试小节'},{'type':'heading','level':4,'text':'1.1.1.1 四级测试'},{'type':'table','caption':'表1-1 测试表','rows':[['项目','说明'],['用途','单元测试']]}],conclusion='工具测试结束。',acknowledgements='此为测试占位，非真实致谢。',references=['测试作者. 测试文献[J]. 测试刊名, 2026.'])

class ToolsTest(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.path=Path(self.tmp.name)/'test.docx';build(fixture(),self.path)
    def tearDown(self):self.tmp.cleanup()
    def mutate(self,func):
        doc=Document(self.path);func(doc);doc.save(self.path);return audit(self.path)
    def test_source_immutable_and_no_overwrite(self):
        src=ROOT/'assets/reference.docx';before=hashlib.sha256(src.read_bytes()).hexdigest()
        with self.assertRaises(FileExistsError):build(fixture(),self.path)
        self.assertEqual(before,hashlib.sha256(src.read_bytes()).hexdigest())
    def test_valid_except_unupdated_toc(self):
        failures=audit(self.path)['failures']
        self.assertTrue(failures)
        self.assertTrue(all('目录待更新' in x['location'] for x in failures),failures)
    def test_run_override_detected(self):
        out=self.mutate(lambda d:setattr(next(p for p in d.paragraphs if p.style.name=='XZITBody').runs[0].font,'size',Pt(16)))
        self.assertTrue(any(x['rule']=='F-font' for x in out['failures']))
    def test_margin_detected(self):
        out=self.mutate(lambda d:setattr(d.sections[0],'left_margin',Pt(20)))
        self.assertTrue(any(x['rule']=='O01-O03' for x in out['failures']))
    def test_heading_half_lines(self):
        def f(d):
            p=next(p for p in d.paragraphs if p.style.name=='XZITH2');child(p._p.get_or_add_pPr(),'w:spacing',beforeLines=0)
        self.assertTrue(any(x['rule']=='F13-F16' for x in self.mutate(f)['failures']))
    def test_keywords_case(self):
        data=fixture();data['keywords_en']=['Bad Keyword']
        with self.assertRaises(ValueError):build(data,Path(self.tmp.name)/'bad.docx')
    def test_missing_identity(self):
        data=fixture();del data['student_id']
        with self.assertRaises(ValueError):build(data,Path(self.tmp.name)/'bad.docx')
    def test_unsupported_formula_not_silently_dropped(self):
        data=fixture();data['blocks'].append({'type':'equation','text':'x=y'})
        with self.assertRaises(ValueError):build(data,Path(self.tmp.name)/'bad.docx')
    def test_false_bold_value_detected(self):
        def f(d):next(p for p in d.paragraphs if p.style.name=='XZITAbstractTitleEN').runs[0].bold=False
        self.assertTrue(any(x['rule']=='F08' for x in self.mutate(f)['failures']))
    def test_always_requires_manual_review(self):self.assertFalse(audit(self.path)['ready_for_submission'])

if __name__=='__main__':unittest.main()
