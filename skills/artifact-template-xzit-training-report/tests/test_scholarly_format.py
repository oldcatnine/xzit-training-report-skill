import sys,tempfile,unittest
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from scholarly_format import format_document,audit_links
from build_report import build
from test_tools import fixture

class ScholarlyTests(unittest.TestCase):
    def setUp(self):self.tmp=tempfile.TemporaryDirectory()
    def tearDown(self):self.tmp.cleanup()
    def make(self):
        data=fixture();data['references'][0]+=' DOI:10.1234/example.'
        p=Path(self.tmp.name)/'sample.docx';build(data,p);return Document(p)
    def test_links_and_superscript(self):
        d=self.make();a=audit_links(d);self.assertEqual(a['status'],'pass');self.assertEqual(a['linked_citation_count'],1)
        link=d.element.xpath('.//w:hyperlink[@w:anchor="xzit_ref_1"]')[0]
        self.assertEqual(''.join(link.xpath('.//w:t/text()')),'[1]')
    def test_doi_relationship(self):
        d=self.make();urls=[r.target_ref for r in d.part.rels.values() if r.is_external]
        self.assertNotIn('https://doi.org/10.1234/example',urls)
        self.assertNotIn('10.1234/example',next(p.text for p in d.paragraphs if p.style.name=='XZITReference'))
    def test_modern_docx_keeps_modern_compatibility(self):
        d=self.make();mode=d.settings.element.xpath('.//w:compatSetting[@w:name="compatibilityMode"]')[0]
        self.assertEqual(mode.get(qn('w:val')),'15')
    def test_repeat_is_idempotent(self):
        d=self.make();before=d.element.xml;format_document(d);self.assertEqual(before,d.element.xml)
    def test_thick_outer_and_thin_header_borders(self):
        d=self.make();table=d.tables[-1]
        for cell in table.rows[0].cells:
            self.assertEqual(cell._tc.xpath('./w:tcPr/w:tcBorders/w:top')[0].get(qn('w:sz')),'12')
            self.assertEqual(cell._tc.xpath('./w:tcPr/w:tcBorders/w:bottom')[0].get(qn('w:sz')),'4')
        for cell in table.rows[-1].cells:self.assertEqual(cell._tc.xpath('./w:tcPr/w:tcBorders/w:bottom')[0].get(qn('w:sz')),'12')
    def test_split_run_citation(self):
        d=self.make();p=next(p for p in d.paragraphs if p.style.name=='XZITBody');p.clear();p.add_run('Text [');p.add_run('1');p.add_run('] end.')
        format_document(d);self.assertEqual(p.text,'Text [1] end.');self.assertEqual(len(p._p.xpath('./w:hyperlink')),1)
    def test_unresolved_citation_rejected(self):
        d=self.make();p=next(p for p in d.paragraphs if p.style.name=='XZITBody');p.text='Text [99].'
        with self.assertRaises(ValueError):format_document(d)
    def test_doi_crosses_word_proofing_markers(self):
        d=self.make();p=next(p for p in d.paragraphs if p.style.name=='XZITReference');p.clear()
        p.add_run('[1] Source DOI:10.1234/j.example')
        mark=OxmlElement('w:proofErr');mark.set(qn('w:type'),'spellStart');p._p.append(mark)
        p.add_run('.2026.001.')
        format_document(d)
        self.assertEqual(p.text,'[1] Source')
    def test_loss_of_superscript_fails_audit(self):
        d=self.make();e=d.element.xpath('.//w:hyperlink[@w:anchor="xzit_ref_1"]//w:vertAlign')[0];e.getparent().remove(e)
        self.assertEqual(audit_links(d)['status'],'fail')
    def test_legacy_doi_link_removed_and_only_author_case_changes(self):
        d=self.make();p=next(p for p in d.paragraphs if p.style.name=='XZITReference')
        p.text="[1] MARTÍNEZ M, O’NEIL J. A Study on PV[J]. IEEE Access, 2020. DOI: "
        rid=d.part.relate_to('https://doi.org/10.1234/ex',RT.HYPERLINK,is_external=True)
        link=OxmlElement('w:hyperlink');link.set(qn('r:id'),rid)
        run=p.add_run('10.1234/example.');link.append(run._r);p._p.append(link)
        self.assertEqual(audit_links(d)['status'],'fail')
        format_document(d)
        self.assertEqual(p.text,"[1] Martínez M, O’Neil J. A Study on PV[J]. IEEE Access, 2020.")
        self.assertNotIn(rid,d.part.rels)
        self.assertEqual(audit_links(d)['status'],'pass')

    def test_combines_old_links_and_compacts_consecutive_numbers(self):
        d=self.make()
        for n in (2,3):
            d.add_paragraph(f'[{n}] Example {n}.',style='XZITReference')
        p=next(p for p in d.paragraphs if p.style.name=='XZITBody')
        p.clear()
        from scholarly_format import hyperlink
        source=p.add_run('')
        for n in (1,3):
            p._p.append(hyperlink(source._r,f'[{n}]',anchor=f'xzit_ref_{n}',superscript=True))
        p.add_run(' 与 [1,2,3]。')
        format_document(d)
        self.assertEqual(p.text,'[1,3] 与 [1-3]。')
        self.assertEqual(audit_links(d)['status'],'pass')
        xml=d.element.xml
        format_document(d)
        self.assertEqual(xml,d.element.xml)

    def test_word_saved_hyperlink_fields(self):
        d=self.make()
        d.add_paragraph('[2] Other Author. Example.',style='XZITReference')
        body=next(p for p in d.paragraphs if p.style.name=='XZITBody')
        ref=next(p for p in d.paragraphs if p.style.name=='XZITReference')
        from report_format import field
        body.clear()
        field(body,r'HYPERLINK \l "xzit_ref_1"','[1]')
        body.add_run('[2]')
        ref.add_run(' DOI:')
        field(ref,'HYPERLINK "https://doi.org/10.1234/example"','10.1234/example')
        ref.add_run('.')
        original=ref.text.split(' DOI:')[0]
        format_document(d)
        self.assertEqual(body.text,'[1,2]')
        self.assertEqual(ref.text,original)
        self.assertEqual(audit_links(d)['status'],'pass')

if __name__=='__main__':unittest.main()
