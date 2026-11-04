import copy
import json
import sys
import tempfile
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from check_ledger_review import verify, DIMENSIONS
from check_minutes_review import digest
from build_deliverables import review_packet

class LedgerReviewTests(unittest.TestCase):
    def test_complete_packet_preserves_assertions_and_sources(self):
        record={'id':'X1','nature':'待核实线索','title':'标题','description':'描述',
                'impact':'没有责任人','verification':{'cause_hypothesis':'假设','conclusion':'待查'},'related_ids':['X2']}
        evidence={'id':'E1','record_ids':['X1'],'source_file':'source.txt','note':'单位尚未核实'}
        packet=review_packet({'records':[record],'evidence':[evidence]})['records'][0]
        for key, value in record.items():
            self.assertEqual(packet[key], value)
        self.assertIn('没有',packet['attention_terms'])
        self.assertEqual(packet['evidence'][0]['note'],evidence['note'])
        self.assertEqual(packet['evidence'][0]['source_file'],'source.txt')

    def test_coverage_bindings_and_joint_review(self):
        with tempfile.TemporaryDirectory() as d:
            data=Path(d)/'data.json';data.write_text(json.dumps({'records':[{'id':'X1'}]}))
            ledger=Path(d)/'ledger.xlsx';ledger.write_bytes(b'current ledger')
            minutes=Path(d)/'minutes.docx';minutes.write_bytes(b'current minutes')
            check={'id':'X1','verdict':'PASS',**{k:{'assessment':'judgment','basis':'source position'} for k in DIMENSIONS}}
            review={'rules_version':'ledger-quality-v1','reviewer':'B','reviewed_at':'2026-11-03','independent':True,
                    'data_sha256':digest(data),'ledger_sha256':digest(ledger),'record_checks':[check],
                    'cross_document':{'verdict':'NOT_APPLICABLE','reason':'ledger only'},'verdict':'PASS','issues':[]}
            self.assertEqual(verify(review,data,ledger),[])
            for mutate in [lambda r:r.update(record_checks=[]),lambda r:r['record_checks'].append(copy.deepcopy(check)),
                           lambda r:r['record_checks'][0].pop('source_and_time'),
                           lambda r:r.update(issues=[{'status':'open'}]),lambda r:r.update(independent=False)]:
                bad=copy.deepcopy(review);mutate(bad)
                self.assertTrue(verify(bad,data,ledger))
            self.assertTrue(verify(review,data,ledger,minutes))
            review.update(minutes_sha256=digest(minutes),cross_document={'verdict':'PASS','checks':[{
                'claim':'quantity','minutes_location':'section 1','record_ids':['X1'],'assessment':'consistent','basis':'source line 2'}]})
            self.assertEqual(verify(review,data,ledger,minutes),[])
            minutes.write_bytes(b'new minutes')
            self.assertTrue(verify(review,data,ledger,minutes))
            ledger.write_bytes(b'new ledger')
            self.assertTrue(verify(review,data,ledger,minutes))
