import copy
import tempfile
import unittest
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from check_minutes_review import verify, digest

class ReviewTests(unittest.TestCase):
    def test_stale_empty_or_open_reviews_cannot_pass(self):
        with tempfile.TemporaryDirectory() as directory:
            data=Path(directory)/'data.json'; data.write_text('{}')
            minutes=Path(directory)/'minutes.docx'; minutes.write_bytes(b'reviewed artifact')
            dimension={'verdict':'PASS','paragraph_checks':[{'location':'section 1','assessment':'complete','basis':'specific passage'}]}
            review={'rules_version':'minutes-quality-v4','reviewer':'reviewer-b','independent':True,'reviewed_at':'2026-11-03',
                    'data_sha256':digest(data),'minutes_sha256':digest(minutes),'delivery_suitability':dimension,'evidence':dimension,'verdict':'PASS','issues':[]}
            for name in ('synthesis','section_coherence','tracking_selection','challenge_checks'):
                review[name]={'verdict':'PASS','checks':[{'location':'section 1','assessment':'verified','basis':'source passage and synthesis'}]}
            self.assertEqual(verify(review,data,minutes),[])
            for name in ('synthesis','section_coherence','tracking_selection','challenge_checks'):
                invalid=copy.deepcopy(review); invalid.pop(name)
                self.assertTrue(verify(invalid,data,minutes))
            invalid=copy.deepcopy(review); invalid['rules_version']='minutes-quality-v3'
            self.assertTrue(verify(invalid,data,minutes))
            invalid=copy.deepcopy(review); invalid['delivery_suitability']['paragraph_checks']=[]
            self.assertTrue(verify(invalid,data,minutes))
            invalid=copy.deepcopy(review); invalid['issues']=[{'status':'open'}]
            self.assertTrue(verify(invalid,data,minutes))
            minutes.write_bytes(b'edited artifact')
            self.assertTrue(verify(review,data,minutes))
