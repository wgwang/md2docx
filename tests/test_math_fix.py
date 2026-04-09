
import unittest
from lxml import etree
from md2docx.converter import MarkdownToDocx

class TestMathFix(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocx()
        self.ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
    
    def test_sqrt_fix(self):
        # Case 1: m:rad with empty m:deg
        omml = """
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:rad>
                <m:deg><m:t></m:t></m:deg>
                <m:e><m:t>x</m:t></m:e>
            </m:rad>
        </m:oMath>
        """
        tree = etree.fromstring(omml)
        self.converter._clean_omml(tree)
        
        # Verify m:deg is gone
        deg = tree.find('.//m:deg', self.ns)
        self.assertIsNone(deg, "Empty m:deg should be removed")
        
        # Verify m:radPr/m:degHide is on
        radPr = tree.find('.//m:radPr', self.ns)
        self.assertIsNotNone(radPr, "m:radPr should be created")
        degHide = radPr.find('./m:degHide', self.ns)
        self.assertIsNotNone(degHide, "m:degHide should be present")
        self.assertEqual(degHide.get(f'{{{self.ns["m"]}}}val'), 'on')

    def test_sqrt_missing_deg(self):
        # Case 2: m:rad with missing m:deg
        omml = """
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:rad>
                <m:e><m:t>x</m:t></m:e>
            </m:rad>
        </m:oMath>
        """
        tree = etree.fromstring(omml)
        self.converter._clean_omml(tree)
        
        # Verify m:radPr/m:degHide is on
        radPr = tree.find('.//m:radPr', self.ns)
        self.assertIsNotNone(radPr, "m:radPr should be created")
        degHide = radPr.find('./m:degHide', self.ns)
        self.assertIsNotNone(degHide, "m:degHide should be present")

    def test_nary_sup_fix(self):
        # Case 3: m:nary with empty m:sup
        omml = """
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:nary>
                <m:sub><m:t>i</m:t></m:sub>
                <m:sup></m:sup>
                <m:e><m:t>x</m:t></m:e>
            </m:nary>
        </m:oMath>
        """
        tree = etree.fromstring(omml)
        self.converter._clean_omml(tree)
        
        sup = tree.find('.//m:sup', self.ns)
        self.assertIsNone(sup, "Empty m:sup should be removed")
        
        naryPr = tree.find('.//m:naryPr', self.ns)
        self.assertIsNotNone(naryPr, "m:naryPr should be created")
        supHide = naryPr.find('./m:supHide', self.ns)
        self.assertIsNotNone(supHide, "m:supHide should be present")

    def test_nary_sub_fix(self):
        # Case 4: m:nary with empty m:sub
        omml = """
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:nary>
                <m:sub></m:sub>
                <m:sup><m:t>n</m:t></m:sup>
                <m:e><m:t>x</m:t></m:e>
            </m:nary>
        </m:oMath>
        """
        tree = etree.fromstring(omml)
        self.converter._clean_omml(tree)
        
        sub = tree.find('.//m:sub', self.ns)
        self.assertIsNone(sub, "Empty m:sub should be removed")
        
        naryPr = tree.find('.//m:naryPr', self.ns)
        self.assertIsNotNone(naryPr, "m:naryPr should be created")
        subHide = naryPr.find('./m:subHide', self.ns)
        self.assertIsNotNone(subHide, "m:subHide should be present")

    def test_nary_missing_sup(self):
        # Case 5: m:nary with missing m:sup (implicit empty)
        omml = """
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:nary>
                <m:sub><m:t>i</m:t></m:sub>
                <m:e><m:t>x</m:t></m:e>
            </m:nary>
        </m:oMath>
        """
        tree = etree.fromstring(omml)
        self.converter._clean_omml(tree)
        
        naryPr = tree.find('.//m:naryPr', self.ns)
        self.assertIsNotNone(naryPr, "m:naryPr should be created")
        supHide = naryPr.find('./m:supHide', self.ns)
        self.assertIsNotNone(supHide, "m:supHide should be present")

if __name__ == '__main__':
    unittest.main()
