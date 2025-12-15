import os
import unittest
from md2docx import MarkdownToDocx

class TestMd2Docx(unittest.TestCase):
    def setUp(self):
        self.input_file = 'test_input.md'
        self.output_file = 'test_output.docx'
        
        # Create a dummy input file
        with open(self.input_file, 'w', encoding='utf-8') as f:
            f.write("# Test Document\n\nHello, **World**!")

    def tearDown(self):
        # Clean up files
        if os.path.exists(self.input_file):
            os.remove(self.input_file)
        if os.path.exists(self.output_file):
            os.remove(self.output_file)

    def test_conversion_default(self):
        converter = MarkdownToDocx()
        converter.convert(open(self.input_file).read(), self.output_file)
        self.assertTrue(os.path.exists(self.output_file))
        self.assertGreater(os.path.getsize(self.output_file), 0)

    def test_conversion_preserve_breaks(self):
        converter = MarkdownToDocx(preserve_breaks=True)
        converter.convert(open(self.input_file).read(), self.output_file)
        self.assertTrue(os.path.exists(self.output_file))

if __name__ == '__main__':
    unittest.main()

