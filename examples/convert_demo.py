import os
import io
from md2docx import MarkdownToDocx

def run_demo():
    # 1. Read the demo markdown file
    md_file = 'examples/demo.md'
    if not os.path.exists(md_file):
        print(f"Error: {md_file} not found.")
        return

    with open(md_file, 'r', encoding='utf-8') as f:
        md_content = f.read()

    # 2. Convert using default settings (to file)
    print("Converting to demo_default.docx...")
    converter = MarkdownToDocx()
    converter.convert(md_content, 'examples/demo_default.docx')

    # 3. Convert using a Professional template
    if os.path.exists('professional_template.docx'):
        print("Converting using professional_template.docx...")
        prof_converter = MarkdownToDocx(template_path='professional_template.docx')
        prof_converter.convert(md_content, 'examples/demo_professional.docx')

    # 4. Convert to a byte stream (for AI agents/Web APIs)
    print("Converting to in-memory byte stream...")
    byte_converter = MarkdownToDocx()
    docx_bytes = byte_converter.convert_to_bytes(md_content)
    
    # Verify the byte stream
    print(f"Generated {docx_bytes.getbuffer().nbytes} bytes of DOCX data.")
    
    # Save the byte stream to a file to verify
    with open('examples/demo_from_bytes.docx', 'wb') as f:
        f.write(docx_bytes.getbuffer())

    print("\nAll demo conversions completed. Check the 'examples/' folder.")

if __name__ == "__main__":
    run_demo()
