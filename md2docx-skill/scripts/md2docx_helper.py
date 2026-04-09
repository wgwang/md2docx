import sys
import argparse
import io
import os
from md2docx import MarkdownToDocx

def main():
    parser = argparse.ArgumentParser(description='Agent Helper for md2docx')
    parser.add_argument('--template', help='Path to docx template')
    parser.add_argument('--output', help='Output file path. If omitted, writes to stdout (not recommended for binary, but possible if redirected)')
    
    args = parser.parse_args()
    
    # Read Markdown from stdin
    md_content = sys.stdin.read()
    
    if not md_content.strip():
        print("Error: No markdown content provided on stdin.", file=sys.stderr)
        sys.exit(1)
    
    converter = MarkdownToDocx(template_path=args.template)
    
    if args.output:
        converter.convert(md_content, args.output)
        print(f"Successfully converted to {args.output}", file=sys.stderr)
    else:
        # Return bytes to stdout
        docx_stream = converter.convert_to_bytes(md_content)
        sys.stdout.buffer.write(docx_stream.getbuffer())

if __name__ == "__main__":
    main()
