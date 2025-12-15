import argparse
import sys
import os
from .converter import MarkdownToDocx

def main():
    parser = argparse.ArgumentParser(description='Convert Markdown to Docx.')
    parser.add_argument('input_file', help='Path to the input markdown file.')
    parser.add_argument('output_file', help='Path to the output docx file.')
    parser.add_argument('--template', help='Path to the template file (docx/dotx).')
    parser.add_argument('--font', help='Default font name to use (default: Microsoft YaHei if no template provided).')
    parser.add_argument('--preserve-breaks', action='store_true', help='Preserve single line breaks as hard breaks.')

    args = parser.parse_args()

    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' not found.")
        sys.exit(1)

    if args.template and not os.path.exists(args.template):
        print(f"Error: Template file '{args.template}' not found.")
        sys.exit(1)

    try:
        with open(args.input_file, 'r', encoding='utf-8') as f:
            md_content = f.read()
        
        # Default font to Microsoft YaHei only if no template is provided and no font is specified
        font_name = args.font
        if font_name is None and args.template is None:
            font_name = 'Microsoft YaHei'

        converter = MarkdownToDocx(template_path=args.template, font_name=font_name, preserve_breaks=args.preserve_breaks)
        converter.convert(md_content, args.output_file)
        
        print(f"Successfully converted '{args.input_file}' to '{args.output_file}'.")

    except Exception as e:
        print(f"Error during conversion: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
