from setuptools import setup, find_packages

import os

# Read the contents of your README files
this_directory = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(this_directory, 'README.md'), encoding='utf-8') as f:
    readme_zh = f.read()
with open(os.path.join(this_directory, 'README_EN.md'), encoding='utf-8') as f:
    readme_en = f.read()

# Combine them: English first (standard for PyPI) followed by Chinese
long_description = f"{readme_en}\n\n---\n\n{readme_zh}"

setup(
    name="md2docx-pro",
    version="0.1.1",  # Incrementing version to allow re-publish
    description="A Python package to convert Markdown to Docx with custom font and embedded images.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Gemini Assistant",
    packages=find_packages(),
    install_requires=[
        "python-docx",
        "markdown-it-py",
        "requests",
        "mdit-py-plugins",
        "latex2mathml",
        "mathml2omml",
    ],
    entry_points={
        "console_scripts": [
            "md2docx = md2docx.cli:main",
        ],
    },
)
