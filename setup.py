from setuptools import setup, find_packages

setup(
    name="md2docx-pro",
    version="0.1.0",
    description="A Python package to convert Markdown to Docx with custom font and embedded images.",
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
