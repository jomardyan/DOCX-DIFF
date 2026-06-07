"""DOCX Diff - Setup configuration for distribution."""
from pathlib import Path
from setuptools import setup, find_packages

# Read the README file
readme_file = Path(__file__).parent / "README.md"
long_description = readme_file.read_text(encoding="utf-8") if readme_file.exists() else ""

# Read requirements
requirements_file = Path(__file__).parent / "requirements.txt"
requirements = []
if requirements_file.exists():
    requirements = [
        line.strip() 
        for line in requirements_file.read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.startswith("#")
    ]

setup(
    name="docxdiff",
    version="2.0.0",
    author="DOCX Diff Team",
    author_email="your.email@example.com",
    description="Compare DOCX files and display differences with multiple output formats",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/docx-diff",
    project_urls={
        "Bug Tracker": "https://github.com/yourusername/docx-diff/issues",
        "Documentation": "https://github.com/yourusername/docx-diff#readme",
        "Source Code": "https://github.com/yourusername/docx-diff",
    },
    py_modules=["docxdiff", "docxdiff_engine", "docxdiff_reports"],
    python_requires=">=3.8",
    install_requires=requirements,
    extras_require={
        "dev": [
            "pytest>=6.0",
            "pytest-cov>=2.0",
            "flake8>=3.9",
            "black>=21.0",
            "mypy>=0.900",
        ],
    },
    entry_points={
        "console_scripts": [
            "docxdiff=docxdiff:main",
        ],
        "gui_scripts": [
            "docxdiff-gui=docxdiff:launch_gui",
        ],
    },
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Intended Audience :: End Users/Desktop",
        "Topic :: Office/Business",
        "Topic :: Text Processing",
        "Topic :: Utilities",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: OS Independent",
        "Environment :: Console",
        "Environment :: Win32 (MS Windows)",
        "Environment :: X11 Applications",
        "Environment :: MacOS X",
    ],
    keywords="docx diff compare word document comparison",
    license="MIT",
)
