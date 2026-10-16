import re
from pathlib import Path

from setuptools import find_packages, setup

VERSION = re.search(r'__version__ = "([^"]+)"', Path('ort/__init__.py').read_text(encoding='utf-8')).group(1)

setup(
    name='ooxml-recovery-tool',
    version=VERSION,
    packages=find_packages(include=['ort', 'ort.*']),
    install_requires=[
        'beautifulsoup4',
        'lxml',
        'python-docx',
        'pillow',
        'openpyxl',
        'python-pptx',
    ],
    entry_points={
        'console_scripts': [
            'ort=ort.main:main'
        ]
    },
    author='Jiyun Kim',
    description='Extracts document content and images from damaged OOXML files (.docx, .xlsx, .pptx)',
    long_description=Path('README.md').read_text(encoding='utf-8'),
    long_description_content_type='text/markdown',
    url='https://github.com/ew6o6/OOXML-Recovery-Tool',
    license='MIT',
    python_requires='>=3.8',
)
