from setuptools import setup, find_packages
'''
pip 설치 및 ooxml-parse 명령어 제공
ooxml_parser.main:main 함수 진입점 지정
'''
setup(
    name='ooxml_parser',
    version='0.1.0',
    packages=find_packages(),
    install_requires=[
        'beautifulsoup4',
        'lxml',
        'tabulate',
        'python-docx'
    ],
    entry_points={
        'console_scripts': [
            'ooxml-parse=ooxml_parser.__main__:main'
        ]
    },
    author='Your Name',
    description='A recovery and extraction tool for damaged OOXML documents (.docx, .xlsx, .pptx)',
    python_requires='>=3.8'
)
