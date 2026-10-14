from setuptools import setup, find_namespace_packages

setup(
    name='ooxml-recovery-tool',
    version='0.2.1',
    # ort/core and ort/core/handlers have no __init__.py, so namespace discovery is required.
    packages=find_namespace_packages(include=['ort', 'ort.*']),
    install_requires=[
        'beautifulsoup4',
        'lxml',
        'python-docx',
        'pillow',
    ],
    entry_points={
        'console_scripts': [
            'ort=ort.__main__:main'
        ]
    },
    author='Jiyun Kim',
    description='Extracts document content and images from damaged OOXML files (.docx, .xlsx, .pptx)',
    long_description=open('README.md', encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/ew6o6/OOXML-Recovery-Tool',
    license='MIT',
    python_requires='>=3.8',
)
