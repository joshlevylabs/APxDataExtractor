from setuptools import setup, find_packages

setup(
    name='APSequenceRunner',
    version='0.1',
    packages=find_packages(),
    install_requires=[
        'openpyxl',
        'pythonnet'
    ],
    entry_points={
        'console_scripts': [
            'APSequenceRunner=APxDataExtractor:main',
        ],
    },
)
