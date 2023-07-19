from setuptools import setup

APP = ['RandomisationParkProReakt_Hamburg.py']
DATA_FILES = ['template.xlsx']
OPTIONS = {'argv_emulation': True}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)