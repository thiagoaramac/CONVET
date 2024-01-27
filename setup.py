from setuptools import find_packages
from cx_Freeze import setup, Executable


options = {
    'build_exe': {
        'includes': [
            'cx_Logging', 'idna',
        ],
        'packages': [
            'asyncio', 'flask', 'jinja2', 'dash', 'plotly', 'waitress'
        ],
        'excludes': ['tkinter']
    }
}

executables = [
    Executable('server.py',
               base='console',
               targetName='CONVET.exe')
]

setup(
    name='CONVET',
    packages=find_packages(),
    version='1.0.0',
    description='Ranking CONVET',
    executables=executables,
    options=options
)
