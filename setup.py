import os
import setuptools

def get_version(version_tuple):
    if not isinstance(version_tuple[-1], int):
        return '.'.join(map(str, version_tuple[:-1]))
    return '.'.join(map(str, version_tuple))

init = os.path.join(
    os.path.dirname(__file__), 'src', 'excel2dict', '__init__.py'
)

version_line = list(
    filter(lambda l: l.startswith('VERSION'), open(init))
)[0]

VERSION = get_version(eval(version_line.split('=')[-1]))

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="excel2dict",
    version=VERSION,
    author="Ryosuke Wakaba",
    author_email="wakaba.ryosule@gmail.com",
    description="A converting excel file to python data structure package",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/rwakaba/excel2dict",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
