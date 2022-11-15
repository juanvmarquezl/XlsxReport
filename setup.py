# Always prefer setuptools over distutils
from setuptools import setup, find_packages

# To use a consistent encoding
from codecs import open
from os import path

# The directory containing this file
HERE = path.abspath(path.dirname(__file__))

# Get the long description from the README file
with open(path.join(HERE, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

# This call to setup() does all the work
setup(
    name="XlsxReport",
    version="0.0.9",
    description="Simple Excel Reports Tool",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/juanvmarquezl/XlsxReport",
    author="Juan Márquez",
    author_email="",
    license="MIT",
    classifiers=[
        "Intended Audience :: Developers",
        "Intended Audience :: Information Technology",
        "Intended Audience :: Other Audience",
        "Development Status :: 2 - Pre-Alpha",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Operating System :: OS Independent"
    ],
    packages=["XlsxReport"],
    include_package_data=True,
    install_requires=['XlsxWriter>=3.0.3']
)

# Compile setup file
#   python setup.py sdist bdist_wheel

# test XlsxReport in test.pypi
#   twine upload --repository-url https://test.pypi.org/legacy/ dist/*