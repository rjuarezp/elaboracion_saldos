# Work around mbcs bug in distutils.
# http://bugs.python.org/issue10945
import codecs

try:
    codecs.lookup('mbcs')
except LookupError:
    ascii = codecs.lookup('ascii')
    func = lambda name, enc=ascii: {True: enc}.get(name=='mbcs')
    codecs.register(func)

import setuptools

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setuptools.setup(
    name="example-pkg-YOUR-USERNAME-HERE", # Replace with your own username
    version="0.0.1",
    author="Ricardo Juarez",
    author_email="rjuarezp@gmail.com",
    description="Elaboracion de saldos",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/rjuarezp/gertru.git",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.7',
)