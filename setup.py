import setuptools
from setuptools import find_packages, setup
import os
with open('requirements.txt') as f:
    pkgs = f.read().splitlines()
    
setup(
    name="DDs",
    version="0.0.1",
    description="Collect DD online",
    long_description_content_type="",
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Operating System :: OS Independent",
    ],
    url="https://github.com/felgabeee",
    author="felgabe",
    author_email="felix.gabet@edhec.com",
    keywords=["AMF","Directors' dealing"],
    install_requires=pkgs,
    packages=find_packages(),
    license="MIT",
    include_package_data=True,
)
