"""
Setup script for building xltable package.
"""

from setuptools import find_packages, setup

setup_params = dict(
    name="xltable",
    description="A Python package to ease writing tables of data to Excel",
    packages=find_packages(),
    test_suite="nose.collector",
    version="0.2.7",
    install_requires=["pandas>=0.12.0"],
    extras_require={
        "xlsxwriter": ["xlsxwriter>=0.7.2"],
        "pywin32": ["pywin32>=219"],
        "xlwt": ["xlwt>=0.7.5"]
    },
    tests_require=["nose>=1.2.1"],
    author="Renshaw Bay",
    author_email="tony@pyxll.com",
    url="https://github.com/fkarb/xltable",
    classifiers=["License :: OSI Approved :: MIT License"],
)

if __name__ == '__main__':
    setup(**setup_params)
