#!/usr/bin/env python

from distutils.core import setup

setup(
    name="exam-grading",
    version="0.1",
    description="python / jupyter tools for visualizing exam grades",
    author="Stefan Schulz",
    author_email="stefan.schulz.2507@gmail.com",
    url="https://github.com/valgarf/exam-grading",
    packages=["exam_grading"],
    python_requires=">=3.8",
    install_requires=[
        "wheel",
        "jupyter",
        "qgrid",
        "pandas",
        "jupyter_contrib_nbextensions",
        "ipysheet",
        "openpyxl",
        "xlrd",
        "matplotlib",
        "nbopen",
        "ipyfilechooser",
    ],
)

# also run:
#   win10_post_install.bat (on windows)
# or
#   linux_post_install.sh (on Linux)
# Enable notebook
