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
# jupyter contrib nbextension install --user
# jupyter nbextension install --py widgetsnbextension --user
# jupyter nbextension enable --py --user widgetsnbextension
# jupyter nbextension install --py qgrid --user
# jupyter nbextension enable --py --user qgrid
# jupyter nbextension install --py ipysheet --user
# jupyter nbextension enable --py --user ipysheet
# python -m nbopen.install_win
