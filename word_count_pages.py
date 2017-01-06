#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
    This module does
    Date created: 4/20/2016
    Date last modified: 4/25/2016
    Python Version: 2.7.10
"""

__author__ = "pengwk"
__copyright__ = "Copyright 2016, pengwk"
__credits__ = [""]
__license__ = "Private"
__version__ = "0.1"
__maintainer__ = "pengwk"
__email__ = "pengwk2@gmail.com"
__status__ = "Brave Heart"


def main():
    from win32com.client import Dispatch

    # open Word
    word = Dispatch('Word.Application')
    word.Visible = True

    doc_path = ur"E:\东莞理工学院\EDA\experment\201441302623 彭未康 实验一.doc"
    word = word.Documents.Open(doc_path)

    # get number of sheets
    word.Repaginate()
    num_of_sheets = word.ComputeStatistics(2)
    print num_of_sheets
    return None


if __name__ == "__main__":
    main()
    