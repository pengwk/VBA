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
__status__ = "BraveHeart"


def wait():
    import time
    time.sleep(2)
    return None


def set_clipboard(text):
    import win32clipboard as wincb
    import win32con

    wincb.OpenClipboard()
    wincb.EmptyClipboard()
    wincb.SetClipboardData(win32con.CF_UNICODETEXT, text)
    wincb.CloseClipboard()
    return None


def excel_find_all(excel_app, what):
    """
    找到全部what所在位置(Row, Column)，以列表的形式返回

    VBA code
    expression.Find(What, After, LookIn, LookAt,
                    SearchOrder, SearchDirection,
                     MatchCase, MatchByte, SearchFormat)
    expression .FindNext(After)
    """
    _result = []
    first_occur = excel_app.Cells.Find(what)
    if first_occur is not None:
        _result.append((first_occur.Row, first_occur.Column))
        next_one = first_occur
        while True:
            next_one = excel_app.Cells.FindNext(next_one)
            if (next_one.Row, next_one.Column) not in _result:
                _result.append((next_one.Row, next_one.Column))
            else:
                return _result
    else:
        return None


def main():
    import os

    from win32com.client import Dispatch

    # real work
    # 打开Word 与 Excel
    excel = Dispatch('Excel.Application')
    excel.Visible = True

    report_template_path = ur"E:\东莞理工学院\IP通信\IP实验模板.doc"

    report_folder = ur"E:\东莞理工学院\IP通信\201441302623 彭未康"

    guide_book_path = ur"E:\东莞理工学院\IP通信\中兴实物实验v2.0.xlsx"
    exp_guide_book = excel.Workbooks.Open(guide_book_path)

    for i in range(15, exp_guide_book.WorkSheets.Count+1):
        exp_guide_book.WorkSheets(i).Activate()
        experiment_name = exp_guide_book.WorkSheets(i).Name

        word = Dispatch("Word.Application")
        word.Visible = True
        report_doc = word.Documents.Open(report_template_path)

        # 查找设置
        # 全局查找常量
        wdFindContinue = 1

        my_find = word.Selection.Find
        my_find.ClearFormatting()
        my_find.Wrap = wdFindContinue

        # 粘贴常量
        wdFormatPlainText = 22
        # 实验名
        report_doc.Activate()
        my_find.Text = "{{experiment_name}}"
        my_find.Execute()

        set_clipboard(experiment_name)
        word.Selection.Paste()
        wait()

        # 拓扑图
        exp_guide_book.WorkSheets(i).Activate()
        exp_guide_book.WorkSheets(i).Shapes.SelectAll()
        wait()
        excel.Selection.Copy()

        report_doc.Activate()
        my_find.Text = "{{topological_graph}}"
        wait()
        my_find.Execute()
        wait()
        word.Selection.PasteSpecial(0,0,0,0,4)
        wait()

        # 代码
        exp_guide_book.WorkSheets(i).Activate()

        special_words = excel_find_all(excel, u"命令内容")
        start_cell = excel.Cells(special_words[0][0], special_words[0][1])

        special_words = excel_find_all(excel, u"exit")
        end_cell = excel.Cells(special_words[-1][0], special_words[-1][1])

        excel.Range(start_cell, end_cell).Select()

        wait()
        excel.Selection.Copy()

        wait()
        report_doc.Activate()
        wait()
        my_find.Text = "{{code}}"
        wait()
        my_find.Execute()
        wait()
        word.Selection.PasteAndFormat(wdFormatPlainText)
        wait()

        # 另存为 退出
        report_doc.Activate()
        wait()
        filename = os.path.join(report_folder, experiment_name)
        word.Documents(1).SaveAs(filename)
        word.Documents(1).Close()
    return None


if __name__ == "__main__":
    main()
    