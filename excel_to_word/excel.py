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


def main():
    
    
    # WorkSheets 工作表
    # 选择工作表
    # exp_guide_book.WorkSheets(1).Select()           # 单选 1开始索引
    # exp_guide_book.WorkSheets([5, 6, 7]).Select()   # 多选
    # 激活工作表
    # exp_guide_book.WorkSheets(6).Activate()         # 只能激活一个工作表

    # 工作表总数
    # exp_guide_book.WorkSheets.Count

    
    import os
    import time

    from win32com.client import Dispatch

    # real work
    # 打开Word 与 Excel
    excel = Dispatch('Excel.Application')
    excel.Visible = True

    report_template_path = ur"E:\东莞理工学院\IP通信\IP实验模板.doc"

    report_folder = ur"E:\东莞理工学院\IP通信\201441302623 彭未康\"

    guide_book_path = ur"E:\东莞理工学院\IP通信\中兴实物实验v2.0.xlsx"
    exp_guide_book = excel.Workbooks.Open(guide_book_path)



    for i in range(2, exp_guide_book.WorkSheets.Count):
        exp_guide_book.WorkSheets(i).Activate()
        experiment_name = exp_guide_book.WorkSheets(1).Name

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
        my_find.Text = "{{experiment_name}}"
        my_find.Execute()
        word.Selection.PasteAndFormat(wdFormatPlainText)

        # 拓扑图
        exp_guide_book.WorkSheets(6).Shapes.SelectAll()
        excel.Selection.Copy()
        my_find.Text = "{{topological_graph}}"
        my_find.Execute()
        word.Selection.PasteSpecial(0,0,0,0,4)

        # 代码
        exp_guide_book.WorkSheets(6).Range("B10:B70").Select()
        excel.Selection.Copy()
        my_find.Text = "{{code}}"
        my_find.Execute()
        word.Selection.PasteAndFormat(wdFormatPlainText)

        # 另存为 退出
        filename = os.path.join(report_folder, experiment_name)
        word.Documents(1).SaveAs(filename)
        word.Documents(1).Close()

    # Shape
    print exp_guide_book.WorkSheets(6).Shapes.Count

    exp_guide_book.WorkSheets(6).Select()       # 重要 没有时会出现 错误：请求的图形已被锁定供选择。
    exp_guide_book.WorkSheets(6).Shapes.SelectAll()

    excel.Selection.Copy()

    # 表格内容
    exp_guide_book.WorkSheets(6).Activate()
    exp_guide_book.WorkSheets(6).Range("B16").Value

    # Word
    word = Dispatch("Word.Application")
    word.Visible = True

    # 新建文档
    exp_doc = word.Documents.Add()
    # word.Selection.Paste()

   # 打开文档
    ip_exp_template_path = ur"E:\东莞理工学院\IP通信\IP实验模板.doc"
    ip_exp_template = word.Documents.Open(ip_exp_template_path)

    # 激活文档
    # Documents("Sales.doc").Activate
    # Documents(1).Activate
    word.Documents(1).Activate()

    # 选择文本
    # With Selection.Find
    #    .Text = "^$"
    #    .Replacement.Text = ""
    #    .Forward = True
    #    .Wrap = wdFindContinue
    #    .Format = False
    #    .MatchCase = False
    #    .MatchWholeWord = False
    #    .MatchByte = False
    #    .MatchWildcards = False
    #    .MatchSoundsLike = False
    #    .MatchAllWordForms = False
    # End With
    # wrap
    #  1 The find operation continues if the beginning or end of the search range is reached.

    wdFindContinue = 1

    my_find = word.Selection.Find
    my_find.ClearFormatting()
    my_find.Wrap = wdFindContinue
    my_find.Text = "{{code}}"
    my_find.Text = "{{topological_graph}}"
    my_find.Text = "{{experiment_name}}"
    my_find.Execute()




    # 粘贴
    # 粘贴为图片
    word.Selection.PasteSpecial(0,0,0,0,4)

    # 粘贴为无格式纯文本
    # Selection.PasteAndFormat (wdFormatPlainText)
    word.Selection.PasteAndFormat(22)

    # 关闭 
    exp_guide_book.Close(True) # 不会关闭Excel自身

    # 保存 另存为
    word.Documents(1).SaveAs("ExperimentName")
    return None

    # excel find
    # excel.Cells.Find("exit").Row
    # excel.Cells.Find("exit").Value
    # excel.Cells.Find("exit").Column
    # excel.Cells.Find("exit").Address u'$B$17'
    # 查找全部使用FindNext实现 FindNext是循环的，全部查找到后会到第一个地方 用Address找到位置



    # 错误
    # 被呼叫方拒绝接收呼叫。
    # 指定的数据类型无效。
    # 命令失败
# 此方法和属性无效，因为剪切板是空的。
#
# com_error: (-2147352567, '\xb7\xa2\xc9\xfa\xd2\xe2\xcd\xe2\xa1\xa3', (0, u'Microsoft Excel', u'\u7c7b Range \u7684 FindNext \u65b9\u6cd5\u65e0\u6548', u'xlmain11.chm', 0, -2146827284), None)
#     Set myUsedRange = Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol))

if __name__ == "__main__":
    main()

