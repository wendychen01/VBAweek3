Attribute VB_Name = "Module1"
Option Explicit

Sub 由大到小()
Attribute 由大到小.VB_Description = "口罩數量排序"
Attribute 由大到小.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 巨集1 巨集
' 口罩數量排序
'
' 快速鍵: Ctrl+q
'
'Create By WEN TI CHEN 2020/3/15

    Range("B1").Select '動作一-選擇B1儲存格
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear '動作二-資料排序設定，根據口罩數量是B欄位遞減順序
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort '全範圍逐行執行排序
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'End of create
End Sub
