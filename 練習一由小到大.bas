Attribute VB_Name = "Module2"
Option Explicit

Sub 由小到大()
Attribute 由小到大.VB_Description = "口罩數量由小到大排序"
Attribute 由小到大.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' 由小到大 巨集
' 口罩數量由小到大排序
'
' 快速鍵: Ctrl+s
'Create By WEN TI CHEN 2020/3/15
    Columns("A:B").Select '動作一-選擇AB欄位
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear '動作二-資料排序設定，根據口罩數量是B欄位遞憎順序
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort '全範圍逐行執行排序
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D9").Select
    'End of create
End Sub
Sub 計算總和()
Attribute 計算總和.VB_Description = "口罩數量總和"
Attribute 計算總和.VB_ProcData.VB_Invoke_Func = "c\n14"
'
' 計算總和 巨集
' 口罩數量總和
'
' 快速鍵: Ctrl+c
'
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("E2").Select
End Sub
Sub 平均()
Attribute 平均.VB_Description = "將數量總合做平均"
Attribute 平均.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' 平均 巨集
' 將數量總合做平均
'
' 快速鍵: Ctrl+a
'
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-5]:R[413]C[-5])"
    Range("G2").Select
End Sub
