Attribute VB_Name = "Module2"
Option Explicit

Sub �Ѥp��j()
Attribute �Ѥp��j.VB_Description = "�f�n�ƶq�Ѥp��j�Ƨ�"
Attribute �Ѥp��j.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' �Ѥp��j ����
' �f�n�ƶq�Ѥp��j�Ƨ�
'
' �ֳt��: Ctrl+s
'Create By WEN TI CHEN 2020/3/15
    Columns("A:B").Select '�ʧ@�@-���AB���
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear '�ʧ@�G-��ƱƧǳ]�w�A�ھڤf�n�ƶq�OB��컼������
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort '���d��v�����Ƨ�
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
Sub �p���`�M()
Attribute �p���`�M.VB_Description = "�f�n�ƶq�`�M"
Attribute �p���`�M.VB_ProcData.VB_Invoke_Func = "c\n14"
'
' �p���`�M ����
' �f�n�ƶq�`�M
'
' �ֳt��: Ctrl+c
'
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("E2").Select
End Sub
Sub ����()
Attribute ����.VB_Description = "�N�ƶq�`�X������"
Attribute ����.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' ���� ����
' �N�ƶq�`�X������
'
' �ֳt��: Ctrl+a
'
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-5]:R[413]C[-5])"
    Range("G2").Select
End Sub
