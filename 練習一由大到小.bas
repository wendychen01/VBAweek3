Attribute VB_Name = "Module1"
Option Explicit

Sub �Ѥj��p()
Attribute �Ѥj��p.VB_Description = "�f�n�ƶq�Ƨ�"
Attribute �Ѥj��p.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' ����1 ����
' �f�n�ƶq�Ƨ�
'
' �ֳt��: Ctrl+q
'
'Create By WEN TI CHEN 2020/3/15

    Range("B1").Select '�ʧ@�@-���B1�x�s��
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear '�ʧ@�G-��ƱƧǳ]�w�A�ھڤf�n�ƶq�OB��컼���
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort '���d��v�����Ƨ�
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'End of create
End Sub
