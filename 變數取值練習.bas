Attribute VB_Name = "Module1"
Option Explicit

Sub cal()
'��k�@ ��Range
Range("E1").Value = Range("A1").Value + Range("C1").Value 'E1���=A1���+C1���
Range("E2").Value = Range("A1").Value - Range("C1").Value 'E2���=A1���-C1���
Range("E3").Value = Range("A1").Value * Range("C1").Value 'E3���=A1���*C1���
Range("E4").Value = Range("A1").Value / Range("C1").Value 'E4���=A1���/C1���
'��k�G ��Cells�Ĥ@�C��Ĥ���
Cells(1, 5).Value = Cells(1, 1).Value + Cells(1, 3).Value 'E1���=A1���+C1���
Cells(2, 5).Value = Cells(1, 1).Value - Cells(1, 3).Value 'E2���=A1���-C1���
Cells(3, 5).Value = Cells(1, 1).Value * Cells(1, 3).Value 'E3���=A1���*C1���
Cells(4, 5).Value = Cells(1, 1).Value / Cells(1, 3).Value 'E4���=A1���/C1���
'��k�T ��Cells�Ĥ@�C���E��
Cells(1, "E").Value = Cells(1, "A").Value + Cells(1, "C").Value 'E1���=A1���+C1���
Cells(2, "E").Value = Cells(1, "A").Value - Cells(1, "C").Value 'E2���=A1���-C1���
Cells(3, "E").Value = Cells(1, "A").Value * Cells(1, "C").Value 'E3���=A1���*C1���
Cells(4, "E").Value = Cells(1, "A").Value / Cells(1, "C").Value 'E4���=A1���/C1���

End Sub
