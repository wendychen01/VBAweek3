Attribute VB_Name = "Module1"
Option Explicit

Sub cal()
'¤èªk¤@ ¥ÎRange
Range("E1").Value = Range("A1").Value + Range("C1").Value 'E1Äæ­È=A1Äæ­È+C1Äæ­È
Range("E2").Value = Range("A1").Value - Range("C1").Value 'E2Äæ­È=A1Äæ­È-C1Äæ­È
Range("E3").Value = Range("A1").Value * Range("C1").Value 'E3Äæ­È=A1Äæ­È*C1Äæ­È
Range("E4").Value = Range("A1").Value / Range("C1").Value 'E4Äæ­È=A1Äæ­È/C1Äæ­È
'¤èªk¤G ¥ÎCells²Ä¤@¦C¨ì²Ä¤­Äæ
Cells(1, 5).Value = Cells(1, 1).Value + Cells(1, 3).Value 'E1Äæ­È=A1Äæ­È+C1Äæ­È
Cells(2, 5).Value = Cells(1, 1).Value - Cells(1, 3).Value 'E2Äæ­È=A1Äæ­È-C1Äæ­È
Cells(3, 5).Value = Cells(1, 1).Value * Cells(1, 3).Value 'E3Äæ­È=A1Äæ­È*C1Äæ­È
Cells(4, 5).Value = Cells(1, 1).Value / Cells(1, 3).Value 'E4Äæ­È=A1Äæ­È/C1Äæ­È
'¤èªk¤T ¥ÎCells²Ä¤@¦C¨ì²ÄEÄæ
Cells(1, "E").Value = Cells(1, "A").Value + Cells(1, "C").Value 'E1Äæ­È=A1Äæ­È+C1Äæ­È
Cells(2, "E").Value = Cells(1, "A").Value - Cells(1, "C").Value 'E2Äæ­È=A1Äæ­È-C1Äæ­È
Cells(3, "E").Value = Cells(1, "A").Value * Cells(1, "C").Value 'E3Äæ­È=A1Äæ­È*C1Äæ­È
Cells(4, "E").Value = Cells(1, "A").Value / Cells(1, "C").Value 'E4Äæ­È=A1Äæ­È/C1Äæ­È

End Sub
