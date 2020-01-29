Attribute VB_Name = "Module1"
Sub PorGrPor()
    ActiveCell.FormulaR1C1 = "=round(rc[-2]-average(rc[-3],rc[-1]),0)"
    ActiveCell.Interior.Color = vbYellow
End Sub
Sub Itog()
    ActiveCell.FormulaR1C1 = "=round(average(rc[-3]:rc[-1]),0)"
    ActiveCell.Interior.Color = "5296274"
End Sub

Sub dot()
For i = 11 To 137
    'Worksheets("Для вставки в расчет").Activate
    dot1 = Cells(i, "a").Value
    'Worksheets("Для расчета базовый").Activate
    For j = 11 To 131
        If Cells(j, "o").Value Like dot1 Then
            Range(Cells(j, "p"), Cells(j, "p")).Select
            Selection.Copy
            'Worksheets("Для вставки в расчет").Activate
            Cells(i, "c").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Exit For
        End If
    Next j
Next i
'Worksheets("Для вставки в расчет").Activate

End Sub
Sub brakes()
    YesNo = MsgBox("Макрос рассчитан на то, что ТОЛЬКО В ПЕРВОЙ СТРОКЕ каждого вида торможений имеются значения. Лучше сделать копию файла! Продолжить?", vbYesNo)
    If YesNo = vbYes Then
    For K = 1 To 8
        flag = 0
        flag1 = 0
        flag2 = 0
        flag3 = 0
        flag4 = 0
        If Cells(K * 4 + 1, 9) <> "" Then
            flag1 = 1
        End If
        If Cells(K * 4 + 1, 10) <> "" Then
            flag2 = 2
        End If
        If Cells(K * 4 + 1, 11) <> "" Then
            flag3 = 4
        End If
        If Cells(K * 4 + 1, 12) <> "" Then
            flag4 = 8
        End If
        flag = flag1 + flag2 + flag3 + flag4
        Select Case K
            Case 1:
                koef = 3
            Case 2:
                koef = 7
            Case 3:
                koef = 5
            Case 4:
                koef = 10
            Case 5:
                koef = 3
            Case 6:
                koef = 7
            Case 7:
                koef = 5
            Case 8:
                koef = 10
        End Select
        Select Case flag
            Case 1:
                'Есть значение только 1ая ось с заводскими настройками
                temp11 = Cells(K * 4 + 1, 9).Value                      '/заполняем первую строку
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 11).Value = temp11 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value                     'заполнили первую строку/
            Case 2:
                'Есть значение только 2ая ось с заводскими настройками
                '/заполняем первую строку
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 9).Value = temp12 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 11).Value = temp11 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                'заполнили первую строку/
            Case 3:
                'Есть значение только 1ая и 2ая ось с заводскими настройками
                '/заполняем первую строку
                temp11 = Cells(K * 4 + 1, 9).Value
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 11).Value = temp11 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                'заполнили первую строку/
            Case 4:
                'Есть значение только 3яя ось с заводскими настройками
                '/заполняем первую строку
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 9).Value = temp13 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                'заполнили первую строку/
            Case 5:
            'Есть значение только 1ая и 3яя ось с заводскими настройками
                '/заполняем первую строку
                temp11 = Cells(K * 4 + 1, 9).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                'заполнили первую строку/
            Case 6:
            'Есть значение только 2ая и 3яя ось с заводскими настройками
                '/заполняем первую строку
                temp12 = Cells(K * 4 + 1, 10).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 9).Value = temp12 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                'заполнили первую строку/
            Case 7:
            'Есть значение только 1ая, 2ая и 3яя ось с заводскими настройками
                '/заполняем первую строку
                temp11 = Cells(K * 4 + 1, 9).Value
                temp12 = Cells(K * 4 + 1, 10).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                'заполнили первую строку/
            Case 8:
            'Есть значение только 4ая ось с заводскими настройками
                '/заполняем первую строку
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 11).Value = temp14 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 9).Value = temp13 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                'заполнили первую строку/
            Case 9:
                'Есть значение только 1ая, 4ая ось с заводскими настройками
                '/заполняем первую строку
                temp11 = Cells(K * 4 + 1, 9).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 11).Value = temp14 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                'заполнили первую строку/
            Case 10:
                'Есть значение только 2ая и 4ая ось с заводскими настройками
                '/заполняем первую строку
                temp12 = Cells(K * 4 + 1, 10).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 9).Value = temp12 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 11).Value = temp14 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                'заполнили первую строку/
            Case 11:
                'Есть значение только 1ая, 2ая и 4ая ось с заводскими настройками
                '/заполняем первую строку
                temp11 = Cells(K * 4 + 1, 9).Value
                temp12 = Cells(K * 4 + 1, 10).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 11).Value = temp14 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                'заполнили первую строку/
            Case 12:
            'Есть значение только 3яя и 4ая ось с заводскими настройками
                '/заполняем первую строку
                temp13 = Cells(K * 4 + 1, 11).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 9).Value = temp13 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                'заполнили первую строку/
            Case 13:
            'Есть значение только 1ая, 3яя и 4ая ось с заводскими настройками
                '/заполняем первую строку
                temp11 = Cells(K * 4 + 1, 9).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                'заполнили первую строку/
            Case 14:
            'Есть значение только 2ая, 3яя и 4ая ось с заводскими настройками
                '/заполняем первую строку
                temp12 = Cells(K * 4 + 1, 10).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 9).Value = temp13 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                'заполнили первую строку/
            Case 15:
            'Есть значение все оси с заводскими настройками
                temp11 = Cells(K * 4 + 1, 9).Value
                temp12 = Cells(K * 4 + 1, 10).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                temp14 = Cells(K * 4 + 1, 12).Value
            Case 0:
            'Не заполнен сектор
            MsgBox ("Не заполнен сектор " & K)
        End Select
        If flag <> 0 Then
            '/ заполняем вторую строку
22:         Cells(K * 4 + 2, 9).Value = temp11 - 25 - Int(Rnd * koef * 4)
            Cells(K * 4 + 2, 10).Value = temp12 - 25 - Int(Rnd * koef * 4)
            If Cells(K * 4 + 2, 10).Value < Cells(K * 4 + 2, 9).Value Then GoTo 22
24:         Cells(K * 4 + 2, 11).Value = temp13 - 25 - Int(Rnd * koef * 4)
            Cells(K * 4 + 2, 12).Value = temp14 - 25 - Int(Rnd * koef * 4)
            If Cells(K * 4 + 2, 11).Value < Cells(K * 4 + 2, 12).Value Then GoTo 24
            'заполнили 2 строку/
            
            '/ заполняем 3 строку
            Cells(K * 4 + 3, 9).Value = temp11 + Int(Rnd * koef * 2.5)
            Cells(K * 4 + 3, 10).Value = temp12 + Int(Rnd * koef * 2.5)
            Cells(K * 4 + 3, 11).Value = temp13 + Int(Rnd * koef * 2.5)
            Cells(K * 4 + 3, 12).Value = temp14 + Int(Rnd * koef * 2.5)
            'заполнили 3 строку/
            '/ заполняем 4 строку
            Cells(K * 4 + 4, 9).Value = temp11 - Int(Rnd * koef * 2.5)
            Cells(K * 4 + 4, 10).Value = temp12 - Int(Rnd * koef * 2.5)
            Cells(K * 4 + 4, 11).Value = temp13 - Int(Rnd * koef * 2.5)
            Cells(K * 4 + 4, 12).Value = temp14 - Int(Rnd * koef * 2.5)
            'заполнили 4 строку/
        End If
            
    Next K
End If
End Sub

