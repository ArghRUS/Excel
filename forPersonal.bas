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
    'Worksheets("��� ������� � ������").Activate
    dot1 = Cells(i, "a").Value
    'Worksheets("��� ������� �������").Activate
    For j = 11 To 131
        If Cells(j, "o").Value Like dot1 Then
            Range(Cells(j, "p"), Cells(j, "p")).Select
            Selection.Copy
            'Worksheets("��� ������� � ������").Activate
            Cells(i, "c").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Exit For
        End If
    Next j
Next i
'Worksheets("��� ������� � ������").Activate

End Sub
Sub brakes()
    YesNo = MsgBox("������ ��������� �� ��, ��� ������ � ������ ������ ������� ���� ���������� ������� ��������. ����� ������� ����� �����! ����������?", vbYesNo)
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
                '���� �������� ������ 1�� ��� � ���������� �����������
                temp11 = Cells(K * 4 + 1, 9).Value                      '/��������� ������ ������
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 11).Value = temp11 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value                     '��������� ������ ������/
            Case 2:
                '���� �������� ������ 2�� ��� � ���������� �����������
                '/��������� ������ ������
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 9).Value = temp12 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 11).Value = temp11 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                '��������� ������ ������/
            Case 3:
                '���� �������� ������ 1�� � 2�� ��� � ���������� �����������
                '/��������� ������ ������
                temp11 = Cells(K * 4 + 1, 9).Value
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 11).Value = temp11 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                '��������� ������ ������/
            Case 4:
                '���� �������� ������ 3�� ��� � ���������� �����������
                '/��������� ������ ������
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 9).Value = temp13 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                '��������� ������ ������/
            Case 5:
            '���� �������� ������ 1�� � 3�� ��� � ���������� �����������
                '/��������� ������ ������
                temp11 = Cells(K * 4 + 1, 9).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                '��������� ������ ������/
            Case 6:
            '���� �������� ������ 2�� � 3�� ��� � ���������� �����������
                '/��������� ������ ������
                temp12 = Cells(K * 4 + 1, 10).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 9).Value = temp12 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                '��������� ������ ������/
            Case 7:
            '���� �������� ������ 1��, 2�� � 3�� ��� � ���������� �����������
                '/��������� ������ ������
                temp11 = Cells(K * 4 + 1, 9).Value
                temp12 = Cells(K * 4 + 1, 10).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 12).Value = temp13 - 10 - Int(Rnd * koef)
                temp14 = Cells(K * 4 + 1, 12).Value
                '��������� ������ ������/
            Case 8:
            '���� �������� ������ 4�� ��� � ���������� �����������
                '/��������� ������ ������
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 11).Value = temp14 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                Cells(K * 4 + 1, 9).Value = temp13 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                '��������� ������ ������/
            Case 9:
                '���� �������� ������ 1��, 4�� ��� � ���������� �����������
                '/��������� ������ ������
                temp11 = Cells(K * 4 + 1, 9).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                Cells(K * 4 + 1, 11).Value = temp14 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                '��������� ������ ������/
            Case 10:
                '���� �������� ������ 2�� � 4�� ��� � ���������� �����������
                '/��������� ������ ������
                temp12 = Cells(K * 4 + 1, 10).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 9).Value = temp12 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 11).Value = temp14 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                '��������� ������ ������/
            Case 11:
                '���� �������� ������ 1��, 2�� � 4�� ��� � ���������� �����������
                '/��������� ������ ������
                temp11 = Cells(K * 4 + 1, 9).Value
                temp12 = Cells(K * 4 + 1, 10).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 11).Value = temp14 + 10 + Int(Rnd * koef)
                temp13 = Cells(K * 4 + 1, 11).Value
                '��������� ������ ������/
            Case 12:
            '���� �������� ������ 3�� � 4�� ��� � ���������� �����������
                '/��������� ������ ������
                temp13 = Cells(K * 4 + 1, 11).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 9).Value = temp13 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                '��������� ������ ������/
            Case 13:
            '���� �������� ������ 1��, 3�� � 4�� ��� � ���������� �����������
                '/��������� ������ ������
                temp11 = Cells(K * 4 + 1, 9).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 10).Value = temp11 + 10 + Int(Rnd * koef)
                temp12 = Cells(K * 4 + 1, 10).Value
                '��������� ������ ������/
            Case 14:
            '���� �������� ������ 2��, 3�� � 4�� ��� � ���������� �����������
                '/��������� ������ ������
                temp12 = Cells(K * 4 + 1, 10).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                temp14 = Cells(K * 4 + 1, 12).Value
                Cells(K * 4 + 1, 9).Value = temp13 - 10 - Int(Rnd * koef)
                temp11 = Cells(K * 4 + 1, 9).Value
                '��������� ������ ������/
            Case 15:
            '���� �������� ��� ��� � ���������� �����������
                temp11 = Cells(K * 4 + 1, 9).Value
                temp12 = Cells(K * 4 + 1, 10).Value
                temp13 = Cells(K * 4 + 1, 11).Value
                temp14 = Cells(K * 4 + 1, 12).Value
            Case 0:
            '�� �������� ������
            MsgBox ("�� �������� ������ " & K)
        End Select
        If flag <> 0 Then
            '/ ��������� ������ ������
22:         Cells(K * 4 + 2, 9).Value = temp11 - 25 - Int(Rnd * koef * 4)
            Cells(K * 4 + 2, 10).Value = temp12 - 25 - Int(Rnd * koef * 4)
            If Cells(K * 4 + 2, 10).Value < Cells(K * 4 + 2, 9).Value Then GoTo 22
24:         Cells(K * 4 + 2, 11).Value = temp13 - 25 - Int(Rnd * koef * 4)
            Cells(K * 4 + 2, 12).Value = temp14 - 25 - Int(Rnd * koef * 4)
            If Cells(K * 4 + 2, 11).Value < Cells(K * 4 + 2, 12).Value Then GoTo 24
            '��������� 2 ������/
            
            '/ ��������� 3 ������
            Cells(K * 4 + 3, 9).Value = temp11 + Int(Rnd * koef * 2.5)
            Cells(K * 4 + 3, 10).Value = temp12 + Int(Rnd * koef * 2.5)
            Cells(K * 4 + 3, 11).Value = temp13 + Int(Rnd * koef * 2.5)
            Cells(K * 4 + 3, 12).Value = temp14 + Int(Rnd * koef * 2.5)
            '��������� 3 ������/
            '/ ��������� 4 ������
            Cells(K * 4 + 4, 9).Value = temp11 - Int(Rnd * koef * 2.5)
            Cells(K * 4 + 4, 10).Value = temp12 - Int(Rnd * koef * 2.5)
            Cells(K * 4 + 4, 11).Value = temp13 - Int(Rnd * koef * 2.5)
            Cells(K * 4 + 4, 12).Value = temp14 - Int(Rnd * koef * 2.5)
            '��������� 4 ������/
        End If
            
    Next K
End If
End Sub

