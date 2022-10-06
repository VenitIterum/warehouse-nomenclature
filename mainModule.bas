Attribute VB_Name = "mainModule"
Option Explicit

Sub startTest()
    
    Dim strFlag As String, strFlagEnter As String, strThisSheet As String, _
    strInstaller As String, strError As String
    Dim longSheetsCount As Long, longRowCount As Long, longLastRow As Long, _
    longCountInstaller As Long
    Dim blSearchCheck As Boolean, blInstaller As Boolean
    blSearchCheck = False
    
    Dim wsMainSheet As Worksheet, wsEnterSheet As Worksheet
    Set wsMainSheet = Workbooks("������-����.xlsm").Worksheets("����")
    
    strThisSheet = ActiveSheet.Name
    'CASE ��� �������� �����
    Select Case strThisSheet
        'CASE "����"
        Case "����"
            If ActiveCell.Address = "$B$3" Then
                '����� ������
                If wsMainSheet.Range("B3") = "enter" Then
                    wsMainSheet.Range("B3") = ""
                    wsMainSheet.Range("Z6") = "enter"
                    wsMainSheet.Range("D7:F7").Interior.Color = RGB(246, 9, 0)
                    wsMainSheet.Range("D7").Interior.Color = RGB(23, 229, 3)
                    wsMainSheet.Range("B6") = "������� ����� �����"            '������� ���������
                    wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                ElseIf wsMainSheet.Range("B3") = "search" Then
                    wsMainSheet.Range("B3") = ""
                    wsMainSheet.Range("Z6") = "search"
                    wsMainSheet.Range("D7:F7").Interior.Color = RGB(246, 9, 0)
                    wsMainSheet.Range("E7").Interior.Color = RGB(23, 229, 3)
                    
                    wsMainSheet.Range("Z10") = ""
                    wsMainSheet.Range("D11:J11").Interior.Color = RGB(246, 9, 0)
                    
                    wsMainSheet.Range("B6") = "������� ����� ������"           '������� ���������
                    wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                ElseIf wsMainSheet.Range("B3") = "enter_search" Then
                    wsMainSheet.Range("B3") = ""
                    wsMainSheet.Range("Z6") = "enter_search"
                    wsMainSheet.Range("D7:F7").Interior.Color = RGB(246, 9, 0)
                    wsMainSheet.Range("F7").Interior.Color = RGB(23, 229, 3)
                    wsMainSheet.Range("B6") = "������� ����� ����� � �������"  '������� ���������
                    wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                Else
                '������ �������
                strFlag = wsMainSheet.Range("Z6")
                Select Case strFlag
                    '����� �����
                    Case "enter"
                        If wsMainSheet.Range("B3") = "unknow" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "unknow"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("D11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "����� ����� � ����" & vbCrLf & _
                            "������������"      '������� ���������
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "blocks" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "blocks"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("E11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "����� ����� � ����" & vbCrLf & _
                            "������ �����"      '������� ���������
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "dut" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "dut"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("F11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "����� ����� � ����" & vbCrLf & _
                            "������ ���"        '������� ���������
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "tachographs" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "tachographs"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("G11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "����� ����� � ����" & vbCrLf & _
                            "������ ���������"  '������� ���������
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "skzi" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "skzi"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("H11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "����� ����� � ����" & vbCrLf & _
                            "������ ����"       '������� ���������
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "heaters" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "heaters"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("I11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "����� ����� � ����" & vbCrLf & _
                            "������ ���������"  '������� ���������
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "auto" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "auto"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("J11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "����� �����" & vbCrLf & _
                            "��������������"    '������� ���������
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        Else
                            '����� �������� �������� �����!!!
                            strFlagEnter = wsMainSheet.Range("Z10")
                            If wsMainSheet.Range("B3") = "new_parish" Then
                                wsMainSheet.Range("B3") = ""
                                Select Case strFlagEnter
                                    Case "unknow"
                                        enterNewParish "������������"
                                    Case "blocks"
                                        enterNewParish "������ �����"
                                    Case "dut"
                                        enterNewParish "������ ���"
                                    Case "tachographs"
                                        enterNewParish "������ ���������"
                                    Case "skzi"
                                        enterNewParish "������ ����"
                                    Case "heaters"
                                        enterNewParish "������ ���������"
                                    Case Else
                                        wsMainSheet.Range("B6") = _
                                        "�������� ���� ��� ����� ����� �����!"
                                        wsMainSheet.Range("B6").Font.Color = _
                                        RGB(246, 9, 0)
                                End Select
                            ElseIf wsMainSheet.Range("B3") = "delete_parish" Then
                                wsMainSheet.Range("B3") = ""
                                Select Case strFlagEnter
                                    Case "unknow"
                                        deleteNewParish "������������"
                                    Case "blocks"
                                        deleteNewParish "������ �����"
                                    Case "dut"
                                        deleteNewParish "������ ���"
                                    Case "tachographs"
                                        deleteNewParish "������ ���������"
                                    Case "skzi"
                                        deleteNewParish "������ ����"
                                    Case "heaters"
                                        deleteNewParish "������ ���������"
                                    Case Else
                                        wsMainSheet.Range("B6") = _
                                        "�������� ���� ��� �������� ����� �����!"
                                        wsMainSheet.Range("B6").Font.Color = _
                                        RGB(246, 9, 0)
                                End Select
                            Else
                                Select Case strFlagEnter
                                    Case "unknow"
                                        enterNewCodes "������������"
                                    Case "blocks"
                                        enterNewCodes "������ �����"
                                    Case "dut"
                                        enterNewCodes "������ ���"
                                    Case "tachographs"
                                        enterNewCodes "������ ���������"
                                    Case "skzi"
                                        enterNewCodes "������ ����"
                                    Case "heaters"
                                        enterNewCodes "������ ���������"
                                    Case Else
                                        wsMainSheet.Range("B6") = _
                                        "�������� ���� ��� ����� ������ �������!"
                                        wsMainSheet.Range("B6").Font.Color = _
                                        RGB(246, 9, 0)
                                End Select
                            End If
                        End If
                        Debug.Print "���� ��������"
                    '����� ������
                    Case "search"
                        longSheetsCount = 3
                        Do While longSheetsCount <> Workbooks("������-����.xlsm").Worksheets.Count And blSearchCheck = False
                            longRowCount = 2
                            longLastRow = Workbooks("������-����.xlsm").Worksheets(longSheetsCount).Cells(Rows.Count, 6).End(xlUp).Row
                            Do While Workbooks("������-����.xlsm").Worksheets(longSheetsCount).Cells(longRowCount, 6) <> wsMainSheet.Range("B3") And longRowCount <> longLastRow + 1
                                longRowCount = longRowCount + 1
                            Loop
                            If Workbooks("������-����.xlsm").Worksheets(longSheetsCount).Cells(longRowCount, 6) = wsMainSheet.Range("B3") Then
                                blSearchCheck = True
                            End If
                            longSheetsCount = longSheetsCount + 1
                        Loop
                        If blSearchCheck = True Then
                            Workbooks("������-����.xlsm").Worksheets(longSheetsCount - 1).Select
                            Workbooks("������-����.xlsm").Worksheets(longSheetsCount - 1).Cells(longRowCount, 8).Select
                            wsMainSheet.Range("B6") = "������ � �����: " & Workbooks("������-����.xlsm").Worksheets(longSheetsCount - 1).Name  '������� ���������
                            wsMainSheet.Range("B6").Font.Color = _
                            RGB(23, 229, 3)
                        Else
                            wsMainSheet.Range("B6") = "������ �� �������"       '������� ���������
                            wsMainSheet.Range("B6").Font.Color = _
                            RGB(246, 9, 0)
                        End If
                        Debug.Print "����� ��������"
                    '����� ����� � �������
                    Case "enter_search"
                        wsMainSheet.Range("B6") = "���� � ������� �� ��������"  '������� ���������
                        wsMainSheet.Range("B6").Font.Color = _
                        RGB(246, 9, 0)
                        Debug.Print "���� � �������"
                End Select
                End If
            Else
                If ActiveCell <> "" Then
                    Application.Undo
                End If
                Workbooks("������-����.xlsm").Worksheets("����").Range("B3").Activate
                wsMainSheet.Range("B6").Font.Color = RGB(246, 9, 0)
                wsMainSheet.Range("B6") = "������� ��� ���!"  '������� ���������
            End If
        'CASE "������ �����", "������ ���", "������ ���������", "������ ����", "������ ���������"
        Case "������������", "������ �����", "������ ���", "������ ���������", "������ ����", "������ ���������"
            '������� ����� ��� �������?? ������� ��� � �����!!!
            If ActiveCell.Value = "main_sheet" Then
                Application.Undo
                Workbooks("������-����.xlsm").Worksheets("����").Activate
                Workbooks("������-����.xlsm").Worksheets("����").Range("B3").Activate
            '�������� ������������
            ElseIf ActiveCell Like "T##" Then
                strInstaller = ActiveCell
                longCountInstaller = 14
                Do While wsMainSheet.Range("D" & longCountInstaller - 1) <> "" And wsMainSheet.Range("D" & longCountInstaller - 1) <> strInstaller
                    If wsMainSheet.Range("D" & longCountInstaller) = strInstaller Then
                        ActiveCell = wsMainSheet.Range("E" & longCountInstaller)
                        fillInstaller
                    End If
                    longCountInstaller = longCountInstaller + 1
                Loop
            ElseIf ActiveCell = "replacement" Then
                Application.Undo
                ActiveCell.Offset(0, 1) = "���������"
            Else
                otherFuncional "������� �����������!"
                Range("H" & ActiveCell.Row).Activate
            End If
        'CASE ������ ������������
        Case "������ ������������"
            otherFuncional "���� ������ ������������ �� �������� ��������"
        'CASE ���� ��������� ���� �� �������� ��������
        Case Else
            otherFuncional "� ������� ����� ��� �������"
    End Select
End Sub

Sub fillInstaller()
    '������ �����������
    ActiveCell.Offset(0, 3) = ActiveCell.Offset(0, 1)
    ActiveCell.Offset(0, 4) = ActiveCell.Offset(0, 2)
    If ActiveCell.Offset(0, 1) <> "���������" Then
        ActiveCell.Offset(0, 1) = "����� ���-��"
    End If
    ActiveCell.Offset(0, 2) = Date
    Workbooks("������-����.xlsm").Worksheets("����").Range("B6") = _
    "����: " & Date & vbCrLf & "������: " & ActiveCell '������� ���������
    Workbooks("������-����.xlsm").Worksheets("����").Range("B6").Font.Color = _
    RGB(23, 229, 3)
    
End Sub

Sub enterNewCodes(strNameSheet As String)
    '���� ����� �����-�����
    Dim longLastRow As Long
    Dim wsEnterSheet As Worksheet
    Set wsEnterSheet = Workbooks("������-����.xlsm").Worksheets(strNameSheet)
    
    longLastRow = wsEnterSheet.Cells(Rows.Count, 6).End(xlUp).Row
    '��������, ���� ������ �����
    If wsEnterSheet.Cells(longLastRow + 1, 6).Interior.Color = RGB(255, 255, 0) Then
        longLastRow = longLastRow + 1
    End If
    wsEnterSheet.Cells(longLastRow + 1, 6) = _
    Workbooks("������-����.xlsm").Worksheets("����").Range("B3")
    '��������� �����
    If wsEnterSheet.Cells(longLastRow, 1) = "" Then
        wsEnterSheet.Cells(longLastRow + 1, 1) = 1
    Else
        wsEnterSheet.Cells(longLastRow + 1, 1) = wsEnterSheet.Cells(longLastRow, 1).Value + 1
    End If
    '   '   '
    wsEnterSheet.Cells(longLastRow + 1, 9) = _
    "�����"
    wsEnterSheet.Cells(longLastRow + 1, 10) = _
    Workbooks("������-����.xlsm").Worksheets("����").Range("B9")
    Workbooks("������-����.xlsm").Worksheets("����").Range("B6") = _
    "����-��� ������ � ����" & vbCrLf & wsEnterSheet.Name '������� ���������
    Workbooks("������-����.xlsm").Worksheets("����").Range("B6").Font.Color = _
    RGB(23, 229, 3)
    
End Sub

Sub enterNewParish(strNameSheet As String)
    '���� ����� �����
    Dim longLastRow As Long, longLastColumn As Long
    Dim wsEnterSheet As Worksheet
    Set wsEnterSheet = Workbooks("������-����.xlsm").Worksheets(strNameSheet)

    longLastColumn = wsEnterSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    longLastRow = wsEnterSheet.Cells(Rows.Count, 6).End(xlUp).Row
    If wsEnterSheet.Range("A" & longLastRow + 1 & ":" & wsEnterSheet.Cells(longLastRow + 1, _
        longLastColumn).Address).Interior.Color <> RGB(255, 255, 0) Then
        wsEnterSheet.Range("A" & longLastRow + 1 & ":" & wsEnterSheet.Cells(longLastRow + 1, _
        longLastColumn).Address).Interior.Color = RGB(255, 255, 0)
        Workbooks("������-����.xlsm").Worksheets("����").Range("B6") = _
        "Ƹ���� ������ � ����" & vbCrLf & wsEnterSheet.Name   '������� ���������
        Workbooks("������-����.xlsm").Worksheets("����").Range("B6").Font.Color = _
        RGB(23, 229, 3)
        wsEnterSheet.Cells(longLastRow + 2, 2) = _
        Workbooks("������-����.xlsm").Worksheets("����").Range("B9")
    Else
        Workbooks("������-����.xlsm").Worksheets("����").Range("B6") = _
        "Ƹ���� ����� ��� ����!"   '������� ���������
        Workbooks("������-����.xlsm").Worksheets("����").Range("B6").Font.Color = _
        RGB(246, 9, 0)
    End If
    
End Sub

Sub deleteNewParish(strNameSheet As String)
    '�������� ����� ����� (�� ��������!!!)
    Dim longLastRow As Long, longLastColumn As Long, longCountColumn As Long
    Dim wsEnterSheet As Worksheet
    Set wsEnterSheet = Workbooks("������-����.xlsm").Worksheets(strNameSheet)
    longCountColumn = 1

    longLastColumn = wsEnterSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    longLastRow = wsEnterSheet.Cells(Rows.Count, 6).End(xlUp).Row
    If wsEnterSheet.Cells(longLastRow + 1, longCountColumn).Interior.Color = RGB(255, 255, 0) Then
        Do While longCountColumn <> longLastColumn + 1
            If wsEnterSheet.Cells(longLastRow, longCountColumn).Interior.Color = RGB(0, 255, 255) Then
                wsEnterSheet.Cells(longLastRow + 1, longCountColumn).Interior.Color = RGB(0, 255, 255)
            Else
                wsEnterSheet.Cells(longLastRow + 1, longCountColumn).Interior.Color = xlNone
            End If
            longCountColumn = longCountColumn + 1
        Loop
        Workbooks("������-����.xlsm").Worksheets("����").Range("B6") = _
        "Ƹ���� ������ �� �����" & vbCrLf & wsEnterSheet.Name   '������� ���������
        Workbooks("������-����.xlsm").Worksheets("����").Range("B6").Font.Color = _
        RGB(23, 229, 3)
        wsEnterSheet.Cells(longLastRow + 2, 2) = ""
    Else
        Workbooks("������-����.xlsm").Worksheets("����").Range("B6") = _
        "Ƹ���� ����� �� �������!"   '������� ���������
        Workbooks("������-����.xlsm").Worksheets("����").Range("B6").Font.Color = _
        RGB(246, 9, 0)
    End If
    
End Sub

Sub otherFuncional(strError As String)
    '�������, ������ (�� ��������!!!), ����� ����������� ������
    If ActiveCell.Value = "main_sheet" Then
        Application.Undo
        Workbooks("������-����.xlsm").Worksheets("����").Activate
        Workbooks("������-����.xlsm").Worksheets("����").Range("B3").Activate
    '�������� �� ��� ������������. �������� ���!!!
    ElseIf ActiveCell.Value = "cancel" Then
        Application.Undo
    Else
        If ActiveCell <> "" Then
            Application.Undo
        End If
        MsgBox strError
    End If

End Sub
