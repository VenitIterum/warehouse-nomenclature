Attribute VB_Name = "sortIMEI"
Option Explicit

Sub theSortOfImei()
'���������� ����������
    Dim intCount As Integer             '��� 3-�� � 4-�� ���� (������ �� ���?)
    Dim intCount_gt As Integer          '��� 1-�� ����� (������ �� IMEI?)
    Dim intCount_gt_start As Integer    '��� 1-�� ����� (�������� ��������� � ��������� 100 IMEI-��)
    Dim intCount_result As Integer      '��� 2-�� ����� (����������� ����� �������)
    Dim strValue_1 As String            '��������� ��� 3-�� � 4-�� ����
    Dim strValue_2 As String            '��������� ��� 1-�� �����
    '1 > 3
    '2 > 4
    '3 > 1
    '4 > 2
    intCount = 2
    intCount_gt_start = 2
    intCount_result = 2
    
    Dim waWialon_1 As Worksheet
    Set waWialon_1 = Workbooks(3).Worksheets("�������")
    Dim waWialon_2 As Worksheet
    Set waWialon_2 = Workbooks(4).Worksheets("�������")
    Dim waGoogleTable As Worksheet
    Set waGoogleTable = Workbooks(1).Worksheets("������ �����")
    Dim wsResultWorksheet As Worksheet
    Set wsResultWorksheet = Workbooks("������������_������.xlsm").Worksheets("Result")
        
'������� ��������
    Do While wsResultWorksheet.Range("A" & intCount) <> ""
        wsResultWorksheet.Range("A" & intCount & ":F" & intCount) = ""
        wsResultWorksheet.Range("A" & intCount & ":F" & intCount).Borders.LineStyle = False
        intCount = intCount + 1
    Loop
    intCount = 2
'
'������ � ����� 1-��� ���������
    Do While waGoogleTable.Range("F" & intCount_gt_start) <> "" Or waGoogleTable.Range("F" & intCount_gt_start + 1) <> "" Or waGoogleTable.Range("F" & intCount_gt_start + 2) <> "" Or waGoogleTable.Range("F" & intCount_gt_start + 3) <> ""
        intCount_gt_start = intCount_gt_start + 1
    Loop
    intCount_gt_start = intCount_gt_start - wsResultWorksheet.Range("G6")
'
'��������� �� 3-��� � 4-��� ����������
    intCount_gt = intCount_gt_start
    Do While waGoogleTable.Range("F" & intCount_gt) <> "" Or waGoogleTable.Range("F" & intCount_gt + 1) <> "" Or waGoogleTable.Range("F" & intCount_gt + 2) <> "" Or waGoogleTable.Range("F" & intCount_gt + 3) <> ""
        strValue_2 = waGoogleTable.Range("F" & intCount_gt)
        If waGoogleTable.Range("F" & intCount_gt) <> "" Then
            intCount = 2
'�������� �� 3-�� �����
            strValue_1 = waWialon_1.Range("F" & intCount)
            Do While waWialon_1.Range("A" & intCount) <> "" And strValue_1 <> strValue_2
                If waWialon_1.Range("F" & intCount) <> "" Then
                    strValue_1 = waWialon_1.Range("F" & intCount)
                End If
                intCount = intCount + 1
            Loop
            If strValue_1 = strValue_2 Then
                intCount_result = 2
                Do While wsResultWorksheet.Range("A" & intCount_result) <> ""
                    intCount_result = intCount_result + 1
                Loop
                wsResultWorksheet.Range("A" & intCount_result) = intCount_gt
                wsResultWorksheet.Range("B" & intCount_result) = waWialon_1.Range("A" & intCount - 1)
                wsResultWorksheet.Range("C" & intCount_result) = waWialon_1.Range("B" & intCount - 1)
                wsResultWorksheet.Range("D" & intCount_result) = Left(waWialon_1.Range("I" & intCount - 1), 10)
                wsResultWorksheet.Range("E" & intCount_result) = waWialon_1.Range("F" & intCount - 1)
                wsResultWorksheet.Range("F" & intCount_result) = 1
                
                waGoogleTable.Range("M" & intCount_gt) = waWialon_1.Range("A" & intCount - 1)
                waGoogleTable.Range("N" & intCount_gt) = waWialon_1.Range("B" & intCount - 1)
                waGoogleTable.Range("O" & intCount_gt) = Left(waWialon_1.Range("I" & intCount - 1), 10)
            End If
'�������� �� 4-�� �����
            intCount = 2
            strValue_1 = waWialon_2.Range("F" & intCount)
            Do While waWialon_2.Range("A" & intCount) <> "" And strValue_1 <> strValue_2
                If waWialon_2.Range("F" & intCount) <> "" Then
                    strValue_1 = waWialon_2.Range("F" & intCount)
                End If
                intCount = intCount + 1
            Loop
            If strValue_1 = strValue_2 Then
                intCount_result = 2
                Do While wsResultWorksheet.Range("A" & intCount_result) <> ""
                    intCount_result = intCount_result + 1
                Loop
                wsResultWorksheet.Range("A" & intCount_result) = intCount_gt
                wsResultWorksheet.Range("B" & intCount_result) = waWialon_2.Range("A" & intCount - 1)
                wsResultWorksheet.Range("C" & intCount_result) = waWialon_2.Range("B" & intCount - 1)
                wsResultWorksheet.Range("D" & intCount_result) = Left(waWialon_2.Range("I" & intCount - 1), 10)
                wsResultWorksheet.Range("E" & intCount_result) = waWialon_2.Range("F" & intCount - 1)
                wsResultWorksheet.Range("F" & intCount_result) = 3
                
                waGoogleTable.Range("M" & intCount_gt) = waWialon_2.Range("A" & intCount - 1)
                waGoogleTable.Range("N" & intCount_gt) = waWialon_2.Range("B" & intCount - 1)
                waGoogleTable.Range("O" & intCount_gt) = Left(waWialon_2.Range("I" & intCount - 1), 10)
            End If
        End If
'������� �����
        If intCount_result > 2 Then
            If wsResultWorksheet.Range("A" & intCount_result) - wsResultWorksheet.Range("A" & intCount_result - 1) > 1 Then
                With wsResultWorksheet.Range("A" & intCount_result & ":F" & intCount_result).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
            End If
        End If
        intCount_gt = intCount_gt + 1
    Loop
'
    
End Sub

Sub msgTest()

    Dim intValue_1 As String
    Dim intValue_2 As String

    intValue_1 = Workbooks(3).Worksheets("�������").Range("F2335")
    intValue_2 = Workbooks(1).Worksheets("������ �����").Range("F2675")

    MsgBox intValue_1 & " | " & intValue_2
    MsgBox Workbooks(3).Worksheets("�������").Range("F2335").NumberFormatLocal & " | " & Workbooks(1).Worksheets("������ �����").Range("E2675").NumberFormatLocal
    If intValue_1 = intValue_2 Then
    MsgBox "IMEI �����"
    Else
    MsgBox "IMEI �� �����"
    End If

End Sub

Sub msgTest2()

    Dim intCount_gt_start As Integer
    intCount_gt_start = 2

    Do While Workbooks(1).Worksheets("������ �����").Range("E" & intCount_gt_start) <> "" Or Workbooks(1).Worksheets("������ �����").Range("E" & intCount_gt_start + 1) <> "" Or Workbooks(1).Worksheets("������ �����").Range("E" & intCount_gt_start + 2) <> "" Or Workbooks(1).Worksheets("������ �����").Range("E" & intCount_gt_start + 3) <> ""
        intCount_gt_start = intCount_gt_start + 1
    Loop
    intCount_gt_start = intCount_gt_start - 100

    MsgBox intCount_gt_start
End Sub
