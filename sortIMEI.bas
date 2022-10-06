Attribute VB_Name = "sortIMEI"
Option Explicit

Sub theSortOfImei()
'Объявление переменных
    Dim intCount As Integer             'для 3-ой и 4-ой книг (пустое ли имя?)
    Dim intCount_gt As Integer          'для 1-ей книги (пустой ли IMEI?)
    Dim intCount_gt_start As Integer    'для 1-ей книги (начинает сравнение с послелних 100 IMEI-ев)
    Dim intCount_result As Integer      'для 2-ой книги (определения конца таблицы)
    Dim strValue_1 As String            'сравнение для 3-ой и 4-ой книг
    Dim strValue_2 As String            'сравнение для 1-ей книги
    '1 > 3
    '2 > 4
    '3 > 1
    '4 > 2
    intCount = 2
    intCount_gt_start = 2
    intCount_result = 2
    
    Dim waWialon_1 As Worksheet
    Set waWialon_1 = Workbooks(3).Worksheets("Объекты")
    Dim waWialon_2 As Worksheet
    Set waWialon_2 = Workbooks(4).Worksheets("Объекты")
    Dim waGoogleTable As Worksheet
    Set waGoogleTable = Workbooks(1).Worksheets("Приход БЛОКИ")
    Dim wsResultWorksheet As Worksheet
    Set wsResultWorksheet = Workbooks("Актуализация_данных.xlsm").Worksheets("Result")
        
'Очистка таблички
    Do While wsResultWorksheet.Range("A" & intCount) <> ""
        wsResultWorksheet.Range("A" & intCount & ":F" & intCount) = ""
        wsResultWorksheet.Range("A" & intCount & ":F" & intCount).Borders.LineStyle = False
        intCount = intCount + 1
    Loop
    intCount = 2
'
'Отсчёт с конца 1-его документа
    Do While waGoogleTable.Range("F" & intCount_gt_start) <> "" Or waGoogleTable.Range("F" & intCount_gt_start + 1) <> "" Or waGoogleTable.Range("F" & intCount_gt_start + 2) <> "" Or waGoogleTable.Range("F" & intCount_gt_start + 3) <> ""
        intCount_gt_start = intCount_gt_start + 1
    Loop
    intCount_gt_start = intCount_gt_start - wsResultWorksheet.Range("G6")
'
'Сортиовка по 3-ому и 4-ому документам
    intCount_gt = intCount_gt_start
    Do While waGoogleTable.Range("F" & intCount_gt) <> "" Or waGoogleTable.Range("F" & intCount_gt + 1) <> "" Or waGoogleTable.Range("F" & intCount_gt + 2) <> "" Or waGoogleTable.Range("F" & intCount_gt + 3) <> ""
        strValue_2 = waGoogleTable.Range("F" & intCount_gt)
        If waGoogleTable.Range("F" & intCount_gt) <> "" Then
            intCount = 2
'Проверка по 3-ой книге
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
'Проверка по 4-ой книге
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
'Рисовка линий
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

    intValue_1 = Workbooks(3).Worksheets("Объекты").Range("F2335")
    intValue_2 = Workbooks(1).Worksheets("Приход БЛОКИ").Range("F2675")

    MsgBox intValue_1 & " | " & intValue_2
    MsgBox Workbooks(3).Worksheets("Объекты").Range("F2335").NumberFormatLocal & " | " & Workbooks(1).Worksheets("Приход БЛОКИ").Range("E2675").NumberFormatLocal
    If intValue_1 = intValue_2 Then
    MsgBox "IMEI равны"
    Else
    MsgBox "IMEI не равны"
    End If

End Sub

Sub msgTest2()

    Dim intCount_gt_start As Integer
    intCount_gt_start = 2

    Do While Workbooks(1).Worksheets("Приход БЛОКИ").Range("E" & intCount_gt_start) <> "" Or Workbooks(1).Worksheets("Приход БЛОКИ").Range("E" & intCount_gt_start + 1) <> "" Or Workbooks(1).Worksheets("Приход БЛОКИ").Range("E" & intCount_gt_start + 2) <> "" Or Workbooks(1).Worksheets("Приход БЛОКИ").Range("E" & intCount_gt_start + 3) <> ""
        intCount_gt_start = intCount_gt_start + 1
    Loop
    intCount_gt_start = intCount_gt_start - 100

    MsgBox intCount_gt_start
End Sub
