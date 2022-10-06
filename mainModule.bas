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
    Set wsMainSheet = Workbooks("Приход-уход.xlsm").Worksheets("Ввод")
    
    strThisSheet = ActiveSheet.Name
    'CASE для проверки листа
    Select Case strThisSheet
        'CASE "Ввод"
        Case "Ввод"
            If ActiveCell.Address = "$B$3" Then
                'Выбор режима
                If wsMainSheet.Range("B3") = "enter" Then
                    wsMainSheet.Range("B3") = ""
                    wsMainSheet.Range("Z6") = "enter"
                    wsMainSheet.Range("D7:F7").Interior.Color = RGB(246, 9, 0)
                    wsMainSheet.Range("D7").Interior.Color = RGB(23, 229, 3)
                    wsMainSheet.Range("B6") = "Включен режим ВВОДА"            'КОНСОЛЬ СОСТОЯНИЯ
                    wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                ElseIf wsMainSheet.Range("B3") = "search" Then
                    wsMainSheet.Range("B3") = ""
                    wsMainSheet.Range("Z6") = "search"
                    wsMainSheet.Range("D7:F7").Interior.Color = RGB(246, 9, 0)
                    wsMainSheet.Range("E7").Interior.Color = RGB(23, 229, 3)
                    
                    wsMainSheet.Range("Z10") = ""
                    wsMainSheet.Range("D11:J11").Interior.Color = RGB(246, 9, 0)
                    
                    wsMainSheet.Range("B6") = "Включен режим ПОИСКА"           'КОНСОЛЬ СОСТОЯНИЯ
                    wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                ElseIf wsMainSheet.Range("B3") = "enter_search" Then
                    wsMainSheet.Range("B3") = ""
                    wsMainSheet.Range("Z6") = "enter_search"
                    wsMainSheet.Range("D7:F7").Interior.Color = RGB(246, 9, 0)
                    wsMainSheet.Range("F7").Interior.Color = RGB(23, 229, 3)
                    wsMainSheet.Range("B6") = "Включен режим ВВОДА С ПОИСКОМ"  'КОНСОЛЬ СОСТОЯНИЯ
                    wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                Else
                'Работа режимов
                strFlag = wsMainSheet.Range("Z6")
                Select Case strFlag
                    'Режим ВВОДА
                    Case "enter"
                        If wsMainSheet.Range("B3") = "unknow" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "unknow"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("D11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "Режим ВВОДА в лист" & vbCrLf & _
                            "Неопознанные"      'КОНСОЛЬ СОСТОЯНИЯ
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "blocks" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "blocks"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("E11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "Режим ВВОДА в лист" & vbCrLf & _
                            "Приход БЛОКИ"      'КОНСОЛЬ СОСТОЯНИЯ
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "dut" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "dut"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("F11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "Режим ВВОДА в лист" & vbCrLf & _
                            "Приход ДУТ"        'КОНСОЛЬ СОСТОЯНИЯ
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "tachographs" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "tachographs"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("G11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "Режим ВВОДА в лист" & vbCrLf & _
                            "Приход ТАХОГРАФЫ"  'КОНСОЛЬ СОСТОЯНИЯ
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "skzi" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "skzi"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("H11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "Режим ВВОДА в лист" & vbCrLf & _
                            "Приход СКЗИ"       'КОНСОЛЬ СОСТОЯНИЯ
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "heaters" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "heaters"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("I11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "Режим ВВОДА в лист" & vbCrLf & _
                            "Приход ОТОПИТЕЛИ"  'КОНСОЛЬ СОСТОЯНИЯ
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        ElseIf wsMainSheet.Range("B3") = "auto" Then
                            wsMainSheet.Range("B3") = ""
                            wsMainSheet.Range("Z10") = "auto"
                            wsMainSheet.Range("D11:J11").Interior.Color = "0 255 0"
                            wsMainSheet.Range("J11").Interior.Color = "0 255 255"
                            wsMainSheet.Range("B6") = "Режим ВВОДА" & vbCrLf & _
                            "автоматический"    'КОНСОЛЬ СОСТОЯНИЯ
                            wsMainSheet.Range("B6").Font.Color = RGB(255, 255, 255)
                        Else
                            'Далее основной алгоритм ввода!!!
                            strFlagEnter = wsMainSheet.Range("Z10")
                            If wsMainSheet.Range("B3") = "new_parish" Then
                                wsMainSheet.Range("B3") = ""
                                Select Case strFlagEnter
                                    Case "unknow"
                                        enterNewParish "Неопознанные"
                                    Case "blocks"
                                        enterNewParish "Приход БЛОКИ"
                                    Case "dut"
                                        enterNewParish "Приход ДУТ"
                                    Case "tachographs"
                                        enterNewParish "Приход ТАХОГРАФЫ"
                                    Case "skzi"
                                        enterNewParish "Приход СКЗИ"
                                    Case "heaters"
                                        enterNewParish "Приход ОТОПИТЕЛИ"
                                    Case Else
                                        wsMainSheet.Range("B6") = _
                                        "Выберите лист для ввода жёлтой линии!"
                                        wsMainSheet.Range("B6").Font.Color = _
                                        RGB(246, 9, 0)
                                End Select
                            ElseIf wsMainSheet.Range("B3") = "delete_parish" Then
                                wsMainSheet.Range("B3") = ""
                                Select Case strFlagEnter
                                    Case "unknow"
                                        deleteNewParish "Неопознанные"
                                    Case "blocks"
                                        deleteNewParish "Приход БЛОКИ"
                                    Case "dut"
                                        deleteNewParish "Приход ДУТ"
                                    Case "tachographs"
                                        deleteNewParish "Приход ТАХОГРАФЫ"
                                    Case "skzi"
                                        deleteNewParish "Приход СКЗИ"
                                    Case "heaters"
                                        deleteNewParish "Приход ОТОПИТЕЛИ"
                                    Case Else
                                        wsMainSheet.Range("B6") = _
                                        "Выберите лист для удаления жёлтой линии!"
                                        wsMainSheet.Range("B6").Font.Color = _
                                        RGB(246, 9, 0)
                                End Select
                            Else
                                Select Case strFlagEnter
                                    Case "unknow"
                                        enterNewCodes "Неопознанные"
                                    Case "blocks"
                                        enterNewCodes "Приход БЛОКИ"
                                    Case "dut"
                                        enterNewCodes "Приход ДУТ"
                                    Case "tachographs"
                                        enterNewCodes "Приход ТАХОГРАФЫ"
                                    Case "skzi"
                                        enterNewCodes "Приход СКЗИ"
                                    Case "heaters"
                                        enterNewCodes "Приход ОТОПИТЕЛИ"
                                    Case Else
                                        wsMainSheet.Range("B6") = _
                                        "Выберите лист для ввода нового прихода!"
                                        wsMainSheet.Range("B6").Font.Color = _
                                        RGB(246, 9, 0)
                                End Select
                            End If
                        End If
                        Debug.Print "ВВОД работает"
                    'Режим ПОИСКА
                    Case "search"
                        longSheetsCount = 3
                        Do While longSheetsCount <> Workbooks("Приход-уход.xlsm").Worksheets.Count And blSearchCheck = False
                            longRowCount = 2
                            longLastRow = Workbooks("Приход-уход.xlsm").Worksheets(longSheetsCount).Cells(Rows.Count, 6).End(xlUp).Row
                            Do While Workbooks("Приход-уход.xlsm").Worksheets(longSheetsCount).Cells(longRowCount, 6) <> wsMainSheet.Range("B3") And longRowCount <> longLastRow + 1
                                longRowCount = longRowCount + 1
                            Loop
                            If Workbooks("Приход-уход.xlsm").Worksheets(longSheetsCount).Cells(longRowCount, 6) = wsMainSheet.Range("B3") Then
                                blSearchCheck = True
                            End If
                            longSheetsCount = longSheetsCount + 1
                        Loop
                        If blSearchCheck = True Then
                            Workbooks("Приход-уход.xlsm").Worksheets(longSheetsCount - 1).Select
                            Workbooks("Приход-уход.xlsm").Worksheets(longSheetsCount - 1).Cells(longRowCount, 8).Select
                            wsMainSheet.Range("B6") = "Найден в листе: " & Workbooks("Приход-уход.xlsm").Worksheets(longSheetsCount - 1).Name  'КОНСОЛЬ СОСТОЯНИЯ
                            wsMainSheet.Range("B6").Font.Color = _
                            RGB(23, 229, 3)
                        Else
                            wsMainSheet.Range("B6") = "Ничего не найдено"       'КОНСОЛЬ СОСТОЯНИЯ
                            wsMainSheet.Range("B6").Font.Color = _
                            RGB(246, 9, 0)
                        End If
                        Debug.Print "ПОИСК работает"
                    'Режим ВВОДА С ПОИСКОМ
                    Case "enter_search"
                        wsMainSheet.Range("B6") = "ВВОД С ПОИСКОМ не работает"  'КОНСОЛЬ СОСТОЯНИЯ
                        wsMainSheet.Range("B6").Font.Color = _
                        RGB(246, 9, 0)
                        Debug.Print "ВВОД С ПОИСКОМ"
                End Select
                End If
            Else
                If ActiveCell <> "" Then
                    Application.Undo
                End If
                Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B3").Activate
                wsMainSheet.Range("B6").Font.Color = RGB(246, 9, 0)
                wsMainSheet.Range("B6") = "Введите ещё раз!"  'КОНСОЛЬ СОСТОЯНИЯ
            End If
        'CASE "Приход БЛОКИ", "Приход ДУТ", "Приход ТАХОГРАФЫ", "Приход СКЗИ", "Приход ОТОПИТЕЛИ"
        Case "Неопознанные", "Приход БЛОКИ", "Приход ДУТ", "Приход ТАХОГРАФЫ", "Приход СКЗИ", "Приход ОТОПИТЕЛИ"
            'Нахрена здесь это условие?? Запихни его в конец!!!
            If ActiveCell.Value = "main_sheet" Then
                Application.Undo
                Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Activate
                Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B3").Activate
            'Проверка установщиков
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
                ActiveCell.Offset(0, 1) = "Подменный"
            Else
                otherFuncional "Укажите установщика!"
                Range("H" & ActiveCell.Row).Activate
            End If
        'CASE ремонт оборудования
        Case "РЕМОНТ ОБОРУДОВАНИЯ"
            otherFuncional "Лист РЕМОНТ ОБОРУДОВАНИЯ не содержит макросов"
        'CASE если выбранный лист не содержит макросов
        Case Else
            otherFuncional "У данного листа нет макроса"
    End Select
End Sub

Sub fillInstaller()
    'Вводим установщика
    ActiveCell.Offset(0, 3) = ActiveCell.Offset(0, 1)
    ActiveCell.Offset(0, 4) = ActiveCell.Offset(0, 2)
    If ActiveCell.Offset(0, 1) <> "Подменный" Then
        ActiveCell.Offset(0, 1) = "Выдан уст-ку"
    End If
    ActiveCell.Offset(0, 2) = Date
    Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6") = _
    "Дата: " & Date & vbCrLf & "Принял: " & ActiveCell 'КОНСОЛЬ СОСТОЯНИЯ
    Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6").Font.Color = _
    RGB(23, 229, 3)
    
End Sub

Sub enterNewCodes(strNameSheet As String)
    'Ввод новых штрих-кодов
    Dim longLastRow As Long
    Dim wsEnterSheet As Worksheet
    Set wsEnterSheet = Workbooks("Приход-уход.xlsm").Worksheets(strNameSheet)
    
    longLastRow = wsEnterSheet.Cells(Rows.Count, 6).End(xlUp).Row
    'Проверка, если строка жёлтая
    If wsEnterSheet.Cells(longLastRow + 1, 6).Interior.Color = RGB(255, 255, 0) Then
        longLastRow = longLastRow + 1
    End If
    wsEnterSheet.Cells(longLastRow + 1, 6) = _
    Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B3")
    'Нумерация строк
    If wsEnterSheet.Cells(longLastRow, 1) = "" Then
        wsEnterSheet.Cells(longLastRow + 1, 1) = 1
    Else
        wsEnterSheet.Cells(longLastRow + 1, 1) = wsEnterSheet.Cells(longLastRow, 1).Value + 1
    End If
    '   '   '
    wsEnterSheet.Cells(longLastRow + 1, 9) = _
    "Склад"
    wsEnterSheet.Cells(longLastRow + 1, 10) = _
    Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B9")
    Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6") = _
    "Штри-код вписан в лист" & vbCrLf & wsEnterSheet.Name 'КОНСОЛЬ СОСТОЯНИЯ
    Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6").Font.Color = _
    RGB(23, 229, 3)
    
End Sub

Sub enterNewParish(strNameSheet As String)
    'Ввод жёлтой линии
    Dim longLastRow As Long, longLastColumn As Long
    Dim wsEnterSheet As Worksheet
    Set wsEnterSheet = Workbooks("Приход-уход.xlsm").Worksheets(strNameSheet)

    longLastColumn = wsEnterSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    longLastRow = wsEnterSheet.Cells(Rows.Count, 6).End(xlUp).Row
    If wsEnterSheet.Range("A" & longLastRow + 1 & ":" & wsEnterSheet.Cells(longLastRow + 1, _
        longLastColumn).Address).Interior.Color <> RGB(255, 255, 0) Then
        wsEnterSheet.Range("A" & longLastRow + 1 & ":" & wsEnterSheet.Cells(longLastRow + 1, _
        longLastColumn).Address).Interior.Color = RGB(255, 255, 0)
        Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6") = _
        "Жёлтая строка в лист" & vbCrLf & wsEnterSheet.Name   'КОНСОЛЬ СОСТОЯНИЯ
        Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6").Font.Color = _
        RGB(23, 229, 3)
        wsEnterSheet.Cells(longLastRow + 2, 2) = _
        Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B9")
    Else
        Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6") = _
        "Жёлтая линия уже есть!"   'КОНСОЛЬ СОСТОЯНИЯ
        Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6").Font.Color = _
        RGB(246, 9, 0)
    End If
    
End Sub

Sub deleteNewParish(strNameSheet As String)
    'Удаление жёлтой линии (не работает!!!)
    Dim longLastRow As Long, longLastColumn As Long, longCountColumn As Long
    Dim wsEnterSheet As Worksheet
    Set wsEnterSheet = Workbooks("Приход-уход.xlsm").Worksheets(strNameSheet)
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
        Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6") = _
        "Жёлтая строка из листа" & vbCrLf & wsEnterSheet.Name   'КОНСОЛЬ СОСТОЯНИЯ
        Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6").Font.Color = _
        RGB(23, 229, 3)
        wsEnterSheet.Cells(longLastRow + 2, 2) = ""
    Else
        Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6") = _
        "Жёлтая линия не найдена!"   'КОНСОЛЬ СОСТОЯНИЯ
        Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B6").Font.Color = _
        RGB(246, 9, 0)
    End If
    
End Sub

Sub otherFuncional(strError As String)
    'Возврат, отмена (не работает!!!), вывод критической ошибки
    If ActiveCell.Value = "main_sheet" Then
        Application.Undo
        Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Activate
        Workbooks("Приход-уход.xlsm").Worksheets("Ввод").Range("B3").Activate
    'Работает не как задумывалось. Подумать ещё!!!
    ElseIf ActiveCell.Value = "cancel" Then
        Application.Undo
    Else
        If ActiveCell <> "" Then
            Application.Undo
        End If
        MsgBox strError
    End If

End Sub
