Attribute VB_Name = "Module1"
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Public Sub FastWB(Optional ByVal opt As Boolean = True)
    With Application
        .Calculation = IIf(opt, xlCalculationManual, xlCalculationAutomatic)
        .DisplayAlerts = Not opt
        .DisplayStatusBar = Not opt
        .EnableAnimations = Not opt
        .EnableEvents = Not opt
        .ScreenUpdating = Not opt
    End With
    FastWS , opt
End Sub

Sub ПрайсПрестонДляЗагрузкиВ1С()

FastWB True


Workbooks.Open Filename:= _
        "D:\Obmen\mnglog_01\Прайс.xlsx"
Set bk1 = Workbooks("Прайс.xlsx")
bk1.Activate

Dim myList As Object
Set myList = CreateObject("Scripting.Dictionary")

botrow2 = Cells(Rows.Count, 1).End(xlUp).Row

'добавляем в массив данные из листа2, столбец1 (данные о адресах ячеек пикинга)
For x = 8 To botrow2
    On Error Resume Next
    myList.Add Cells(x, 1).Value, Cells(x, 14).Value
Next

'Проверяем на наличие элементов, если есть, в счетчик элемента добавляем 1, выводим в столбец 6 счетчик, если нет - о
'
'For x = 2 To BotRowOld - 1 'использовать botrowold
'If myList.exists(Cells(x, 7).Text) Then
'myList(Cells(x, 7).Text) = myList(Cells(x, 7).Text) + 1
'Cells(x, 6).Value = myList(Cells(x, 7).Text)
'Else
'Cells(x, 6).Value = 0
'End If
'Next
'End If
'



Workbooks.Open Filename:= _
        "D:\Obmen\mnglog_01\Цены на престон.xlsx"
Set bk2 = Workbooks("Цены на престон.xlsx")
bk2.Activate

BotRow = Cells(Rows.Count, 9).End(xlUp).Row

For x = 2 To BotRow
Cells(x, 8).Value = myList(Cells(x, 7).Value)
Cells(x, 10).Value = CDbl(FormatNumber(Cells(x, 8).Value * Cells(x, 9).Value, 2))

Next

FastWB False
Debug.Print Timer - T & " seconds"

End Sub

Sub ПрайсДляОтправкиПрестону()
'Send prices to "Preston" customer

Range("A1").CurrentRegion.Copy
Workbooks.Add
Range("A1").PasteSpecial xlPasteValues
BotRow = Cells(Rows.Count, 1).End(xlUp).Row

Range("A2:A" & BotRow).NumberFormat = "0"
Range("C:I").EntireColumn.Delete
Range("C2:C" & BotRow).NumberFormat = "0.00"
Cells(1, 1).Select
Cells(1, 1).Copy
Cells(1, 3).Value = "Цена без НДС"
Columns(1).ColumnWidth = 15
Columns(2).ColumnWidth = 55
Columns(3).ColumnWidth = 15

ActiveWorkbook.SaveAs Filename:="D:\Obmen\mnglog_01\" & Date & " " & " прайс" & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False


End Sub
Sub Хайна()
'
' Хайна Создание реализаций

    Dim ilastRow, ilastCol As Integer
    ilastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ilastCol = Cells(2, Columns.Count).End(xlToLeft).Column
    Range(Cells(1, 1), Cells(ilastRow, ilastCol)).Copy
    Workbooks.Add
    Columns("B:B").ColumnWidth = 74
    ActiveSheet.Paste
    Range(Cells(1, 1), Cells(ilastRow, ilastCol)).PasteSpecial Paste:=xlPasteValues
    Range(Cells(1, 1), Cells(ilastRow, ilastCol)).NumberFormat = "General"
    Rows(1).EntireRow.Delete
    Dim iCount As Integer
    For iCount = ilastCol To 3 Step -1
        If Application.Sum(Range(Cells(2, iCount), Cells(ilastRow, iCount))) = 0 Then
        Columns(iCount).Delete
        End If
     Next
    Dim iCount1 As Integer
    For iCount1 = ilastRow To 2 Step -1
        If Application.Sum(Range(Cells(iCount1, 3), Cells(iCount1, ilastCol))) = 0 Then
        Rows(iCount1).EntireRow.Delete
        End If
        Next
    Columns("A:A").ColumnWidth = 15
    Range(Cells(2, 1), Cells(ilastRow, 1)).NumberFormat = "0"
    Static int1 As Integer
    int1 = int1 + 1
    ChDir "D:\Obmen\mnglog_01"
    ActiveWorkbook.SaveAs Filename:="D:\Obmen\mnglog_01\Расход" & int1 & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
        
        
        
End Sub


Sub ПриходХайны()
    'верхнюю границу таблицы (ряд) задаем как 17 (переменная iTopRow)
    'правая граница таблицы (столбец) задаем как 35 (переменная iRightCol)
    'нижнюю границу определяем как последнее значение в первом столбце минус 10 (переменная iBotRow)
    Dim iTopRow, iRightCol, iBotRow, iBotRow1, ibotrow2 As Integer
    iTopRow = 17
    iRightCol = 35
    iBotRow = Cells(Rows.Count, 1).End(xlUp).Row - 10
    Range(Cells(iTopRow, 1), Cells(iBotRow, iRightCol)).Copy
    Workbooks.Add
    ActiveSheet.Paste
    iBotRow1 = Cells(Rows.Count, 2).End(xlUp).Row
    Range(Cells(1, 1), Cells(iBotRow1, iRightCol)).UnMerge
    Range(Cells(1, 3), Cells(iBotRow1, 3)).Value = "=RIGHT(RC[-1],13)"
    ActiveSheet.Usedrange.Value = ActiveSheet.Usedrange.Value
    Range("A1:AD1").AutoFilter
    ActiveWorkbook.Worksheets("Лист1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Лист1").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Лист1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ibotrow2 = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(ibotrow2 + 1, 1), Cells(iBotRow1, 1)).EntireRow.Delete
    Rows(1).EntireRow.Delete
    Range("A1:B1,D1:L1,N1,P1:AI1").EntireColumn.Delete
    

End Sub





Function Format1(iRow1, iColl1, iRow2, iColl2 As Integer, sFormula As String)
     Range(Cells(iRow1, iColl1), Cells(iRow2, iColl2)).FormatConditions.Add Type:=xlExpression, Formula1:=sFormula
     Range(Cells(iRow1, iColl1), Cells(iRow2, iColl2)).FormatConditions(Range(Cells(iRow1, iColl1), Cells(iRow2, iColl2)).FormatConditions.Count).SetFirstPriority
    With Range(Cells(iRow1, iColl1), Cells(iRow2, iColl2)).FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
    End With
    Range(Cells(iRow1, iColl1), Cells(iRow2, iColl2)).FormatConditions(1).StopIfTrue = True

End Function


Sub ПроверкаЛогТаблиц0918()
'
' Проверка Логистических таблиц
'

     Cells.FormatConditions.Delete
     
'Находим последнюю строку в 1ом столбце
     
    Dim ilastRow As Integer
    ilastRow = Cells(Rows.Count, 1).End(xlUp).Row
 'Проверяем количество символов в первом стоблце (штрих-код) должно быть 13
    
    Call Format1(3, 1, ilastRow, 1, "=ДЛСТР(RC)<>13")
 
    
'Проверяем заполненность веса брутто шт. и размеров шт. (в полях должны быть значения),

    Call Format1(3, 5, ilastRow, 8, "=ЕПУСТО(RC)")
    
'Проверяем заполненность внутренних упаково, если заполнены то количество товара во внутренней должно быть кратно количеству во внешней
'(столбец Кол_тов_внутр_уп W2:W60   , условие:  =ОСТАТ(W2;X2)=0

    Call Format1(3, 9, ilastRow, 9, "=И(НЕ(ЕПУСТО(RC));ОСТАТ(RC[6];RC)<>0)")
    
'Проверяем вес внутренней упаковки, если она есть ее вес должен быть больше или равен весу брутто шт*кол.во в внутр упаковке.
'(столбцы Масса_брутто_внутр_уп AF2:AF60    , условие: =ЕСЛИ($W2>0;ЕСЛИ(AF2>=W2*N2;ЛОЖЬ;ИСТИНА);ЛОЖЬ)

    Call Format1(3, 11, ilastRow, 11, "=ЕСЛИ(RC9>0;ЕСЛИ(RC>=RC[-6]*RC[-2];ЛОЖЬ;ИСТИНА);ЛОЖЬ)")
    
'Проверяем заполненость размеров внутренней упаковки (если заполнена информация о количестве товара в_
'внутренней упаковке, то должны быть заполнены размеры (произведение размеров шт и их кол-во в внутренней упаковке
'должно быть меньше произведения размеров внутренней упаковки)

    Call Format1(3, 12, ilastRow, 14, "=ЕСЛИ(RC9>0;R[0]C6*R[0]C7*R[0]C8*R[0]C9>R[0]C12*R[0]C13*R[0]C14;ЛОЖЬ)")


'Проверяем внешние упаковки (их количество дожно быть больше чем внутренних), если меньше либо равно, ошибка
'(столбцы Кол_тов_внеш_уп X2:X60, условие: =ЕСЛИ(W2>0;ЕСЛИ(X2<=W2;ИСТИНА;ЛОЖЬ);ЛОЖЬ)
    
    Call Format1(3, 15, ilastRow, 15, "=ЕСЛИ(ИЛИ(RC[-6]>0;RC<1);ЕСЛИ(ИЛИ(RC<=RC[-6];RC<1);ИСТИНА;ЛОЖЬ);ЛОЖЬ)")
    
'Проверяем количество внешних упаковок на паллете, ОБъем внешних упаковок на паллете должен соответствовать высоте паллета. Допускается расхождение?
'(столбцы Кол_внеш_уп_палет Z2:Z60, условие: =ОСТАТ(Z2/Y2;1)>0
    
    'Call Format1(3, 16, ilastRow, 16, "=ИЛИ(ЕПУСТО(RC);ЕСЛИОШИБКА(ОСТАТ(RC/RC[1];1)>0;ИСТИНА))")   'проверка по ряду (устаревшая)
    Call Format1(3, 16, ilastRow, 16, "=ИЛИ(ЕПУСТО(RC);RC*RC[4]*RC[5]*RC[6]/1000000000<(1,2*0,8*RC[11]/1000)*0,8;RC*RC[4]*RC[5]*RC[6]/1000000000>1,2*0,8*RC[11]/1000)")

'выводим подсказку в виде комментария что не так.

For x = 3 To ilastRow
    'определяем объем коробок на паллете в куб.м.
    b = WorksheetFunction.Product(Range(Cells(x, 20), Cells(x, 22))) * Cells(x, 16).Value / 1000000000
    'определяем объем паллеты
    g = (Cells(x, 27).Value * 1.2 * 0.8) / 1000
    g1 = ((Cells(x, 27).Value * 1.2 * 0.8) / 1000) * 0.8
    If b > g Then
    On Error Resume Next
    Range("P" & x).AddComment
    Range("P" & x).Comment.Text Text:="Объем коробок на паллете равен " & b & " куб.м. и он больше объема паллеты " & g & " куб.м. расчитанного согласно указанной высоты паллеты"
  
    ElseIf b < g1 Then
    
    On Error Resume Next
    Range("P" & x).AddComment
    Range("P" & x).Comment.Text Text:="Объем коробок на паллете равен " & b & " куб.м. и он меньше объема паллеты " & g1 & " куб.м. уменьшенного на 20% расчитанного согласно указанной высоты паллеты"
    
    Else
    On Error Resume Next
    Range("P" & x).Comment.Delete 'удаляем комменты
    
    End If
    Next

'Проверяем заполнение кол-ва внеш упаковок в ряду.
    
    Call Format1(3, 17, ilastRow, 17, "=ЕПУСТО(RC)")



'Проверяем массу бруто внешней упаковки (должны быть большее или равна вес.брутто шт*кол.во шт.)
'(столбцы Масса_бруто_внеш_тр_уп AK2:AK60   , условие: =AK2<N2*X2

    Call Format1(3, 19, ilastRow, 19, "=RC<RC[-14]*RC[-4]")


'Проверяем размеры внешней упаковки (произведение размеров шт и их кол-во в внешней упаковке
'должно быть меньше произведения размеров внутренней упаковки)

    Call Format1(3, 20, ilastRow, 22, "=R[0]C20*R[0]C21*R[0]C22<R[0]C6*R[0]C7*R[0]C8*R[0]C15")

'выводим подсказку в виде комментария что не так.

For x = 3 To ilastRow
    'определяем объем коробки в куб.м.
    b = WorksheetFunction.Product(Range(Cells(x, 20), Cells(x, 22))) / 1000000000
    'определяем объем штук в коробке куб.м.
    g = WorksheetFunction.Product(Range(Cells(x, 6), Cells(x, 8))) * Cells(x, 15).Value / 1000000000
    If g > b Then
    On Error Resume Next
    Range("T" & x).AddComment
    Range("T" & x).Comment.Text Text:="Объем штук товара в коробке равен " & g & " куб.м. и он больше объема самой коробки (" & b & ") куб.м."
    Else
    On Error Resume Next
    Range("T" & x).Comment.Delete 'удаляем комменты
  
    End If
    Next
    

'Проверяем вес паллета, должен быть больше или равен весу кол-во упаковок*вес упаковки
'(столбцы Высота_товар+палета AN2:AN60   , условие: =AN2<Z2*AK2, при этом меньше 1050

    Call Format1(3, 24, ilastRow, 24, "=ИЛИ(ЕПУСТО(RC);RC<RC[-5]*RC[-8];RC>1050)")
    



'Проверяем высоту паллета - должна быть больше 500 и меньше или равна 2200
'(столбцы Высота_товар+палета AL2:AL60   , условие: =ЕСЛИ(AL>0;ИЛИ(AL<210;AL2>1800);ЛОЖЬ)
'Так же необходима проверка на объем паллеты и количества коробов на паллете...

    Call Format1(3, 27, ilastRow, 27, "=ИЛИ(RC<500;RC>2200)")
    

    


End Sub

Sub ОтправкаПисемПоЛогТаблицам()

    str1 = InputBox("Введите название таблицы")
    ActiveWorkbook.SaveAs Filename:="T:\06 ОТДЕЛ ЛОГИСТИКИ\Общая\Логистические данные\Логистические данные_МАКСИМ\Новые\" & Date & " " & str1 & " лог таблица" & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.SendMail Array("a.sachuk@sosedi.by", "v.nesterava@sosedi.by", "splog_1@sosedi.by", "splog_2@sosedi.by", "splog_3@sosedi.by"), str1 & " лог. таблица"




    
End Sub


    


Sub ПриходХайныНовые()
    

    'правая граница таблицы (столбец) задаем как 15 (переменная iRightCol)
    'нижнюю границу определяем как последнее значение в первом столбце минус 16 (переменная iBotRow)
    Application.ScreenUpdating = False
    Dim iTopRow, iRightCol, iBotRow, iBotRow1, ibotrow2 As Integer
    
    Set rTopCell = Range("A1:Y192").Find("I ТОВАРНЫЙ РАЗДЕЛ", , , , , xlPrevious).Offset(3, 0) 'находим Товарный раздел (поиск снизу вверх) и смещаем ячейку на 3 вниз.
    'iTopRow = Mid(x.Address, InStr(2, x.Address, "$") + 1) + 3 'не используем, решение с оффсет более корректное. из найденно range вычленяем значение ряда, присваиваем его переменной TopRow (+3) чтобы попасть на начало таблицы с товаром
    
    iRightCol = 25
    

    iBotRow = Cells(Rows.Count, 1).End(xlUp).Row - 16
    Range(rTopCell, Cells(iBotRow, iRightCol)).Copy
    Workbooks.Add
    ActiveSheet.Paste
    iBotRow1 = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(iBotRow1, iRightCol)).UnMerge
    Dim iCount As Integer
    
    For iCount = (iBotRow1 - 2) To 1 Step -3
        Cells(iCount, 2) = Mid(Cells(iCount + 2, 1), 12, 12) ' вписываем штрих-код товара в 2ой столбец (данные берем из строчки на 2 строки ниже)
    
    'вызываем страшуню функцию которая должна нам разрезать сертификаты и поместить их в строку iCount, столбцы 8 9 10 11
    Call Slice1(Cells(iCount + 1, 1), iCount, 8, 9, 10, 11)
                    
              
    Next
    'немного фильтров
    ActiveSheet.Usedrange.Value = ActiveSheet.Usedrange.Value
    Range("A1:Y1").AutoFilter
    ActiveWorkbook.Worksheets("Лист1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Лист1").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Лист1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1:Y1").AutoFilter 'снимаем фильтр
    'удаляем все лишнее
    ibotrow2 = Cells(Rows.Count, 4).End(xlUp).Row
    Range(Cells(ibotrow2 + 1, 1), Cells(iBotRow1, 1)).EntireRow.Delete
    Range("C1:D1,F1,L1:Y1").EntireColumn.Delete
    
    Application.ScreenUpdating = True

    'сохраняем файл
    Static int1 As Integer
    int1 = int1 + 1
    ChDir "D:\Obmen\mnglog_01"
    ActiveWorkbook.SaveAs Filename:="D:\Obmen\mnglog_01\Приход" & int1 & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False



End Sub

Function Slice1(OneString As String, RowIndex, ColInd1, ColInd2, ColInd3, ColInd4 As Integer)

Dim String1, String2, String3, String4 As String

Dim x, y, z, z1, z2, z3 As Integer

y = InStr(OneString, "№") + 1 'ишем символ № в тексте
If y = 1 Then
y = InStr(OneString, "N") + 1 'если не нашли символ № ищем N
End If
If y <> 1 Then 'если нашли какой либо из симоволов начинаем разрезать строку
x = InStrRev(OneString, ".", y) + 1 'ищем первую точку в тексте, в обратную строну от символа № или N
String1 = Mid(OneString, x, y - x - 1) 'Отрезаем первый кусок строки (по идее это название сертификата)

' ниже попытки найти где заканчивается номер и начинается дата сертификата
z = InStr(y, OneString, "с")
z1 = InStr(y, OneString, "от")

If z = 0 Then
    If z1 > z Then
        z2 = z1
        z3 = 15 'длинна строки при условии что будем "резать" начиная с символов "от"
    Else
        z2 = 0
    End If
ElseIf z1 = 0 Then
z2 = z
z3 = 31 'длинна строки при условии что будем "резать" начиная с символа "с"
ElseIf z > z1 Then
z2 = z1
z3 = 15 'длинна строки при условии что будем "резать" начиная с символов "от"
Else
z2 = z
z3 = 31 'длинна строки при условии что будем "резать" начиная с символа "с"
End If

' мрачное решение, если есть и "с" и "от" то будет выбрано наименьшее число (номер символа)
If z2 <> 0 Then
String2 = Mid(OneString, y, z2 - y)
String3 = Mid(OneString, z2, z3)
String4 = Mid(OneString, z2 + z3, Len(OneString))

Else
String2 = Mid(OneString, y, Len(OneString))
String3 = ""
String4 = ""

End If

Else

String1 = ""
String2 = ""
String3 = ""
String4 = ""

End If

Cells(RowIndex, ColInd1).Value = String1
Cells(RowIndex, ColInd2).Value = String2
Cells(RowIndex, ColInd3).Value = String3
Cells(RowIndex, ColInd4).Value = String4


End Function




Sub Хайна2()


'
' Хайна Создание реализаций

Dim ilastRow, ilastCol As Integer
ilastRow = Cells(Rows.Count, 1).End(xlUp).Row
ilastCol = Cells(1, Columns.Count).End(xlToLeft).Column
Range(Cells(1, 1), Cells(ilastRow, ilastCol)).Copy
Workbooks.Add
ActiveSheet.Paste
Range(Cells(1, 1), Cells(ilastRow, ilastCol)).PasteSpecial xlPasteValues
Range(Cells(1, 1), Cells(ilastRow, ilastCol)).NumberFormat = "General"
Range(Cells(2, 1), Cells(ilastRow, 1)).NumberFormat = "0"
Columns("B:B").ColumnWidth = 74

    Dim iCount As Integer
    For iCount = ilastCol To 3 Step -1
        If Application.Sum(Range(Cells(2, iCount), Cells(ilastRow, iCount))) = 0 Then
        Columns(iCount).Delete
        End If
    Next
    Dim iCount1 As Integer
    For iCount1 = ilastRow To 2 Step -1
        If Application.Sum(Range(Cells(iCount1, 3), Cells(iCount1, ilastCol))) = 0 And Not VarType(Cells(iCount1, 3)) = vbString Then
        Rows(iCount1).EntireRow.Delete
        End If
        Next
Columns("A:A").ColumnWidth = 15
Range(Cells(2, 1), Cells(ilastRow, 1)).NumberFormat = "0"
    
    Static int1 As Integer
    int1 = int1 + 1
    ChDir "D:\Obmen\mnglog_01"
    ActiveWorkbook.SaveAs Filename:="D:\Obmen\mnglog_01\Расход" & int1 & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
        
End Sub

Sub ПриходРошен()
    'верхнюю границу таблицы (ряд) задаем как 71 (переменная iTopRow)
    'правая граница таблицы (столбец) задаем как 15 (переменная iRightCol)
    'нижнюю границу определяем как последнее значение в первом столбце минус 16 (переменная iBotRow)
    Application.ScreenUpdating = False
    Dim iTopRow, iRightCol, iBotRow, iBotRow1, ibotrow2, b As Integer
    iTopRow = 19
    iRightCol = 38
    iBotRow = Cells(Rows.Count, 2).End(xlUp).Row - 36
    Range(Cells(iTopRow, 1), Cells(iBotRow, iRightCol)).Copy
    Workbooks.Add
    ActiveSheet.Paste
    iBotRow1 = Cells(Rows.Count, 2).End(xlUp).Row
    Range(Cells(1, 1), Cells(iBotRow1, iRightCol)).UnMerge
    ActiveSheet.Usedrange.Value = ActiveSheet.Usedrange.Value
    b = 1
    For x = 1 To iBotRow1
    c = Cells(x, 2).Value
    If IsNumeric(c) And c > 1 Then
    Cells(x, 1).Value = b
    b = b + 1
    End If
    Next

    Range("A1:AK1").AutoFilter
    ActiveWorkbook.Worksheets("Лист1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Лист1").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Лист1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Rows(1).EntireRow.Delete
    
    iBotRow1 = Cells(Rows.Count, 1).End(xlUp).Row
    ibotrow2 = Cells(Rows.Count, 2).End(xlUp).Row
    Range(Cells(iBotRow1 + 1, 1), Cells(ibotrow2, 1)).EntireRow.Delete
    Range("C1:M1,O1,Q1:AL1").EntireColumn.Delete
   
    For Each Cell In Range(Cells(1, 3), Cells(iBotRow1, 3))
    Cell.Value = CDbl(Cell.Value)
    Next
    Application.ScreenUpdating = True

    Static int1 As Integer
    int1 = int1 + 1
    ChDir "D:\Obmen\mnglog_01"
    ActiveWorkbook.SaveAs Filename:="D:\Obmen\mnglog_01\Приход" & int1 & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False


End Sub

Sub test1()
Set x = Range("A1:Y192").Find("I ТОВАРНЫЙ РАЗДЕЛ", , , , , xlPrevious).Offset(3, 0)
MsgBox x.Address

'b = Mid(x.Address, InStr(2, x.Address, "$") + 1) + 3
iRightCol = 25
    

iBotRow = Cells(Rows.Count, 1).End(xlUp).Row - 16
Range(x, Cells(iBotRow, iRightCol)).Copy


End Sub

Sub tes3()
'
'x = 3
'b = WorksheetFunction.Product(Range(Cells(x, 20), Cells(x, 22))) * Cells(x, 15).Value / 1000000000
'
'
'MsgBox b
'
x = 4
b = 10
g = 15


Range("T" & x).AddComment

Range("T" & x).Comment.Text Text:="Объем коробки равен " & b & " При этом объем товара в коробке равен " & g


End Sub
Sub test214124()
g = Cells(3, 27).Value * 1.2
MsgBox g

End Sub


Sub testing_insert()

Dim ReturnValue, i
ReturnValue = Shell("CALC.EXE")
Application.Wait (Now + TimeValue("0:00:01"))
AppActivate "Калькулятор"
For i = 1 To 100 ' ???????? 100 ???
SendKeys i & "{+}", True ' ????????? ??????? ?? ??????? ? ???????????

Next i ' ????????? ? ???????? ? ???????????? ?????????? I

SendKeys "=", True ' ????????? ??????? ?? ???? ?????????
SendKeys "%{F4}", True ' ????????? Alt+F4 ??? ???????? ????????????
End Sub
