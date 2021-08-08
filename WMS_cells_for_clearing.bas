Attribute VB_Name = "Module3"
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Function DeleteRows(FieldNum As Integer, Criteria As String)
' take num of field (field used for filtering), criteria for filtering
' returns only those part of table that meet criteria, other part of table will be droped
'�� ����� ��������� ����� ���� (�� �������� ����������� ������, �������� ��� ����������
'�� ������ ��������� ������ �� ����� ������� ������� ������������ ������� (��������� ����� �������)


Dim wsName As String, T As Double, oldusedrng As Range

Set oldWs = Worksheets(1)
wsName = oldWs.Name

Set oldusedrng = oldWs.Range("A1", GetMaxCell(oldWs.Usedrange))

If oldusedrng.Rows.Count > 1 Then                           'If sheet is not empty
    Set newWs = Sheets.Add(after:=oldWs)                    'Add new sheet
        With oldusedrng
            .AutoFilter Field:=FieldNum, Criteria1:=Criteria
            .Copy                                               'Copy visible data
        End With
        With newWs.Cells
            .PasteSpecial xlPasteColumnWidths
            .PasteSpecial xlPasteAll                            'Paste data on new sheet
            .Cells(1, 1).Select                                 'Deselect paste area
            .Cells(1, 1).Copy                                   'Clear Clipboard
        End With
        oldWs.Delete                                            'Delete old sheet
        newWs.Name = wsName
End If

End Function

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

Sub ������_���()


FastWB True


Sheets(1).Activate

botrow1 = Cells(Rows.Count, 1).End(xlUp).Row

Range("A1").CurrentRegion.Copy


Workbooks.Open Filename:= _
        "C:\Users\MngLOg_01\Desktop\�������!\�������� �������\������.xlsx"
Windows("������.xlsx").Activate

Sleep (500)

Range("A1").PasteSpecial xlPasteValues

Dim x As Integer
' create an array of data
'������� ������ ������
'����������� ������� (������� ������) � ��������� ���� (�������) � 4�� �������

Set r = Range("A1:T10").Find("���� �������").Resize(botrow1, 1) '�������� ����� ��� ������ � ������ �������

Range("I1:I" & botrow1).Value = r.Value '� 7�� ������� ������ ��� ���� �������


Range("D1,G1,J1:U1").EntireColumn.Delete
Cells(1, 5).Value = "����"
Cells(1, 6).Value = "��������"
Cells(1, 7).Value = "���� �������"

Range("A2:A" & botrow1).NumberFormat = "0" '������ �������� �������� � int �������
Range("D2:D" & botrow1).NumberFormat = "0"
Range("G2:G" & botrow1).NumberFormat = "m/d/yyyy"

Dim myList As Object
Set myList = CreateObject("Scripting.Dictionary")



'��������� � ������ ������ �� �����2, �������1 (������ � ������� ����� �������)
For x = 2 To ActiveWorkbook.Sheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    myList.Add ActiveWorkbook.Sheets(2).Cells(x, 1).Text, 1
Next

'����������� � ������ ������� ��������, 1 ���� ������ ���� � �������, ����� ���� ���.

For x = 2 To Cells(Rows.Count, 1).End(xlUp).Row
If myList.exists(Cells(x, 3).Text) Then
Cells(x, 6).Value = 1
End If
Next

'������� ������ ������ (��� ������ � ������� � 6�� ������� �������� �� ����� 1)

Call DeleteRows(6, "=1")

x = Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(2, 5), Cells(x, 5)).Value = Date

Dim BotRowOld As Integer

y = Cells(2, Columns.Count).End(xlToLeft).Column

Range(Cells(2, 1), Cells(x, y)).Copy



Workbooks.Open Filename:= _
        "T:\06 ����� ���������\�������� �������.xlsx"
Windows("�������� �������.xlsx").Activate
Sleep (3000)
Sheets(1).Select
BotRowOld = Cells(Rows.Count, 1).End(xlUp).Row + 1
Cells(BotRowOld, 1).PasteSpecial xlPasteAll
Cells(1, 1).Select                                 'Deselect paste area
Cells(1, 1).Copy


BotRow = Cells(Rows.Count, 1).End(xlUp).Row
For x = BotRowOld To BotRow 'x �������� �� BOTROWOLD!!
Cells(x, 8).Value = Cells(x, 7).Value
Cells(x, 7).Value = Cells(x, 1).Text + Cells(x, 3).Text + Cells(x, 4).Text
Next

'������� ���� Dictionary
myList.RemoveAll

'��������� ��������

For x = BotRowOld To BotRow
    On Error Resume Next
    myList.Add Cells(x, 7).Text, 1
Next

'��������� �� ������� ���������, ���� ����, � ������� �������� ��������� 1, ������� � ������� 6 �������, ���� ��� - �
If BotRowOld > 2 Then
For x = 2 To BotRowOld - 1 '������������ botrowold
If myList.exists(Cells(x, 7).Text) Then
myList(Cells(x, 7).Text) = myList(Cells(x, 7).Text) + 1
Cells(x, 6).Value = myList(Cells(x, 7).Text)
Else
Cells(x, 6).Value = 0
End If
Next
End If


Call DeleteRows(6, ">0")

Dim oldusedrng As Range

Sheets(2).Usedrange.ClearContents
x = Cells(Rows.Count, 1).End(xlUp).Row
y = Cells(2, Columns.Count).End(xlToLeft).Column

Set oldusedrng = Sheets(1).Range(Cells(1, 1), Cells(x, y))

FieldNum = 6
Criteria = ">14" '����� 14 ���������� ����� �������� � ���� ���������!!!

With oldusedrng
            .AutoFilter Field:=FieldNum, Criteria1:=Criteria
            .Copy                                               'Copy visible data
'            .AutoFilter
End With
'
Sheets(2).Select
ActiveSheet.Cells(1, 1).PasteSpecial xlPasteAll

oldusedrng.AutoFilter

Cells(1, 1).Select


    FastWB False

Debug.Print Timer - T & " seconds"

Windows("�������� �������.xlsx").Activate
Cells(1, 1).Select

    
End Sub

Sub ����������2()

    FastWB True

Dim bk1, bk2 As Workbook

'������� ����� ������ � ������� � ��� ���������� ����� 2 (�������� �������), ����������� ����� ����� bk1

Range("A1").CurrentRegion.Copy
Workbooks.Add
Range("A1").PasteSpecial xlPasteValues
Set bk1 = ActiveWorkbook

' ������� ������ �������, ����� ���� ������� ���������
Columns(5).Delete
Columns(5).Delete




Cells(1, 5).Value = "������� ���"
Cells(1, 6).Value = "��� ��������"
Cells(1, 7).Value = "���-�� 1�"
Cells(1, 8).Value = "�����"
Cells(1, 9).Value = "�������"
Cells(1, 10).Value = "���� �������"

Range("$A$1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6), _
        Header:=xlYes

BotRow = Cells(Rows.Count, 1).End(xlUp).Row

Range("J2:J" & BotRow).Value = Range("F2:F" & BotRow).Value
Range("J2:J" & BotRow).NumberFormat = "m/d/yyyy"

Columns("A:K").EntireColumn.AutoFit

Sheets.Add after:=Worksheets(1)


Workbooks.Open Filename:= _
        "D:\Obmen\mnglog_01\�������.xlsx"
Windows("�������.xlsx").Activate
Set bk2 = Workbooks("�������.xlsx")

bk2.Activate
BotRow = Cells(Rows.Count, 1).End(xlUp).Row - 1

'������� ������� � ������� ��������� ������ (������� �������)
art = Mid(Range("A1", GetMaxCell()).Find("�������").Address, 2, 1)
potok = Mid(Range("A1", GetMaxCell()).Find("�����").Address, 2, 1)
wms = Mid(Range("A1", GetMaxCell()).Find("��� WMS").Address, 2, 1)
ost = Mid(Range("A1", GetMaxCell()).Find("�������� �������").Address, 2, 1)


bk1.Activate


For x = 1 To BotRow - 9
bk1.Sheets(2).Range("A" & x).Value = Right(bk2.Sheets(1).Range(wms & x + 9).Value, 6) '��� WMS

bk1.Sheets(2).Range("B" & x).Value = bk2.Sheets(1).Range(ost & x + 9) '������� � 1�
bk1.Sheets(2).Range("C" & x).Value = bk2.Sheets(1).Range(potok & x + 9) '�����
bk1.Sheets(2).Range("D" & x).Value = bk2.Sheets(1).Range(art & x + 9) '�������

Next

bk2.Close

Worksheets(1).Activate

botrow1 = Cells(Rows.Count, 1).End(xlUp).Row
botrow2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row


For x = 2 To botrow1 '��� ��������� ������ ���� �� ����� ����� �������� � �������� (� ����� ����� �����)
On Error Resume Next
Range(Cells(x, 7), Cells(x, 9)).Value = Worksheets(2).Columns(1).Find(Cells(x, 1).Value, lookat:=xlWhole).Offset(0, 1).Resize(1, 3).Value
Cells(x, 5).Value = Application.Max(Cells(x, 4).Value - Cells(x, 7).Value, 0)
If Cells(x, 8).Value > 8 Then
Cells(x, 6).Value = "���"
Else
Cells(x, 6).Value = "���"
End If
Next
'
'��������� ������� �� ������� � �� �������� � ��������, � ��������� �� ����������� ������� ������
Range("A1").CurrentRegion.Sort key1:=Range("E1:E" & botrow1), order1:=xlDescending, _
    Header:=xlYes


Range("A1:J1").Font.Bold = True '������ ����� ������ ������
Range("A1:J" & botrow1).Borders.LineStyle = True '������ ����� ��� �����
Range("I2:I" & botrow1).NumberFormat = "0"

Cells(1, 1).Select
Cells(1, 1).Copy

    FastWB False

ActiveWorkbook.SaveAs Filename:="D:\Obmen\mnglog_01\������ " & Date & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False


End Sub




