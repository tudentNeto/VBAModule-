Attribute VB_Name = "Module1"
Option Compare Database

Public Sub dhSendEXOSTATOK() '(strSource As String)
Const constShift = 2
Const constShiftj = 0
Const iShiftMerge = 2
Const iShiftMergeAll = 5
Const iShiftGridStart = 5
Const iShiftGridEnd = 22
Const conpath = "---"
Dim arrHead(13) As myTypeArr
Dim objExcApp As Excel.Application
Dim objWorkBook As Excel.Workbook
Dim objSheet As Excel.Worksheet
Dim rst As DAO.Recordset
Dim iCols, jCols As Integer
Dim kLoc, kAll As Integer
Dim SumLoc, SumAll As Double
Dim npp, nppRest As Integer
Dim rstCount As Integer
Dim iFor, jFor As Integer
Dim fl As Boolean
Dim nLoop As Integer

arrHead(0).SizeCol = 3
arrHead(0).IsUp = False
arrHead(0).NameCol = "Hом. по порядк"
arrHead(0).NameColUp = ""
arrHead(0).shift = 0

arrHead(1).SizeCol = 8.43
arrHead(1).IsUp = False
arrHead(1).NameCol = "Счет. cубсчет"
arrHead(1).NameColUp = ""
arrHead(1).shift = 0

arrHead(2).SizeCol = 40
arrHead(2).IsUp = False
arrHead(2).NameCol = "Наименование, характеристика    (вид, сорт, группа)"
arrHead(2).NameColUp = ""
arrHead(2).shift = 2

arrHead(3).SizeCol = 15
arrHead(3).IsUp = True
arrHead(3).NameCol = "Kод (номенклатурный номeр)"
arrHead(3).NameColUp = "Товарно-материвальые ценности"
arrHead(3).shift = 2

arrHead(4).SizeCol = 8.43
arrHead(4).IsUp = False
arrHead(4).NameCol = "Код по ОКЕИ"
arrHead(4).NameColUp = ""
arrHead(4).shift = 2

arrHead(5).SizeCol = 8.43
arrHead(5).IsUp = True
arrHead(5).NameCol = "Наиме-нование"
arrHead(5).NameColUp = "Ед.изм."
arrHead(5).shift = 2

arrHead(6).SizeCol = 8.86
arrHead(6).IsUp = False
arrHead(6).NameCol = "Цена руб.коп."
arrHead(6).NameColUp = ""
arrHead(6).shift = 0

arrHead(7).SizeCol = 8.43
arrHead(7).IsUp = False
arrHead(7).NameCol = "Инвен-тарный"
arrHead(7).NameColUp = ""
arrHead(7).shift = 2

arrHead(8).SizeCol = 8.43
arrHead(8).IsUp = True
arrHead(8).NameCol = "Паспо-рта"
arrHead(8).NameColUp = "Номер"
arrHead(8).shift = 2

arrHead(9).SizeCol = 8.43
arrHead(9).IsUp = False
arrHead(9).NameCol = "Коли-чество"
arrHead(9).NameColUp = ""
arrHead(9).shift = 2

arrHead(10).SizeCol = 8.86
arrHead(10).IsUp = True
arrHead(10).NameCol = "Сумма руб.коп."
arrHead(10).NameColUp = "Фактическое наличие"
arrHead(10).shift = 2

arrHead(11).SizeCol = 8.43
arrHead(11).IsUp = False
arrHead(11).NameCol = "Коли-чество"
arrHead(11).NameColUp = ""
arrHead(11).shift = 2

arrHead(12).SizeCol = 8.86
arrHead(12).IsUp = True
arrHead(12).NameCol = "Сумма руб.коп."
arrHead(12).NameColUp = "По данным бухгалтерии"
arrHead(12).shift = 2

strPath = conpath
Set objExcApp = New Excel.Application
Set objWorkBook = objExcApp.Workbooks.Add
Set objSheet = objWorkBook.Sheets("Лист1")
'ОТРИСОВЫВАЕМ ШАПКУ
'**********************************************
iCols = constShift
npp = 1
nLoop = 0
Set rst = CurrentDb.OpenRecordset("---")
rstCount = rst.RecordCount
rst.MoveFirst
Do While Not rstCount - npp < 0

For jCols = 1 To 13
i = iCols + arrHead(jCols - 1).shift
    objSheet.Cells(i, constShiftj + jCols).Value = arrHead(jCols - 1).NameCol
    objSheet.Cells(iCols + 6, constShiftj + jCols).Value = jCols
    objSheet.Cells(iCols + 6, constShiftj + jCols).HorizontalAlignment = xlCenter
    objSheet.Range(objSheet.Cells(iCols, constShiftj + jCols), objSheet.Cells(iCols + 22, constShiftj + jCols)).Cells.ColumnWidth = arrHead(jCols - 1).SizeCol
    objSheet.Range(objSheet.Cells(i, constShiftj + jCols), objSheet.Cells(i + 5 - arrHead(jCols - 1).shift, constShiftj + jCols)).Cells.Merge
    objSheet.Range(objSheet.Cells(i, constShiftj + jCols), objSheet.Cells(i + 5 - arrHead(jCols - 1).shift, constShiftj + jCols)).Cells.WrapText = True
    objSheet.Range(objSheet.Cells(i, constShiftj + jCols), objSheet.Cells(i + 5 - arrHead(jCols - 1).shift, constShiftj + jCols)).Cells.HorizontalAlignment = xlCenter
    objSheet.Range(objSheet.Cells(i, constShiftj + jCols), objSheet.Cells(i + 5 - arrHead(jCols - 1).shift, constShiftj + jCols)).Cells.VerticalAlignment = xlTop
    objSheet.Range(objSheet.Cells(i, constShiftj + jCols), objSheet.Cells(i + 5 - arrHead(jCols - 1).shift, constShiftj + jCols)).Borders.LineStyle = xlContinuous
    If arrHead(jCols - 1).IsUp Then
        'строка заголовка двойная
        'вписываем заголовок
        objSheet.Cells(iCols, constShiftj + jCols - 1).Value = arrHead(jCols - 1).NameColUp
        'объединяем, форматируем, прорисовывает сетку соответствующим образом
        objSheet.Range(objSheet.Cells(iCols, constShiftj + jCols - 1), objSheet.Cells(iCols + 1, constShiftj + jCols)).Cells.Merge
        objSheet.Range(objSheet.Cells(iCols, constShiftj + jCols - 1), objSheet.Cells(iCols + 1, constShiftj + jCols)).Cells.WrapText = True
        objSheet.Range(objSheet.Cells(iCols, constShiftj + jCols - 1), objSheet.Cells(iCols + 1, constShiftj + jCols)).Cells.HorizontalAlignment = xlCenter
        objSheet.Range(objSheet.Cells(iCols, constShiftj + jCols - 1), objSheet.Cells(iCols + 1, constShiftj + jCols)).Cells.VerticalAlignment = xlTop
        objSheet.Range(objSheet.Cells(iCols, constShiftj + jCols - 1), objSheet.Cells(iCols + 1, constShiftj + jCols)).Borders.LineStyle = xlContinuous
    End If
Next jCols
objSheet.Range(objSheet.Cells(iCols + 6, constShiftj + 1), objSheet.Cells(iCols + 22, constShiftj + 13)).Borders.LineStyle = xlContinuous

kLoc = 0
SumLoc = 0
'вносим данные в таблицу, накапливаем кол-во и сумму, меняем npp
For iFor = 1 To 16
    objSheet.Cells(iCols + 6 + iFor, constShiftj + 1).Value = npp
    For jFor = 2 To 13
        objSheet.Cells(iCols + 6 + iFor, constShiftj + jFor).Value = rst(jFor - 2)
    Next jFor
    kLoc = kLoc + rst("KOL1")
    SumLoc = SumLoc + rst("SUM1")
    npp = npp + 1
rst.MoveNext
    If (rst.EOF) Then
        rst.MovePrevious
        fl = True
        nppRest = rstCount - nLoop * 16
        Exit For
    Else
        fl = False
    End If
Next iFor
'накапливаем сумму и количество за этот проход в общие
kAll = kAll + kLoc
SumAll = SumAll + SumLoc
objSheet.Range(objSheet.Cells(iCols + 7, 13), objSheet.Cells(iCols + 23, 13)).Cells.NumberFormat = "0.00"
objSheet.Range(objSheet.Cells(iCols + 7, 7), objSheet.Cells(iCols + 23, 7)).Cells.NumberFormat = "0.00"
objSheet.Range(objSheet.Cells(iCols + 7, 11), objSheet.Cells(iCols + 23, 11)).Cells.NumberFormat = "0.00"
'постраничные итоги
objSheet.Cells(iCols + 23, 1).Value = "Итого по странице:"
objSheet.Cells(iCols + 23, 9).Value = "Итого:"
objSheet.Cells(iCols + 23, 10).Value = kLoc
objSheet.Cells(iCols + 23, 11).Value = SumLoc
objSheet.Cells(iCols + 23, 12).Value = kLoc
objSheet.Cells(iCols + 23, 13).Value = SumLoc
objSheet.Cells(iCols + 24, 3).Value = "а)количество порядковых номеров"
objSheet.Cells(iCols + 24, 4).Value = Пропись_без_руб(IIf(fl, nppRest, 16), False)
objSheet.Range(objSheet.Cells(iCols + 24, 4), objSheet.Cells(iCols + 24, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 24, 4), objSheet.Cells(iCols + 24, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 24, 4), objSheet.Cells(iCols + 24, 12)).Cells.Borders(xlEdgeBottom).LineStyle = 1
objSheet.Cells(iCols + 25, 4).Value = "прописью"
objSheet.Range(objSheet.Cells(iCols + 25, 4), objSheet.Cells(iCols + 25, 12)).Cells.Font.Size = 6
objSheet.Range(objSheet.Cells(iCols + 25, 4), objSheet.Cells(iCols + 25, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 25, 4), objSheet.Cells(iCols + 25, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 25, 4), objSheet.Cells(iCols + 25, 12)).Cells.VerticalAlignment = xlTop
objSheet.Cells(iCols + 27, 3).Value = "б)общее количество единиц фактически"
objSheet.Cells(iCols + 27, 4).Value = Пропись_без_pуб(kLoc, False)
objSheet.Range(objSheet.Cells(iCols + 27, 4), objSheet.Cells(iCols + 27, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 27, 4), objSheet.Cells(iCols + 27, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 27, 4), objSheet.Cells(iCols + 27, 12)).Cells.Borders(xlEdgeBottom).LineStyle = 1
objSheet.Cells(iCols + 28, 4).Value = "прописью"
objSheet.Range(objSheet.Cells(iCols + 28, 4), objSheet.Cells(iCols + 28, 12)).Cells.Font.Size = 6
objSheet.Range(objSheet.Cells(iCols + 28, 4), objSheet.Cells(iCols + 28, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 28, 4), objSheet.Cells(iCols + 28, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 28, 4), objSheet.Cells(iCols + 28, 12)).Cells.VerticalAlignment = xlTop
objSheet.Cells(iCols + 30, 3).Value = "в)на сумму фактически"
objSheet.Cells(iCols + 30, 4).Value = Пропись(SumLoc, True)
objSheet.Range(objSheet.Cells(iCols + 30, 4), objSheet.Cells(iCols + 30, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 30, 4), objSheet.Cells(iCols + 30, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 30, 4), objSheet.Cells(iCols + 30, 12)).Cells.Borders(xlEdgeBottom).LineStyle = 1
objSheet.Cells(iCols + 31, 4).Value = "прописьо"
objSheet.Range(objSheet.Cells(iCols + 31, 4), objSheet.Cells(iCols + 31, 12)).Cells.Font.Size = 6
objSheet.Range(objSheet.Cells(iCols + 31, 4), objSheet.Cells(iCols + 31, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 31, 4), objSheet.Cells(iCols + 31, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 31, 4), objSheet.Cells(iCols + 31, 12)).Cells.VerticalAlignment = xlTop
iCols = iCols + 32
nLoop = nLoop + 1
Loop
'в конце, после последних постраничных итогов, выводщим общие итоги
iCols = iCols - 32
objSheet.Cells(iCols + 31, 1).Value = "Итого по описи:"
objSheet.Cells(iCols + 32, 3).Value = "а)количество порядковыгX номеров"
objSheet.Cells(iCols + 32, 4).Value = Пропись_без_руб(npp - 1, False)
objSheet.Range(objSheet.Cells(iCols + 32, 4), objSheet.Cells(iCols + 32, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 32, 4), objSheet.Cells(iCols + 32, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 32, 4), objSheet.Cells(iCols + 32, 12)).Cells.Borders(xlEdgeBottom).LineStyle = 1
objSheet.Cells(iCols + 33, 4).Value = "прописьо"
objSheet.Range(objSheet.Cells(iCols + 33, 4), objSheet.Cells(iCols + 33, 12)).Cells.Font.Size = 6
objSheet.Range(objSheet.Cells(iCols + 33, 4), objSheet.Cells(iCols + 33, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 33, 4), objSheet.Cells(iCols + 33, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 33, 4), objSheet.Cells(iCols + 33, 12)).Cells.VerticalAlignment = xlTop
objSheet.Cells(iCols + 35, 3).Value = "б)общее количество единиц фактически"
objSheet.Cells(iCols + 35, 4).Value = Пропись_без_руб(kAll, False)
objSheet.Range(objSheet.Cells(iCols + 35, 4), objSheet.Cells(iCols + 35, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 35, 4), objSheet.Cells(iCols + 35, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 35, 4), objSheet.Cells(iCols + 35, 12)).Cells.Borders(xlEdgeBottom).LineStyle = 1
objSheet.Cells(iCols + 36, 4).Value = "прописью"
objSheet.Range(objSheet.Cells(iCols + 36, 4), objSheet.Cells(iCols + 36, 12)).Cells.Font.Size = 6
objSheet.Range(objSheet.Cells(iCols + 36, 4), objSheet.Cells(iCols + 36, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 36, 4), objSheet.Cells(iCols + 36, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 36, 4), objSheet.Cells(iCols + 36, 12)).Cells.VerticalAlignment = xlTop
objSheet.Cells(iCols + 38, 3).Value = "в)на сумму фактически"
objSheet.Cells(iCols + 38, 4).Value = Пропись(SumAll, True)
objSheet.Range(objSheet.Cells(iCols + 38, 4), objSheet.Cells(iCols + 38, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 38, 4), objSheet.Cells(iCols + 38, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 38, 4), objSheet.Cells(iCols + 38, 12)).Cells.Borders(xlEdgeBottom).LineStyle = 1
objSheet.Cells(iCols + 39, 4).Value = "прописью"
objSheet.Range(objSheet.Cells(iCols + 39, 4), objSheet.Cells(iCols + 39, 12)).Cells.Font.Size = 6
objSheet.Range(objSheet.Cells(iCols + 39, 4), objSheet.Cells(iCols + 39, 12)).Cells.Merge
objSheet.Range(objSheet.Cells(iCols + 39, 4), objSheet.Cells(iCols + 39, 12)).Cells.HorizontalAlignment = xlCenter
objSheet.Range(objSheet.Cells(iCols + 39, 4), objSheet.Cells(iCols + 39, 12)).Cells.VerticalAlignment = xlTop
objWorkBook.SaveAs conpath & "---.xlsx"
Set objExcApp = Nothing
Set objWorkBook = Nothing
Set objSheet = Nothing

End Sub

