Option Explicit
Const xlUp = -4162, xlPasteValues = -4163, xlNone = -4142
Public Const firstCol = 39, lastCol = 45
'''Public Const firstCol = 40, lastCol = 40

Public Const resNoTemplate = " template not found. Check the template.  "
Public Const resNoBOM = "Nothing is inside this BOM. First make the BOM."

Dim tblArea
Dim qtn, plant, sorg, template, serno
Dim qtyRows, rowCount, visibleRows, sapRow, goto_pos, grid, cell 
Dim bExit, bAbort, txtStatus
Dim intRow : intRow = 4
Dim iCol

'1. Запрашиваем файл QTN и получаем массив значений для последующего заполнения SAP Quotation
Dim excelFile
excelFile = selectExcel()

' Объявляем объект FileSystemObject
Dim fso, filePath, fileName, arrWords
' Устанавливаем путь к файлу
filePath = excelFile
' Создаем объект FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")
' Получаем имя файла из пути
fileName = fso.GetFileName(filePath)
' Выводим имя файла
WScript.Echo "Имя файла: " & fileName
' Освобождаем объект
Set fso = Nothing
arrWords = Split(fileName, " ")
qtn = arrWords(0)
If Not IsNumeric(qtn) Then
	MsgBox "Первое слово имени файла должно являтся номером QTN. " & qtn & " - не числовое значение!", vbSystemModal Or vbExclamation
	WScript.Quit
End If

WScript.Quit


'2.0 - открываем транзакцию
 qtn = "50002648"
 session.findById("wnd[0]").maximize
 session.findById("wnd[0]/tbar[0]/okcd").text = "VA22"
 session.findById("wnd[0]").sendVKey 0
 session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = qtn
 session.findById("wnd[0]").sendVKey 0

'2. Заполняем открытый SAP Quotation
Dim ArticlesExcel, objWorkbook, pmu, TextSheet
Set ArticlesExcel = CreateObject("Excel.Application")
Set objWorkbook = ArticlesExcel.Workbooks.Open(excelFile)
objWorkbook.Sheets("PMU").Activate
Set pmu = objWorkbook.Worksheets("PMU")
Dim iLastRow: iLastRow = CInt(0)
iLastRow =pmu.Range("A" & pmu.Rows.Count).End(xlUp).Row  
'WScript.Echo iLastRow


'On Error Resume Next
sapRow = 0
Do Until ArticlesExcel.Cells(intRow, firstCol).Value = ""
	'ReDim Preserve arrExcel(intRow - 4, 6)
	'WScript.Echo ArticlesExcel.Cells(intRow, firstCol).Value
'    Err.Clear
    tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
'''    tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_ANGEBOT", "GuiTableControl").Id
     Set grid = session.findById(tblArea)
'    sapRow = grid.currentRow                'Here is the current visible row of the QTN
'MsgBox "sap Row: " & sapRow, vbSystemModal Or vbInformation

	If sapRow > 7 Then
		rowCount = grid.RowCount
		goto_pos = session.findById(tblArea & "/txtVBAP-POSNR[0," & sapRow - 5 & "]").text
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
'''		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
		session.findById("wnd[1]/usr/txtRV45A-POSNR").text = goto_pos
		session.findById("wnd[1]/usr/txtRV45A-POSNR").caretPosition = 3
		session.findById("wnd[1]").sendVKey 0
		WScript.Sleep 300

		tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
'''		tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_ANGEBOT", "GuiTableControl").Id		
		Set grid = session.findById(tblArea)
		sapRow = grid.currentRow                'Here is the current visible row of the QTN
		Set cell = grid.GetCell(sapRow + 5, 1)
		cell.setFocus()
		sapRow = grid.currentRow                'Here is the current visible row of the QTN

	End If    

	For iCol = firstCol to lastCol
		grid.GetCell(sapRow, iCol - firstCol + 1).Text = Left(ArticlesExcel.Cells(intRow, iCol).Value, 40)
	Next 
	'WScript.Echo arrExcel(intRow - 4, 0)
	intRow = intRow + 1
	sapRow = sapRow + 1

Loop


'3. Транспонируем
transpose
Sub transpose()
  Dim sourceRange
  Dim targetRange

  Set sourceRange = pmu.Range(pmu.Cells(3, 3), pmu.Cells(iLastRow, 8))
  objWorkbook.Sheets("Text").Activate
  Set TextSheet = objWorkbook.Worksheets("Text")
  Set targetRange = TextSheet.Cells(1, 1)

  sourceRange.Copy
  targetRange.PasteSpecial xlPasteValues, xlNone, False, True
  pmu.Cells(3,3).Copy
End Sub



'4.Вставляем текстовые значения из транспонированной таблицы
tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
goto_pos = session.findById(tblArea & "/txtVBAP-POSNR[0,0]").text
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
'''		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
session.findById("wnd[1]/usr/txtRV45A-POSNR").text = 10
session.findById("wnd[1]/usr/txtRV45A-POSNR").caretPosition = 3
session.findById("wnd[1]").sendVKey 0
WScript.Sleep 300

Set grid = session.findById(tblArea)
sapRow = grid.currentRow  
Set cell = grid.GetCell(sapRow, 1)
cell.setFocus()	'перешли на нужную строку


'Подготовка заголовка
objWorkbook.Sheets("Text").Activate
Set TextSheet = objWorkbook.Worksheets("Text")
Dim spaces
Dim arrTexts(5, 1)
For intRow = 1 To 6		'делаем заголовки с пробелами до 18 символов
	spaces = ""
	if Len(TextSheet.Cells(intRow, 1).Value) < 18 Then
		spaces = Space(18 - Len(TextSheet.Cells(intRow, 1).Value))
	End if	
	arrTexts(intRow - 1, 0) = TextSheet.Cells(intRow, 1).Value & spaces
    'MsgBox arrTexts(intRow - 1, 0)
Next

Dim strText
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\08").select	'зашли в тексты позиции
intRow = 1
Do Until TextSheet.Cells(1, intRow + 1).Value = ""
	strText = ""
	For iCol = 1 To 6		'склеиваем заголовки со значениями
		strText = strText & arrTexts(iCol - 1, 0) & TextSheet.Cells(iCol, intRow + 1).Value & vbCrLf 
	Next	
	'MsgBox strText
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = strText
	if intRow < iLastRow then
		session.findById("wnd[0]/tbar[1]/btn[19]").press 'кнопка перехода по позициям
	end if	
    intRow = intRow + 1
Loop


'5. Вставка цен
session.findById("wnd[0]/tbar[0]/btn[3]").press()
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/tblSAPMV45ATCTRL_U_ERF_KONTRAKT/txtVBAP-ARKTX[4,0]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/tblSAPMV45ATCTRL_U_ERF_KONTRAKT/txtVBAP-ARKTX[4,0]").caretPosition = 2
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").select()
WScript.Sleep 300
tblArea = UserArea.findByName("SAPLV69ATCTRL_KONDITIONEN", "GuiTableControl").Id
Set grid = session.findById(tblArea)
objWorkbook.Sheets("PMU").Activate
intRow = 4
Dim iRow
Do Until ArticlesExcel.Cells(intRow, 12).Value = ""
	' Цикл для каждой строки в ценовых условиях
	tblArea = UserArea.findByName("SAPLV69ATCTRL_KONDITIONEN", "GuiTableControl").Id
	Set grid = session.findById(tblArea)
	qtyRows = grid.rowCount - 1
	MsgBox qtyRows
	iRow = 0
	MsgBox grid.GetCell(7, 3).Text 
	WScript.Sleep 300
	Do Until iRow > qtyRows
		'MsgBox "Row: " & intRow
		if grid.GetCell(iRow, 1).Text = "ZLS3" Then
			WScript.Sleep 100
			' ** didn't work
			' ** grid.GetCell(iRow, 3).Text = ArticlesExcel.Cells(intRow, 12).Value
			' ** grid.GetCell(iRow, 3).setFocus()
			' ** grid.GetCell(iRow, 3)..caretPosition = 16
			grid.GetCell(iRow, 1).setFocus()
			grid.GetCell(iRow, 1).caretPosition = 2	
			' get in the condition ZLS3
			session.findById("wnd[0]").sendVKey 2
			session.findById("wnd[0]/usr/txtKOMV-KBETR").text =  ArticlesExcel.Cells(intRow, 12).Value
			session.findById("wnd[0]/usr/txtKOMV-KBETR").caretPosition = 15
			session.findById("wnd[0]").sendVKey(0)
			session.findById("wnd[0]/tbar[0]/btn[3]").press()		
		
			session.findById("wnd[0]").sendVKey(0)
			session.findById("wnd[0]/tbar[1]/btn[19]").press()
			iRow = qtyRows + 100
		End If	
		iRow = iRow + 1
	Loop
	intRow = intRow + 1

	' If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[1,7]").text = "ZLS3" Then
	' 	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,7]").text = ArticlesExcel.Cells(intRow, 12).Value
	' 	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,7]").setFocus
	' 	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,7]").caretPosition = 9
	' 	session.findById("wnd[0]").sendVKey 0
	' 	session.findById("wnd[0]/tbar[1]/btn[19]").press
	' Else 
	' 	MsgBox "Условие ZLS3 не найдено! Переход к следующей строке.", vbSystemModal Or vbInformation
	' Endif
Loop

session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press

objWorkbook.Close False
ArticlesExcel.Quit
MsgBox "Script finished! ", vbSystemModal Or vbInformation

'====== Functions ans Subs ========

'returns an unique array from an Excel file chosen by a user
Function GetExcelArray()
	'excelFile = "C:\VBScript\articles.xlsx" ' Полный путь к выбранному файлу
	Dim ArticlesExcel, objWorkbook, ws
	Set ArticlesExcel = CreateObject("Excel.Application")
	Set objWorkbook = ArticlesExcel.Workbooks.Open(excelFile)
	objWorkbook.Sheets("PMU").Activate
	Set ws = objWorkbook.Worksheets("PMU")
	Dim collTemp : Set collTemp = CreateObject("Scripting.Dictionary")

	Dim iLastRow: iLastRow = CInt(0)
   	iLastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row  
	'WScript.Echo iLastRow

	' Считаем, что в 4 строке - начало таблицы для обработки
	Dim intRow : intRow = 4
	Dim iCol
	' Цикл для каждой строки
	On Error Resume Next
	Do Until ArticlesExcel.Cells(intRow, firstCol).Value = ""
		'ReDim Preserve arrExcel(intRow - 4, 6)
		'WScript.Echo ArticlesExcel.Cells(intRow, firstCol).Value
		For iCol = firstCol to lastCol
			arrExcel(intRow - 4, iCol - firstCol) = ArticlesExcel.Cells(intRow, iCol).Value
		Next 
		WScript.Echo arrExcel(intRow - 4, 0)
		intRow = intRow + 1
	Loop
	objWorkbook.Close False
	ArticlesExcel.Quit
	WScript.Echo Join(arrExcel)
	GetExcelArray = arrExcel
End Function





Sub OutputToExcel
	Dim ReportExcel, objWorkbook
	Set ReportExcel = CreateObject("Excel.Application")
	Set objWorkbook = ReportExcel.Workbooks.Add()
	ReportExcel.Visible = True
	
	arrReport = dicReport.Items
	intRow = 0
	For Each serno In arrSerno
		strReport = strReport & serno & " : " & arrReport(intRow) & VbCrLf
		ReportExcel.cells(intRow + 1, 1).value = serno
		ReportExcel.cells(intRow + 1, 2).value = arrReport(intRow)
		intRow = intRow + 1
	Next
End Sub



Sub notused


'StartTransaction("ZIB07")
session.findById("wnd[0]/tbar[0]/okcd").text = "ZIB07"
session.findById("wnd[0]").sendVKey 0

bAbort = vbFalse
For Each serno In arrSerno
	bExit = vbFalse
	session.findById("wnd[0]/usr/ctxtP_EQUNR").text = serno
	session.findById("wnd[0]/usr/ctxtP_WERKS2").text = plant
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	WScript.Sleep 500 'Delay for SAP processing
	If session.findById("wnd[0]/usr/ctxtP_EQUNR", False) Is Nothing Then
		Do While session.findById("wnd[0]/usr/chkJOB", False) Is Nothing
			If session.findById("wnd[1]/usr/txtLV_MATNR1", False) Is Nothing Then
				dicReport.Add serno, resNoBOM
				bExit = vbTrue
				session.findById("wnd[1]").sendVKey 0
				Exit Do
			Else
				session.findById("wnd[1]/tbar[0]/btn[8]").press 'V
				'session.findById("wnd[1]/tbar[0]/btn[2]").press       'X        
			End If
		Loop
		
		If Not bExit Then
			session.findById("wnd[0]/usr/chkJOB").selected = False
			session.findById("wnd[0]/usr/chkJOB").setFocus
			
			Set grid = session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell")
			
			qtyRows = grid.rowCount - 1
			'MsgBox "Rows amount: " & qtyRows
			visibleRows = grid.VisibleRowCount

			' Цикл для каждой строки
			'On Error Resume Next
			intRow = 0
			Do Until intRow > qtyRows
				'Err.Clear
				'MsgBox "Row: " & intRow
				grid.modifyCell intRow, "TEMPLATE", template
				grid.currentCellRow = intRow
				intRow = intRow + 1
			Loop
			grid.triggerModified
			session.findById("wnd[0]/tbar[1]/btn[8]").press
			'    MsgBox "Next Control - btn[3]", vbSystemModal Or vbInformation

			' It can be error that mat number not found - If for that
			If session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
			Else
				bAbort = vbTrue
				dicReport.Add serno, template & resNoTemplate
				session.findById("wnd[1]/tbar[0]/btn[0]").press
			End If
			
			If Not bAbort Then
				session.findById("wnd[0]/tbar[0]/btn[3]").press
				'    MsgBox "Next Control - wnd[1]/tbar[0]/btn[0]", vbSystemModal Or vbInformation
				session.findById("wnd[1]/tbar[0]/btn[0]").press
				dicReport.Add serno, resOK
			End If
		End If
	Else
		' Same selection window - check for status bar
		If session.ActiveWindow.findById("sbar", False) Is Nothing Then
			dicReport.Add serno, resExists
		Else
			txtStatus = session.ActiveWindow.findById("sbar").Text
			dicReport.Add serno, txtStatus
		End If
		
	End If
Next

OutputToExcel

End Sub