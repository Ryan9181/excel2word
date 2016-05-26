'''
''' @File 指标Excel转Word.vbs
''' @Auth Ryan
''' @Date 2016/05/12  ***
'''

Dim xlsfile, docfile
Dim xlsObj, docObj

Const sheetIndex  = 2

CALL Main


'
' 主入口
'
Function Main()
	if ShowUsage() Then
		Exit Function
	End IF
	
	xlsfile  = "C:\Downloads\1.xlsx"
	docfile = "C:\Downloads\1.docx"
	
	Wscript.echo "CreateObject..." & vbNewLine
	Set xlsObj 	= CreateObject("Excel.Application")
	Set docObj = CreateObject("Word.Application")
	xlsObj.Visible = True
	docObj.Visible = True
	
	Wscript.echo "DoTransform..." & vbNewLine
	
	
	Call DoTransform
	
	
	Wscript.echo "Clan...." & vbNewLine
	xlsObj.Quit
	'docObj.Quit
	
	Set xlsObj 	= Nothing
	Set docObj = Nothing
	
	Wscript.echo "Quit" & vbNewLine
End Function


Function ShowUsage()
	If Wscript.Arguments.count < 2 Then
		Wscript.echo "Usage: This.VBS <Excel file> <Word file>"
		ShowUsage = True
	Else 
		xlsfile = Wscript.Arguments(0)
		docfile = Wscript.Arguments(1)
		ShowUsage = False
	End If
End Function



Function DoTransform()
	Dim workbook, sheet
	Dim document, range, table
	Dim i, j, endpos
	Dim rowtext
	
	
	Set workbook = xlsObj.Workbooks.Open(xlsfile,,True)
	Set document = docObj.Documents.open(docfile)
	
	Set sheet = workbook.Sheets(sheetIndex)
	
	' 保存样式表格
	document.Tables(1).Range.Copy
	
	For i = 2 To sheet.UsedRange.Rows.Count
		endpos = document.Tables(i-1).Range.End
		
		Set range = document.Range(endpos, endpos)
		range.InsertAfter(sheet.Cells(i, 7))
		range.Style = document.Styles("H5")
		
		document.Range(range.End, range.End).Paste
		
		set table = document.Tables(i)
		With table
			.Cell(4,2).Range = sheet.Cells(i,7)   	'名称
			.Cell(5,2).Range = sheet.Cells(i,8)		'定义
			.Cell(7,2).Range = sheet.Cells(i,7)		'指标用途
			.Cell(8,2).Range = sheet.Cells(i,12)	'指标类别
			.Cell(11,2).Range = sheet.Cells(i,15)	'数据流来源
			.Cell(12,2).Range = sheet.Cells(i,17)	'适用范围
			.Cell(14,2).Range = sheet.Cells(i,18)	'计算频率
			.Cell(15,2).Range = sheet.Cells(i,9)	'计量单位
			.Cell(16,2).Range = sheet.Cells(i,18)	'考察时间
			.Cell(17,2).Range = sheet.Cells(i,10)	'维度
			.Cell(18,2).Range = sheet.Cells(i,5)	'数据流程对应	
			
			rowtext = (i-1) & "  " & sheet.Cells(i,5) & "  " & sheet.Cells(i,7) & "  " & sheet.Cells(i,8) & "  " & sheet.Cells(i,9) & "  " & sheet.Cells(i,10) & "  " & sheet.Cells(i,12) & "  " & sheet.Cells(i,15) & "  " & sheet.Cells(i,17) &  "  " & sheet.Cells(i,18)
			Wscript.echo rowtext & vbNewLine
		End With
	Next
	
	workbook.Saved = True
	xlsObj.Workbooks.Close
	
	document.Saved = True
	'docObj.Documents.Close
	
	set sheet = Nothing
	set workbook = Nothing	
End Function
