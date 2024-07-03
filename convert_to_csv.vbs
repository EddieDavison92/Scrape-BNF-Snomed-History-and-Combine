Dim objExcel, objWorkbook
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False

Dim args
Set args = WScript.Arguments
inputFile = args(0)
outputFile = args(1)

Set objWorkbook = objExcel.Workbooks.Open(inputFile)
objWorkbook.SaveAs outputFile, 6 ' 6 specifies CSV format
objWorkbook.Close False

objExcel.Quit
