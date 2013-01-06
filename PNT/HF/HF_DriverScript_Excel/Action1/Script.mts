'Driver Script for HF

ExecuteFile "C:\Users\Srikanth\Desktop\PNT\HF\Keyword_Functions.qfl"

'DataTable.AddSheet "dtTC"
'DataTable.AddSheet "dtTS"

'Import from External Excel Sheet

'DataTable.ImportSheet "C:\Users\Srikanth\Desktop\PNT\HF\HF_TestData.xls","TestCases","dtTC"
'DataTable.ImportSheet "C:\Users\Srikanth\Desktop\PNT\HF\HF_TestData.xls","TestScripts","dtTS"


Set objExcl=CreateObject("Excel.Application")
Set objExclBook=objExcl.Workbooks.Open("C:\Users\Srikanth\Desktop\PNT\HF\HF_TestData.xls")
Set objTCSheet=objExclBook.Worksheets("TestCases")
Set objTSSheet=objExclBook.Worksheets("TestScripts")

'Row Count for Test Cases and Test Scripts Sheet
rcnt_tc=objTCSheet.usedrange.rows.count
rcnt_ts=objTSSheet.usedrange.rows.count

For i=1 to rcnt_tc
'	DataTable.GetSheet("dtTC").SetCurrentRow(i)
	If objTCSheet.Cells(i,4)="Y" Then
		For j=1 to rcnt_ts
'			DataTable.GetSheet("dtTS").SetCurrentRow(j)
			If objTCSheet.Cells(i,1)=objTSSheet.Cells(j,1) Then
				'Execute Keyword Function
				vIP1=objTSSheet.Cells(j,6)
				vIP2=objTSSheet.Cells(j,7)
				vKeyword=objTSSheet.Cells(j,5)
				KeywordFunction_Executor vKeyword,vIP1,vIP2
				If Environment.Value("envResult")="Passed" Then
					objTCSheet.Cells(i,5)="Passed"
				Else
					objTCSheet.Cells(i,5)="Failed"	
				End If
			End If
		Next
	Else
		'
	End If
Next

objExclBook.Save
objExcl.Quit
Set objExcl=Nothing
Set objExclBook=Nothing
Set objTCSheet=Nothing
Set objTSSheet=Nothing


'DataTable.Export "C:\Users\Srikanth\Desktop\PNT\HF\HF_TestData.xls"


