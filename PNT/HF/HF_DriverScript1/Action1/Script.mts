'Driver Script for HF

ExecuteFile "C:\Users\Srikanth\Desktop\PNT\HF\Keyword_Functions.qfl"

DataTable.AddSheet "dtTC"
DataTable.AddSheet "dtTS"

'Import from External Excel Sheet

DataTable.ImportSheet "C:\Users\Srikanth\Desktop\PNT\HF\HF_TestData.xls","TestCases","dtTC"
DataTable.ImportSheet "C:\Users\Srikanth\Desktop\PNT\HF\HF_TestData.xls","TestScripts","dtTS"

'Row Count for Test Cases and Test Scripts Sheet
rcnt_tc=DataTable.GetSheet("dtTC").GetRowCount
rcnt_ts=DataTable.GetSheet("dtTS").GetRowCount

For i=1 to rcnt_tc
	DataTable.GetSheet("dtTC").SetCurrentRow(i)
	If DataTable.Value("Execute","dtTC")="Y" Then
		For j=1 to rcnt_ts
			DataTable.GetSheet("dtTS").SetCurrentRow(j)
			If DataTable.Value("TestCase_ID","dtTC")=DataTable.Value("TestCase_ID","dtTS") Then
				Environment.Value("envResult")="Passed"
				'Execute Keyword Function
				vIP1=DataTable.Value("IP1","dtTS")
				vIP2=DataTable.Value("IP2","dtTS")
				vKeyword=DataTable.Value("Keyword_Function","dtTS")
				KeywordFunction_Executor vKeyword,vIP1,vIP2
				If Environment.Value("envResult")="Failed" Then
					DataTable.Value("Result","dtTS")="Failed"	
				Else
					DataTable.Value("Result","dtTS")="Passed"									
				End If
			End If
		Next
	Else
		'
	End If
Next



'DataTable.Export "C:\Users\Srikanth\Desktop\PNT\HF\HF_TestData.xls"



