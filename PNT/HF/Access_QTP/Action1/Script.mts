'Access

'XML with QTP

'Class Name- ADODB.Connection

'Record Set Class Name- ADODB.RecordSet

Set objAccess=CreateObject("ADODB.Connection")

objAccess.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Srikanth\Desktop\PNT\HF\HF_TestData.xls")


strQuery="Select TestCase_ID,TestCase_Name FROM Test_Scripts"
Set objRecSet=CreateObject("ADODB.Recordset")
objRecSet.Open strQuery,objAccess

'Print the Queried Records


'EOF


While not objRecSet.EOF
	Print objRecSet.Fields("TestCase_ID")
	Print objRecSet.Fields("TestCase_Name")
	objRecSet.MoveNext
Wend





