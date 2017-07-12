'Script Name     - MainScript
'Description     - Main Script
'Created By      -
'Created On      -
'Modified By     -
'Modified On     -
'Authour         - CGI
'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------

'Environment Setup

scriptpath = environment.value("TestDir")
'msgbox scriptpath
environment.value("varpath") =Mid(scriptpath,1,Instrrev(Mid(scriptpath,1,instrrev(scriptpath,"\")-1),"\"))
'Variable for the Login Recovery sequence
environment.value("intLoginAttempts")=0

Datatable.AddSheet "SheetMaster"
Datatable.ImportSheet environment.value("varpath")&"TestData\TestCaseSelection.xlsx","Sheet1","SheetMaster"
Rowcount = Datatable.Getsheet("SheetMaster").Getrowcount

'Initialize the report
'Check if the result file exists
If CheckIfFileExists(environment.value("varpath")&"Results\Test.html")=True Then
	OpenFile environment.value("varpath")&"Results\Test.html"
Else
	CreateFile(environment.value("varpath")&"Results\Test.html")
End If

'Close all the open browsers before execution
CloseAllOpenBrowsers

For i = 1 to Rowcount
	Datatable.SetCurrentRow(i)
	' The Test Script to be Run need to be set as 'Yes'
	RunScript = Datatable.GetSheet("SheetMaster").GetParameter("Run")
	'msgbox RunScript
	TestScript = Datatable.GetSheet("SheetMaster").GetParameter("TestCaseName")
	
	ScriptPath1 = Datatable.GetSheet("SheetMaster").GetParameter("ScriptPath1")
	If RunScript = "Yes" Then
		' The relevant Test Script based on 'Yes' will be executed.
		
		ScriptPath1 = Environment.Value("varpath")&ScriptPath1
		
		Environment.Value ("ScriptPath1")= ScriptPath1
	    'msgbox ScriptPath1
		RunAction TestScript
		
	End If

Next

'Closing the report file
CloseFile

'Close all the open browsers after execution
CloseAllOpenBrowsers
