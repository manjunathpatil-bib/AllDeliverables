﻿'Script Name     - PortalSearch
'Description     - Portal Search Main script
'Created By      -
'Created On      -
'Modified By     -
'Modified On     -
'Authour         - CGI
'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------

'Environment Setup

' Variable Declaration
Dim Rowcount
Dim RunScript
Dim TestScript

scriptpath = environment.value("TestDir")

environment.value("varpath") =Mid(scriptpath,1,Instrrev(Mid(scriptpath,1,instrrev(scriptpath,"\")-1),"\"))

Datatable.AddSheet "SheetMaster"
Datatable.ImportSheet environment.value("varpath")&"TestData\PortalTestCase.xlsx","Sheet1","SheetMaster"
Rowcount = Datatable.Getsheet("SheetMaster").Getrowcount

'Initialize the report
OpenFile environment.value("varpath")&"Results\Test.html"

For i = 1 to Rowcount
	Datatable.SetCurrentRow(i)
	' The Test Script to be Run need to be set as 'Yes'
	RunScript = Datatable.GetSheet("SheetMaster").GetParameter("Run")
	TestScript = Datatable.GetSheet("SheetMaster").GetParameter("TestCaseName")
	
	ScriptPath1 = Datatable.GetSheet("SheetMaster").GetParameter("ScriptPath1")
	If RunScript = "Yes" Then
	' The relevant Test Script based on 'Yes' will be executed.
	
	ScriptPath1 = Environment.Value("varpath")&ScriptPath1
	
	Environment.Value ("ScriptPath1")= ScriptPath1
    RunAction TestScript
	End If

Next

'Closing the report file
CloseFile
