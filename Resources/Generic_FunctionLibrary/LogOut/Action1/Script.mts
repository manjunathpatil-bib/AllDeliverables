'-----------------------------------------------------------------------------------------------------------------
'Script Name  - Logout
'Description  - Logout from SFA Application
'Created By   -
'Created On   -
'Modified By  -
'Modified On  -
'Authour      - CGI
'-----------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------
'Environment setup


scriptpath = environment("ScriptPath1")

'msgbox scriptpath
'If scriptpath = Empty Then
'	scriptpath = environment.value("TestDir")
'	'msgbox scriptpath
'End If

'scriptpath = environment.value("TestDir")

environment.value("varpath") = Mid(scriptpath,1,Instrrev(Mid(scriptpath,1,instrrev(scriptpath,"\")-1),"\"))

Repositoriescollection.Add environment.value("varpath")&"ObjectRepository\LogOut.tsr"

wait 5
Browser("Home | Salesforce").Page("Home | Salesforce").WebButton("View profile").Click @@ hightlight id_;_Browser("Home | Salesforce").Page("Home | Salesforce").WebButton("View profile")_;_script infofile_;_ZIP::ssf1.xml_;_
wait 2
Browser("Home | Salesforce_2").Page("Home | Salesforce").Link("Log Out").Click @@ hightlight id_;_Browser("Home | Salesforce 2").Page("Home | Salesforce").Link("Log Out")_;_script infofile_;_ZIP::ssf2.xml_;_
wait 4
'Browser("Home | Salesforce_2").Close

	wait 5
	Browser("Home | Salesforce").Page("Login | Salesforce").Sync
	checkURL = Browser("Home | Salesforce").Page("Login | Salesforce").GetROProperty("URl")
	'msgbox checkURL
		If checkURL = "https://empower--uftpoc.cs83.my.salesforce.com" Then
		Reporter.ReportEvent micPass, "Logout page", "Logout is successfull"
		Else
		Reporter.ReportEvent micFail, "Logout page", "Logout is failed "
	End If
