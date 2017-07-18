'-----------------------------------------------------------------------------------------------------------------
'Script Name  - Logout
'Description  - Script is used to logout from SFA Application
'Created By   -
'Created On   -
'Modified By  -
'Modified On  -
'Authour      - CGI
'-----------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------
'Environment setup
scriptpath = environment("ScriptPath1")
environment.value("varpath") = Mid(scriptpath,1,Instrrev(Mid(scriptpath,1,instrrev(scriptpath,"\")-1),"\"))
Repositoriescollection.Add environment.value("varpath")&"ObjectRepository\LogOut.tsr"
Browser("Home | Salesforce").Page("Home | Salesforce").Sync
wait 5
Browser("Home | Salesforce").Page("Home | Salesforce").WebButton("View profile").Click @@ hightlight id_;_Browser("Home | Salesforce").Page("Home | Salesforce").WebButton("View profile")_;_script infofile_;_ZIP::ssf1.xml_;_
wait 2
Browser("Home | Salesforce").Page("Home | Salesforce").Link("Log Out").Click
wait 4
Browser("Home | Salesforce").Page("Login | Salesforce").Sync
checkURL = Browser("Home | Salesforce").Page("Login | Salesforce").GetROProperty("URl")
	If checkURL = "https://empower--uftpoc.cs83.my.salesforce.com" Then
		AddNewCase strTCID,"","Logout verification","Logout is successfull","Pass"
	Else
		AddNewCase strTCID,"","Logout verification","Logout is not successfull","Fail"
End If
