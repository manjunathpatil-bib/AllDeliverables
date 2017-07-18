'-----------------------------------------------------------------------------------------------------------------
'Script Name  - Login
'Description  - Login to the SFA Application
'Created By   -
'Created On   -
'Modified By  -
'Modified On  -
'Authour      - CGI
'-----------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------
'Environment setup

scriptpathLogin = environment("ScriptPath1")

environment.value("varpathLogin") = Mid(scriptpathLogin,1,Instrrev(Mid(scriptpathLogin,1,instrrev(scriptpathLogin,"\")-1),"\"))
'msgbox environment.value("varpath")

'Associate the repository only for the first login attempt
If environment.value("intLoginAttempts")=0 Then
	Repositoriescollection.Add environment.value("varpathLogin")&"ObjectRepository\Login.tsr"
End If


varpath1 = environment.value("varpathLogin")
'msgbox varpath1

varpath1 = Mid(scriptpathLogin,1,Instrrev(Mid(varpath1,1,instrrev(varpath1,"\")-1),"\"))
'msgbox varpath1

environment.value("varpath1") = varpath1

Datatable.AddSheet "Sheet1"
Datatable.ImportSheet environment.value("varpath1")&"TestData\Login.xlsx","Sheet1","Sheet1"

'Datatable.getsheet("Sheet1").SetCurrentRow 1
RowCount = Datatable.GetSheet("Sheet1").GetRowCount


For i = 1 To RowCount
	Datatable.SetCurrentRow(i)
	RunTest = Datatable.GetSheet("Sheet1").GetParameter("Run")
	
	If RunTest = "Yes" Then
		'Datatable.SetCurrentRow(i)
		Username  = datatable.Value("UserName","Sheet1")
		'msgbox Username
		Password = datatable.Value("Password","Sheet1")
		URLExp = datatable.Value("LoginURL","Sheet1")
		
		BroserName = datatable.Value("Browser","Sheet1")
		BrowserInvoke = datatable.Value("BrowserPath","Sheet1")
		URLApp = datatable.Value("AppURL","Sheet1")
		Version = datatable.Value("BrowserVersion","Sheet1")
		Browser("Accounts | Salesforce").SetTOProperty "version",Version

		SystemUtil.Run BrowserInvoke,URLApp
		wait(5)
		
		Browser("Accounts | Salesforce").Page("Login | Salesforce").Sync
		wait(5)
			If Browser("Accounts | Salesforce").Page("Login | Salesforce").WebEdit("username").Exist(20) Then
				Browser("Accounts | Salesforce").Page("Login | Salesforce").WebEdit("username").Set Username
				Browser("Accounts | Salesforce").Page("Login | Salesforce").WebEdit("pw").Set Password
				wait 2
				Browser("Accounts | Salesforce").Page("Login | Salesforce").WebButton("Log In to Sandbox").Click
				'wait 15
			End If
			'Some buffer time for network latency
			Wait 10
			'URLAct = Browser("Accounts | Salesforce").Page("Login | Salesforce").GetROProperty("URL")
			If Browser("Dashboards | Salesforce").Page("Home | Salesforce").WebButton("App Launcher").Exist(conExistTimeout) Then
				AddNewCase strTCID,"Login to Salesforce","User should be able to login to the application","User is able to login to the application","Pass"
				Else
				'Reporter.ReportEvent micFail, "Login page", "Home page is not shown"
				'Initiate Login Recovery on error
				If environment.value("intLoginAttempts")=0 Then
					environment.value("intLoginAttempts")=1
					LoginRecoverySequence
				End If
				AddNewCase strTCID,"Login to Salesforce","User should be able to login to the application","User is not able to login to the application","Fail"
			End If
			'If URLExp = URLAct Then
				'Reporter.ReportEvent micPass, "Login page", "Home page is shown successfully"
			'End If
	End If

Next








