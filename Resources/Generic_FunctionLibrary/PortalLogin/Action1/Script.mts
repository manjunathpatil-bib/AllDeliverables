'-----------------------------------------------------------------------------------------------------------------
'Script Name  - Portal Login
'Description  - Login to the Portal for Inventory and Tire Search
'Created By   -
'Created On   -
'Modified By  -
'Modified On  -
'Authour      - CGI
'-----------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------
'Environment setup

' Variable Declartion
Dim RowCount
Dim RunTest
Dim Username
Dim Password
Dim URLExp
Dim BroserName
Dim BrowserInvoke
Dim URLApp
Dim Version
Dim URLAct



scriptpathLogin = environment("ScriptPath1")

environment.value("varpathLogin") = Mid(scriptpathLogin,1,Instrrev(Mid(scriptpathLogin,1,instrrev(scriptpathLogin,"\")-1),"\"))
'msgbox environment.value("varpath")

Repositoriescollection.Add environment.value("varpathLogin")&"ObjectRepository\PortalLogin_Tire.tsr"

varpath1 = environment.value("varpathLogin")
'msgbox varpath1

varpath1 = Mid(scriptpathLogin,1,Instrrev(Mid(varpath1,1,instrrev(varpath1,"\")-1),"\"))
'msgbox varpath1

environment.value("varpath1") = varpath1

Datatable.AddSheet "Sheet1"
Datatable.ImportSheet environment.value("varpath1")&"TestData\PortalLogin.xlsx","Sheet1","Sheet1"

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
		Browser("PortalBrowser").SetTOProperty "version",Version

		SystemUtil.Run BrowserInvoke,URLApp
		Browser("PortalBrowser").Page("PortalPage").Sync
		'Browser("PortalBrowser").FullScreen
		Wait(5)		
		If Browser("PortalBrowser").Page("PortalPage").WebEdit("UserName").Exist(20) Then
			Browser("PortalBrowser").Page("PortalPage").WebEdit("UserName").Set Username
		End If	
			Wait(2)
		If Browser("PortalBrowser").Page("PortalPage").WebEdit("Password").Exist(20) Then
			Browser("PortalBrowser").Page("PortalPage").WebEdit("Password").Set Password
		End If

		If Browser("PortalBrowser").Page("PortalPage").WebButton("Sign In").Exist(20) Then
		    Browser("PortalBrowser").Page("PortalPage").WebButton("Sign In").Click	
		End If	
		
		wait 5
		URLAct = Browser("PortalBrowser").Page("PortalPage").GetROProperty("URL")

		If URLExp = URLAct Then
				AddNewCase strTCID,"Login to Portal","User should be able to login to the Portal","User is able to login to the application","Pass"
				Else
				AddNewCase strTCID,"Login to Portal","User should be able to login to the Portal","User is not able to login to the application","Fail"
		End If
				
	End If

Next








