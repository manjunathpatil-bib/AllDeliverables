'Script Name     - TireSearch
'Description     - Script performs basic search on the tire search portal
'Created By      -
'Created On      -
'Modified By     -
'Modified On     -
'Authour         - CGI
'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------

'Environment Setup

Dim RowCount
Dim RunTest
Dim TirePartNo
Dim TireProductNo
Dim TireQty
Dim CheckLenPartNo
Dim CheckLenProductNo
Dim ActErrMessage
Dim ValidMessage


scriptpathLogin = environment("ScriptPath1")
environment.value("varpathLogin") = Mid(scriptpathLogin,1,Instrrev(Mid(scriptpathLogin,1,instrrev(scriptpathLogin,"\")-1),"\"))
Repositoriescollection.Add environment.value("varpathLogin")&"ObjectRepository\TireSearch.tsr"
varpath1 = environment.value("varpathLogin")
varpath1 = Mid(scriptpathLogin,1,Instrrev(Mid(varpath1,1,instrrev(varpath1,"\")-1),"\"))
environment.value("varpath1") = varpath1
Datatable.AddSheet "Sheet1"
Datatable.ImportSheet environment.value("varpath1")&"TestData\TireSearch.xlsx","Sheet1","Sheet1"
RowCount = Datatable.GetSheet("Sheet1").GetRowCount

For i = 1 To RowCount
	Datatable.SetCurrentRow(i)
	RunTest = Datatable.GetSheet("Sheet1").GetParameter("Run")
	
	If RunTest = "Yes" Then
		TirePartNo = datatable.Value("SearchPartNo","Sheet1")
		TireProductNo = datatable.Value("SearchProductNo","Sheet1")
		TireQty = datatable.Value("SearchQty","Sheet1")
		' Validation message for negative search
		ValidMessage = datatable.Value("ExpMessage","Sheet1")
		Browser("PortalSearch").Page("TireSearch").Sync
		Browser("PortalSearch").Page("TireSearch").WebElement("WebElement").Click
		CheckLenPartNo = Len(TirePartNo)
		CheckLenProductNo = Len(TireProductNo)		
		'Enter the Part Number to search
		If CheckLenPartNo > 0 and CheckLenProductNo = 0 Then
			Browser("PortalSearch").Page("Products").Sync
			Browser("PortalSearch").Page("Products").WebEdit("Search Products").Set TirePartNo 
		End If	
		'Enter the Product Number to search		
		If CheckLenProductNo > 0 and CheckLenPartNo = 0 Then
			Browser("PortalSearch").Page("Products").Sync
			Browser("PortalSearch").Page("Products").WebEdit("Search Products").Set TireProductNo 
		End If
		'Enter the quantity to search
		Browser("PortalSearch").Page("Products").WebNumber("Quantity").Set TireQty ' 1
		If Browser("PortalSearch").Page("Products").Link("SEARCH").Exist Then
			wait 2
			Browser("PortalSearch").Page("Products").Link("SEARCH").Click
		End If
		wait 5
		Browser("Products").Sync
		'Validation for Postive search
		If Browser("Products").Page("Products").WebElement("PartNo").Exist Then
			ActPartNo = Browser("Products").Page("Products").WebElement("PartNo").GetROProperty("innertext")
			'msgbox ActPartNo
		End If
		If Browser("Products").Page("Products").WebElement("ProductNo").Exist Then
			ActProductNo = Browser("Products").Page("Products").WebElement("ProductNo").GetROProperty("innertext")
			'msgbox ActProductNo
		End If	
		If TirePartNo = ActPartNo or TireProductNo = ActProductNo  Then
			Reporter.ReportEvent micPass,"Portal search Result","Part/Product Number is as expected"
		Else
			Reporter.ReportEvent micFail,"Portal search Result","Part/Product Number is Not as expected"		
		End If
		'Validation for Negative search
		If Browser("Products").Page("Products").WebElement("ActErrMessage").Exist Then
		   ActErrMessage = Browser("Products").Page("Products").WebElement("ActErrMessage").GetROProperty("innertext")		   
			If ValidMessage = ActErrMessage  Then
				Reporter.ReportEvent micPass,"Portal Negative search Result","Validation message is as expected"   
		   	Else
		   		Reporter.ReportEvent micFail,"Portal Negative search Result","Validation message is Not as expected"   
		   	End If
		End If
    End If
Next









