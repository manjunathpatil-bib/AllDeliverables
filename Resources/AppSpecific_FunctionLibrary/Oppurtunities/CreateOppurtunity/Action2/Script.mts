'Script Name     - CreateOpportunity
'Description     - Action creates an opportunity based on the details from the datasheet
'Created By      -
'Created On      -
'Modified By     -
'Modified On     -
'Authour         - CGI
'-------------------------------------------------------------------------------------------------------------
'Entry point - Script clicks on opportunities link in the Sales force home page and starts the flow
'Exit point - Script Creates an opportunity based on values mentioned in datasheet and verifies if the opportunities are created or not
'-------------------------------------------------------------------------------------------------------------

'Environment Setup

Dim RowCount
Dim RunTest

environment.value("varpathLogin")=Environment("RootResourceDirectory")
Repositoriescollection.Add environment.value("varpathLogin")&"ObjectRepository\Oppurtunities\Oppurtunities.tsr"
varpath1=Environment("RootScriptDirectory")
environment.value("varpath1") = varpath1
Datatable.AddSheet "Create Oppurtunities"
Datatable.ImportSheet environment.value("varpath1")&"TestData\Oppurtunities\Create Oppurtunities.xlsx","Sheet1","Create Oppurtunities"

RowCount = Datatable.GetSheet("Create Oppurtunities").GetRowCount

For i = 1 To RowCount
	Do
		Datatable.SetCurrentRow(i)
		RunTest = Datatable.GetSheet("Create Oppurtunities").GetParameter("Run")
			
		'Fetch Values from the datasheet
		RecordType=DataTable.Value("RecordType","Create Oppurtunities")
		OpportunityName=DataTable.Value("OpportunityName","Create Oppurtunities")
		AccountName=DataTable.Value("AccountName","Create Oppurtunities")
		CloseDate=DataTable.Value("CloseDate","Create Oppurtunities")
		Probability=DataTable.Value("Probability","Create Oppurtunities")
		PriceBook=DataTable.Value("PriceBook","Create Oppurtunities")
		PriceBookName=DataTable.Value("PriceBookName","Create Oppurtunities")
		PriceBookDescription=DataTable.Value("PriceBookDescription","Create Oppurtunities")
		PriceBookExternalID=DataTable.Value("PriceBookExternalID","Create Oppurtunities")
		Scenario=DataTable.Value("Scenario","Create Oppurtunities")
			
		If RunTest = "Yes" Then		
			'Wait for page to load properly
			Browser("Home | Salesforce").Page("Home | Salesforce").Sync
			Wait 5 
			'Clicn ok App Launcher
'			Browser("Home | Salesforce").Page("Home | Salesforce").WebButton("App Launcher").Click
'			Wait 3
'			Browser("App Launcher | Salesforce").Page("App Launcher | Salesforce").Link("Opportunities").Click
			
			'Click on Oppurtunities link
			ClickonOpportunitiesLink
'			If Browser("Home | Salesforce").Page("Home | Salesforce").Link("Opportunities").Exist(conExistTimeout) Then
'				Browser("Home | Salesforce").Page("Home | Salesforce").Link("Opportunities").Click
'			End If
'			Wait 5
			'Wait for page to load completly
			Browser("Home | Salesforce").Page("Opportunities | Salesforce").Sync
			'Click on the New button
			If Browser("Home | Salesforce").Page("Opportunities | Salesforce").Link("New").Exist(conExistTimeout) Then
				Browser("Home | Salesforce").Page("Opportunities | Salesforce").Link("New").Click
			End If
			Wait 5
			'Wait for page to load 
			Browser("Home | Salesforce").Page("Opportunities | Salesforce").Sync
			'Click on radio button based on value from datasheet
			If RecordType="TCAR" Then
				Browser("Home | Salesforce").Page("Opportunities | Salesforce").WebElement("TCAR").Click
			ElseIf RecordType="PLNA" Then			
				Browser("Home | Salesforce").Page("Opportunities | Salesforce").WebElement("PLNA").Click
			End If
			'Click on Next button
			Browser("Home | Salesforce").Page("Opportunities | Salesforce").WebButton("Next").Click
			'Wait for new frame to load
			If Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("OpportunityName").Exist(conExistTimeout) Then
				Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("OpportunityName").Set OpportunityName
			End If
			'Set the account name
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("AccountNames").Set AccountName
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("AccountNames").Click
			Set WshShell = CreateObject("WScript.Shell")
			wait 3
			WshShell.SendKeys "{ENTER}"	
			Set WshShell = Nothing
			'Click on Account Name autocomplete link
			Wait 5
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").Link("AccName_Link").SetTOProperty "title",AccountName
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").Link("AccName_Link").Click
			'Enter Close Date
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("CloseDate").Set CloseDate
			'Enter the probability
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("Probability").Set Probability
			'Set stage value
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebButton("Stage").Click
			Wait(5)
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.SendKeys "{DOWN}"
			wait 3
			WshShell.SendKeys "{DOWN}"
			wait 3
			WshShell.SendKeys "{ENTER}"	
			Set WshShell = Nothing
			'Set priceBook
			'If value in datatable is createnew, goto createnew flow else select appropriate pricebook
			If PriceBook="CreateNew" Then
				Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("Search Price Books").Set " "
				Wait 4
				Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebElement("New Price Book").Click
				Wait 5
				If Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("PriceBookDescription").Exist(conExistTimeout) Then
					Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("PriceBookName").Set PriceBookName
					Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("PriceBookDescription").Set PriceBookDescription
					Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("PriceBookExternalID").Set PriceBookExternalID
					Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebCheckBox("PriceBookActive").Set "ON"
					Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebButton("PriceBookSave").Click 
					Wait 5
				End If
			Else	
				Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebEdit("Search Price Books").Set PriceBook
				Wait 5
				Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebElement("SearchBookAutoComplete").Click
			End If
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebButton("Save").Click
			Wait 5
			If Browser("Opportunities | Salesforce").Page("OppurtunityMainPage").Image("Opportunity").Exist(conExistTimeout) Then
				If Trim(Browser("Opportunities | Salesforce").Page("OppurtunityMainPage").WebElement("OppurtunityName").GetROProperty("innertext"))=Trim(OpportunityName) Then
					AddNewCase strTCID,""&Scenario,"Opportunity Creation should be successful","Opportunity "&OpportunityName&" Creation is successful","Pass"
				Else
					AddNewCase strTCID,""&Scenario,"Opportunity Creation should be successful","Created Opportunity name doesnt match the entererd value. Expected Value : "&OpportunityName&" Actual Value : "&Browser("Opportunities | Salesforce").Page("OppurtunityMainPage").WebElement("OppurtunityName").GetROProperty("innertext"),"Fail"
				End If
			Else
				AddNewCase strTCID,""&Scenario,"Opportunity Creation should be successful","Opportunity Creation is unsuccessful","Fail"
			End If
	    End If
	Loop While False
Next


