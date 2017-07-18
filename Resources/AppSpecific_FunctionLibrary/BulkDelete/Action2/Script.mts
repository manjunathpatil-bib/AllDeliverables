'Script Name     - BulkDelete
'Description     - Script is called to delete the accounts in bulk
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
Dim ScenarioType
Dim Filt(20)
Dim arrFiltSplit(20)


environment.value("varpathLogin")=Environment("RootResourceDirectory")
Repositoriescollection.Add environment.value("varpathLogin")&"ObjectRepository\BulkDelete.tsr"
varpath1=Environment("RootScriptDirectory")
environment.value("varpath1") = varpath1
Datatable.AddSheet "BulkDelete"
Datatable.ImportSheet environment.value("varpath1")&"TestData\BulkDelete.xlsx","Sheet1","BulkDelete"

'Datatable.getsheet("Sheet1").SetCurrentRow 1
RowCount = Datatable.GetSheet("BulkDelete").GetRowCount



For i = 1 To RowCount
Do
Datatable.SetCurrentRow(i)
RunTest = Datatable.GetSheet("BulkDelete").GetParameter("Run")

	
'Fetch Values from the datasheet
Filt(1)=DataTable.Value("Filter1","BulkDelete")
Filt(2)=DataTable.Value("Filter2","BulkDelete")
Filt(3)=DataTable.Value("Filter3","BulkDelete")
Filt(4)=DataTable.Value("Filter4","BulkDelete")
Filt(5)=DataTable.Value("Filter5","BulkDelete")
DeleteClosedWinOpportunities=DataTable.Value("DeleteClosedWinOpportunities","BulkDelete")
PermanentDeleteRecords=DataTable.Value("PermanentDeleteRecords","BulkDelete")
DeleteOpportunitiesFromOtherOwners=DataTable.Value("DeleteOpportunitiesFromOtherOwners","BulkDelete")

	If RunTest = "Yes" Then
		'Wait for page to load
		Browser("Home | Salesforce").Page("Home | Salesforce").Sync
		Wait 5 
		If Browser("Home | Salesforce").Page("Home | Salesforce").WebElement("Setup").Exist(conExistTimeout) Then
			'Click on Setting icon
			Browser("Home | Salesforce").Page("Home | Salesforce").WebElement("Setup").Click
			Wait 3
			'Click on Setup link
			Browser("Home | Salesforce").Page("Home | Salesforce").Link("Setup").Click
			Wait 2
			Browser("Home | Salesforce").Close
			'Wait for Setup home page to load
			Browser("Setup Home | Salesforce").Page("Setup Home | Salesforce").Sync
			'Check if Setup Webbutton is present or not
			If Browser("Setup Home | Salesforce").Page("Setup Home | Salesforce").WebButton("Setup").Exist(conExistTimeout) Then
				'Enter the search key in the quick find box
				Browser("Setup Home | Salesforce").Page("Setup Home | Salesforce").WebEdit("Quick Find").Set "Mass Delete Records"
				Browser("Setup Home | Salesforce").Page("Setup Home | Salesforce").WebEdit("Quick Find").Click
				Wait 3
				'Click on Mass Delete Records link
				Browser("Setup Home | Salesforce").Page("Setup Home | Salesforce").Link("Mass Delete Records").Click
				Wait 3
				'Click on the search result
				Browser("Setup Home | Salesforce").Page("Setup Home | Salesforce").Frame("Frame1").Link("Mass Delete Accounts Search Link").Click
				Wait 5 
				'Check if the FilterTable is loaded properly
				If Browser("Mass Delete Records Page").Page("Mass Delete Records").Frame("vfFrameId").WebTable("FilterTable").Exist(conExistTimeout) Then
					'Find how many filters are provided in the datatable
					For Iterator = 1 To 5 Step 1
						If Filt(Iterator)="" Then
							intFiltIndex=Iterator-1
							Exit For
						End If
					Next 
					'Check if atleast 1 filter is given in the datatable
					If intFiltIndex<>0 Then
						'Split the filters and put it in arrays
						For IteratorOne = 1 To intFiltIndex Step 1
							'arrFiltSplit starts from 0
							arrFiltSplit(IteratorOne-1)=Split(Filt(IteratorOne),"|")
						Next
						'Use the split array to populate each filter in the mass bulk records filter page
						For IteratorTwo = 0 To intFiltIndex-1 Step 1
							With Browser("Mass Delete Records Page").Page("Mass Delete Records").Frame("vfFrameId").WebTable("FilterTable")
								.ChildItem(IteratorTwo+1,1,"WebList",0).Select arrFiltSplit(IteratorTwo)(0)
								.ChildItem(IteratorTwo+1,2,"WebList",0).Select arrFiltSplit(IteratorTwo)(1)
								.ChildItem(IteratorTwo+1,3,"WebEdit",0).Set arrFiltSplit(IteratorTwo)(2)
								Wait 2
							End With
						Next
						Wait 5 
						'Check if all Close/Win opportunities need to be deleted
						If DeleteClosedWinOpportunities="Yes" Then
							Browser("Mass Delete Records Page").Page("Mass Delete Records").Frame("vfFrameId").WebCheckBox("closed_opp").Set "ON"
							Wait 5
						End If
						'Check if all the records need to be deleted permanently
						If PermanentDeleteRecords="Yes" Then
							Browser("Mass Delete Records Page").Page("Mass Delete Records").Frame("vfFrameId").WebCheckBox("hardDelete").Set "ON" 
							Wait 5
						End If
						'Check Opportunities from other owners need to be deleted or not
						If DeleteOpportunitiesFromOtherOwners="Yes" Then
							Browser("Mass Delete Records Page").Page("Mass Delete Records").Frame("vfFrameId").WebCheckBox("owner_opp").Set "ON"
							Wait 5 
						End If			
						'Click on search button
						Browser("Mass Delete Records Page").Page("Mass Delete Records").Frame("vfFrameId").WebButton("Search").Click
						Wait 5
						
						If Browser("Mass Delete Records Page").Page("Mass Delete Records").Frame("vfFrameId").WebElement("RecordsNotFound").Exist(5) Then
							AddNewCase strTCID,"Bulk Delete Accounts","User should be able to delete the accounts based on the filter criteria","No accounts found with the given filter criteria ","Pass"
						Else
						'Check all
							Browser("Mass Delete Records Page").Page("Mass Delete Records").Frame("vfFrameId").WebCheckBox("allBox").Set "ON"
							'Click on delete
							Browser("Mass Delete Records Page").Page("Mass Delete Records").Frame("vfFrameId").WebButton("Delete").Click	
							Wait 20
							AddNewCase strTCID,"Bulk Delete Accounts","User should be able to delete the accounts based on the filter criteria","The accounts were successfully deleted ","Pass"
						End If
						
					Else
						AddNewCase strTCID,""&Scenario,"Filters for the mass delete operation should be provided in the datatable","No filters are provided in the datatable","Fail"
					End If				
				Else
					AddNewCase strTCID,""&Scenario,"Mass Delete Accounts page should be loaded successfully","Mass delete records page is not loaded","Fail"
				End If


			Else
				AddNewCase strTCID,""&Scenario,"Setup home page should be loaded successfully","Setup home page is not loaded","Fail"
			End If
		Else
			AddNewCase strTCID,""&Scenario,"Home page should be loaded successfully","Home page is not loaded","Fail"
		End If
		
    End If
Loop While False
Next
