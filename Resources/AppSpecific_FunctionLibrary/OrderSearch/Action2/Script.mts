'Script Name     - OrderSearch
'Description     - Script implements basic and advanced search for inventory
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
Dim OrderNumber
Dim PONumber
Dim ShipTO
Dim BillTo	
Dim OrderDate
Dim BillToFromApplication
Dim Mspn
Dim OrderFromDate
Dim OrderToDate
Dim ShipToFilter
Dim ScenarioType

Repositoriescollection.Add environment.value("varpathLogin")&"ObjectRepository\OrderSearch.tsr"
scriptpathLogin = environment("ScriptPath1")
environment.value("varpathLogin") = Mid(scriptpathLogin,1,Instrrev(Mid(scriptpathLogin,1,instrrev(scriptpathLogin,"\")-1),"\"))
varpath1 = environment.value("varpathLogin")
varpath1 = Mid(scriptpathLogin,1,Instrrev(Mid(varpath1,1,instrrev(varpath1,"\")-1),"\"))
environment.value("varpath1") = varpath1
Datatable.AddSheet "OrderSearch"
Datatable.ImportSheet environment.value("varpath1")&"TestData\OrderSearch.xlsx","Sheet1","OrderSearch"
RowCount = Datatable.GetSheet("OrderSearch").GetRowCount

For i = 1 To RowCount
Do
	Datatable.SetCurrentRow(i)
	RunTest = Datatable.GetSheet("OrderSearch").GetParameter("Run")
		
	'Fetch Values from the datasheet
	OrderNumber=DataTable.Value("OrderNumber","OrderSearch")
	PONumber=DataTable.Value("PONumber","OrderSearch")
	ShipTO=DataTable.Value("ShipTo","OrderSearch")
	BillTo=DataTable.Value("BillTo","OrderSearch")
	OrderDate=DataTable.Value("OrderDate","OrderSearch")
	Scenario=DataTable.Value("Scenario","OrderSearch")
	ScenarioType=DataTable.Value("ScenarioType","OrderSearch")

	If RunTest = "Yes" Then
		'Wait for page to load
		Browser("Portal").Page("Orders").Sync
		'Check if page is loaded properly
		If Browser("Portal").Page("Orders").Link("Orders").Exist(conExistTimeout) Then
			Browser("Portal").Page("Orders").Link("Orders").Click
			Wait(5)
			'Get the Bill To Value from teh application for verification
			BillToFromApplication=Browser("Portal").Page("Orders").WebElement("BillToFromApplication").GetROProperty("innertext")
			BillToFromApplication=Trim(Split(BillToFromApplication,"-")(1))
		Else	
			AddNewCase strTCID,""&Scenario&" : Page Loading","The portal page should be loaded correctly","The portal page is not loaded correctly","Fail"
			ExitTest	
		End If
		
		'Enter the Bill TO detail on top
		Browser("Portal").Page("Orders").WebElement("BillTOClickable").Click
		Wait 2
		Browser("Portal").Page("Orders").WebEdit("BillToSearchField").Set BillTo
		Wait 2
		If Browser("Portal").Page("Orders").WebElement("AutoCompleteSearchResultMismatch").Exist(2) Then
			If ScenarioType="Positive" Then
				AddNewCase strTCID,""&Scenario&" : Bill Number Entry Validation","The Bill TO number entered should have matcheing results","The Bill TO number entered has no matches","Fail"
				Exit Do
			Else
				AddNewCase strTCID,""&Scenario&" : Bill Number Entry Validation with invalid value","The Bill TO number entered should not have matcheing results","The Bill TO number entered has no matches","Pass"
				Exit Do
			End If
		Else
			Wait 2
			Browser("Portal").Page("Orders").WebElement("BillTOSearchResult").Click
		End If
		
		'Enter the Ship TO detail on top
		Browser("Portal").Page("Orders").WebElement("ShipTOClickable").Click
		Wait 2
		Browser("Portal").Page("Orders").WebEdit("ShipToSearchField").Set DataTable.Value("ShipToFilter","OrderSearch")
		If Browser("Portal").Page("Orders").WebElement("AutoCompleteSearchResultMismatch").Exist(2) Then
			If ScenarioType="Positive" Then
				AddNewCase strTCID,""&Scenario&" : Ship TO number Entry Validation","The Ship TO number entered should have matcheing results","The Ship TO number entered has no matches","Fail"
				Exit Do
			Else
				AddNewCase strTCID,""&Scenario&" : Ship TO number Entry Validation with invalid value","The Ship TO number entered should not have matching results","The Ship TO number entered has no matches","Pass"
				Exit Do			
			End If
			
		Else
			Wait 2
			Browser("Portal").Page("Orders").WebElement("ShipToSearchResult").Click
		End If		
		'Basic Search
		
		'Set the search string
		If OrderNumber<>"" Then
			'Set the search string as OrderNumber
			Browser("Portal").Page("Orders").WebEdit("Search Orders").Set OrderNumber
			'Click on the search button
			Browser("Portal").Page("Orders").Link("SEARCH").Click
			Wait(2)
			'Check if results are displayed, if not quit
			If CheckIfSearchResultsAreDisplayed<>1 Then	
				'Get the value of the Order Number from the application
				OrderNumberActual=Browser("Portal").Page("Orders").WebTable("Order Date").GetCellData(2,5) @@ hightlight id_;_65730_;_script infofile_;_ZIP::ssf2.xml_;_
				'Compare the actual and expected values of Order Number
				If OrderNumber=OrderNumberActual Then
					AddNewCase strTCID,""&Scenario&" : -Order Number Verification","The order number in the search results should be as per the value entered in the search field. Expected Value : "&OrderNumber,"The order number in the search results is as per the value entered in the search field. Actual Value : "&OrderNumberActual,"Pass"
				Else
					AddNewCase strTCID,""&Scenario&" : -Order Number Verification","The order number in the search results should be as per the value entered in the search field Expected Value : "&OrderNumber,"The order number in the search results is not as per the value entered in the search field. Actual Value : "&OrderNumberActual,"Fail"
				End If
				'Get the value of the Order Number from the application
				BillToSearchResults=Browser("Portal").Page("Orders").WebTable("Order Date").GetCellData(2,2)
				'Verify the BillTo Number
				If BillToFromApplication=BillToSearchResults Then
					AddNewCase strTCID,""&Scenario&" : -Bill TO Number Verification","The Bill TO Number in the search results should be matching the value from the Bill TO dropdown. Expected Value : "&BillToFromApplication,"The Bill TO Number in the search results matches the value from the Bill TO dropdown. Actual Value : "&BillToSearchResults,"Pass"
				Else
					AddNewCase strTCID,""&Scenario&" : -Bill TO Number Verification","The Bill TO Number in the search results should be matching the value from the Bill TO dropdown. Expected Value : "&BillToFromApplication,"The Bill TO Number in the search results does not match the value from the Bill TO dropdown. Actual Value : "&BillToSearchResults,"Fail"
				End If
				Wait(10)
			Else
				Exit Do	
			End If
		Elseif PONumber<>"" Then
			'Set the search string as OrderNumber
			Browser("Portal").Page("Orders").WebEdit("Search Orders").Set PONumber
			'Click on the search button
			Browser("Portal").Page("Orders").Link("SEARCH").Click
			Wait(2)
			'Check if results are displayed, if not quit
			If CheckIfSearchResultsAreDisplayed<>1 Then
				'Get the value of the Order Number from the application
				PONumberActual=Browser("Portal").Page("Orders").WebTable("Order Date").GetCellData(2,4) @@ hightlight id_;_65730_;_script infofile_;_ZIP::ssf2.xml_;_
				'Compare the actual and expected values of Order Number
				If PONumber=PONumberActual Then
					AddNewCase strTCID,""&Scenario&" : -PO Number Verification","The PO number in the search results should be as per the value entered in the search field. Expected Value : "&PONumber,"The PO number in the search results is as per the value entered in the search field. Actual Value : "&PONumberActual,"Pass"
				Else
					AddNewCase strTCID,""&Scenario&" : -PO Number Verification","The PO number in the search results should be as per the value entered in the search field. Expected Value : "&PONumber,"The PO number in the search results is not as per the value entered in the search field. Actual Value : "&PONumberActual,"Fail"
				End If
				'Get the value of the Order Number from the application
				BillToSearchResults=Browser("Portal").Page("Orders").WebTable("Order Date").GetCellData(2,2)
				'Verify the BillTo Number
				If BillToFromApplication=BillToSearchResults Then
					AddNewCase strTCID,""&Scenario&" : -Bill TO Number Verification","The Bill TO Number in the search results should be matching the value from the Bill TO dropdow. Expected Value : "&BillToFromApplication,"The Bill TO Number in the search results matches the value from the Bill TO dropdown. Actual Value : "&BillToSearchResults,"Pass"
				Else
					AddNewCase strTCID,""&Scenario&" : -Bill TO Number Verification","The Bill TO Number in the search results should be matching the value from the Bill TO dropdown. Expected Value : "&BillToFromApplication,"The Bill TO Number in the search results does not match the value from the Bill TO dropdown. Actual Value : "&BillToSearchResults,"Fail"
				End If
			Else
				Exit Do	
			End If
		Elseif ShipTO<>"" Then
			'Set the search string as OrderNumber
			Browser("Portal").Page("Orders").WebEdit("Search Orders").Set ShipTO
			'Click on the search button
			Browser("Portal").Page("Orders").Link("SEARCH").Click
			Wait(2)
			'Check if results are displayed, if not quit
			If CheckIfSearchResultsAreDisplayed<>1 Then
				'Get Rowcount of the webtable
				RowC=Browser("Portal").Page("Orders").WebTable("Order Date").RowCount
				'Iterate all the search results
				For TabIndex = 2 To RowC Step 1
					'Get the value of the Order Number from the application
					ShipTOActual=Browser("Portal").Page("Orders").WebTable("Order Date").GetCellData(TabIndex,3) @@ hightlight id_;_65730_;_script infofile_;_ZIP::ssf2.xml_;_
					'Compare the actual and expected values of Order Number
					If ShipTO=ShipTOActual Then
						AddNewCase strTCID,""&Scenario&" : -ShipTO Number Verification","The ShipTO number in the search result should be as per the value entered in the search field. Expected Value : "&ShipTO,"The ShipTO number in the search results is as per the value entered in the search field. Actual Value : "&ShipTOActual,"Pass"
					Else
						AddNewCase strTCID,""&Scenario&" : -ShipTO Number Verification","The ShipTO number in the search result  should be as per the value entered in the search field. Expected Value : "&ShipTO,"The ShipTO number in the search results is not as per the value entered in the search field. Actual Value : "&ShipTOActual,"Fail"
					End If
					'Get the value of the Order Number from the application
					BillToSearchResults=Browser("Portal").Page("Orders").WebTable("Order Date").GetCellData(TabIndex,2)
					'Verify the BillTo Number
					If BillToFromApplication=BillToSearchResults Then
						AddNewCase strTCID,""&Scenario&" : -Bill TO Number Verification","The Bill TO Number in the search result should be matching the value from the Bill TO dropdown. Expected Value : "&BillToFromApplication,"The Bill TO Number in the search results matches the value from the Bill TO dropdown. Actual Value : "&BillToSearchResults,"Pass"
					Else
						AddNewCase strTCID,""&Scenario&" : -Bill TO Number Verification","The Bill TO Number in the search result should be matching the value from the Bill TO dropdown. Expected Value : "&BillToFromApplication,"The Bill TO Number in the search results does not match the value from the Bill TO dropdown. Actual Value : "&BillToSearchResults,"Fail"
					End If
				Next
			Else
				Exit Do	
			End If
		End If
		
		'Advanced Search
		'Click on the Advance search link
		Browser("Portal").Page("Orders").Link("Advance Search").Click
		'Fetch values for advance search from datatable
		Mspn=DataTable.Value("MSPN","OrderSearch")
		OrderFromDate=DataTable.Value("OrderFromDate","OrderSearch")
		OrderToDate=DataTable.Value("OrderToDate","OrderSearch")
		'Determine the scenario
		If Mspn<>"" And OrderFromDate="" Then
			ScenarioSwitcher=1
		ElseIf Mspn="" And OrderFromDate<>"" Then
			ScenarioSwitcher=2
		ElseIf Mspn<>"" And OrderFromDate<>"" Then
			ScenarioSwitcher=3	
		End If
		
		Select Case ScenarioSwitcher
			Case 1
				'Set Search string as MSPN
				Browser("Portal").Page("Orders").WebEdit("MSPN #").Set Mspn
				'Click on Search link
				Browser("Portal").Page("Orders").Link("SEARCH").Click
				Wait(2)
				'Check if results are displayed, if not quit
				If CheckIfSearchResultsAreDisplayed<>1 Then
					Wait(5)
					RowC=Browser("Portal").Page("Orders").WebTable("Order Date").RowCount
					'Iterate all the search results
					For TabIndex = 1 To RowC-1 Step 1
						'Click on the Order link
						Set oDesc = Description.Create
						oDesc("micclass").value = "Link"
						oDesc("class").value = "clickableLink"
						oDesc("visible").value = "True"
						'Find all the Links
						Set obj = Browser("Portal").Page("Orders").ChildObjects(oDesc)
						'For i = 1 to obj.Count - 1	
						'MsgBox obj(i).GetROProperty("innertext")
						'---------------------DUE TO COMPATIBILITY ISSUE IN IE
						'If TabIndex=2 Then
								Wait(5)
								obj(TabIndex-1).Click
	
							
							'Next
							'Browser("Portal").Page("Orders").WebTable("Order Date").ChildItem(TabIndex,5,"Link",0).Click
							'Verify MSPN from new screen
							Wait(5)
							MspnActual=Browser("Portal").Page("Order Lines").Link("mspn").GetROProperty("innertext")
							If Mspn=MspnActual Then
								AddNewCase strTCID,""&Scenario&" : -MSPN Verification","The MSPN number in the search result  should be as per the value entered in the search field. Expected Value : "&Mspn,"The MSPN number in the search results is as per the value entered in the search field. Actual Value : "&MspnActual,"Pass"
							Else
								AddNewCase strTCID,""&Scenario&" : -MSPN Number Verification","The MSPN number in the search result should be as per the value entered in the search field. Expected Value : "&Mspn,"The MSPN number in the search results is not as per the value entered in the search field. Actual Value : "&MspnActual,"Fail"
							End If
							
							'Click on MSPN link
							Browser("Portal").Page("Order Lines").Link("mspn").Click
							'Wait for Order Schedules page
							Browser("Portal").Page("Order Schedules").Sync
							Wait(5)
							'Check shipping status
							If Browser("Portal").Page("Order Schedules").WebElement("Shipping Status").GetROProperty("innertext")<>"" And Browser("Portal").Page("Order Schedules").WebElement("Shipping Status").GetROProperty("innertext")<>"Cancelled" Then
								strDispatchDate=Browser("Portal").Page("Order Schedules").WebElement("Dispatch Date").GetROProperty("innertext")
								If strDispatchDate<>"" Then
									AddNewCase strTCID,""&Scenario&" : - Shipping status and Dispatch date verification","Shipping status and Dispatch date should be populated as expected. ","Shipping status and Dispatch date is populated as expected","Pass"
								Else
									AddNewCase strTCID,""&Scenario&" : - Shipping status and Dispatch date verification","Shipping status and Dispatch date should be populated as expected"," Dispatch date is not populated as expected","Fail"
								End If
							Else
									AddNewCase strTCID,""&Scenario&" : - Shipping status and Dispatch date verification","Shipping status and Dispatch date should be populated as expected"," Shipping status is not populated as expected","Fail"						
							End If
	
							'Click on back to order link
							Browser("Portal").Page("Order Schedules").Link("Back to Orders").Click
						
							'Browser("Portal").Page("Order Lines").Link("Back to Orders").Click
							Wait 5
					Next
				Else
					Exit Do
				End If
			Case 2
				'Set Order from and Order to fields
				Browser("Portal").Page("Orders").WebEdit("OrderFrom").Set OrderFromDate
				Browser("Portal").Page("Orders").WebEdit("OrderTo").Set OrderToDate
				'Click on Search link
				Browser("Portal").Page("Orders").Link("SEARCH").Click
				'Check if results are displayed, if not quit
				If CheckIfSearchResultsAreDisplayed<>1 Then
					RowC=Browser("Portal").Page("Orders").WebTable("Order Date").RowCount
					'Iterate all the search results
					For TabIndex = 2 To RowC Step 1
						OrderDateActual=Browser("Portal").Page("Orders").WebTable("Order Date").GetCellData(TabIndex,1)
						If CheckDate(OrderFromDate,OrderToDate,OrderDateActual)=0 Then
							AddNewCase strTCID,""&Scenario&" : -Date Verification","The Order Date in the search result should be between the date range provided in advance search. Expected Date Range: "&OrderFromDate&" to "&OrderToDate,"The Order Date in the search result is between the date range provided in advance search. Actual value :"&OrderDateActual,"Pass"
						Else
							AddNewCase strTCID,""&Scenario&" : -Date Verification","The Order Date in the search result should be between the date range provided in advance search. Expected Date Range: "&OrderFromDate&" to "&OrderToDate,"The Order Date in the search result is not between the date range provided in advance search. Actual value :"&OrderDateActual,"Fail"
						End If
					Next
				Else
					Exit Do	
				End If
			Case 3
				'Set Search string as MSPN
				Browser("Portal").Page("Orders").WebEdit("MSPN #").Set Mspn
				'Set Order from and Order to fields
				Browser("Portal").Page("Orders").WebEdit("OrderFrom").Set OrderFromDate
				Browser("Portal").Page("Orders").WebEdit("OrderTo").Set OrderToDate
				'Click on Search link
				Browser("Portal").Page("Orders").Link("SEARCH").Click
				'Check if results are displayed, if not quit
				If CheckIfSearchResultsAreDisplayed<>1 Then
					Wait(5)
					RowC=Browser("Portal").Page("Orders").WebTable("Order Date").RowCount
						'Iterate all the search results
					For TabIndex = 1 To RowC-1 Step 1
						OrderDateActual=Browser("Portal").Page("Orders").WebTable("Order Date").GetCellData(TabIndex,1)
						If CheckDate(OrderFromDate,OrderToDate,OrderDateActual)=0 Then
							AddNewCase strTCID,""&Scenario&" : -Date Verification","The Order Date in the search result should be between the date range provided in advance search","The Order Date in the search result is between the date range provided in advance search. Actual value :"&OrderDateActual&"  Expected Date Range: "&OrderFromDate&" to "&OrderToDate,"Pass"
						Else
							AddNewCase strTCID,""&Scenario&" : -Date Verification","The Order Date in the search result  should be between the date range provided in advance search","The Order Date in the search result  is not between the date range provided in advance search. Actual value :"&OrderDateActual&"  Expected Date Range: "&OrderFromDate&" to "&OrderToDate,"Fail"
						End If
						'Click on the Order link
						Set oDesc = Description.Create
						oDesc("micclass").value = "Link"
						oDesc("class").value = "clickableLink"
						oDesc("visible").value = "True"
						'Find all the Links
						Set obj = Browser("Portal").Page("Orders").ChildObjects(oDesc)
						'For i = 1 to obj.Count - 1	
						'MsgBox obj(i).GetROProperty("innertext")
						'---------------------DUE TO COMPATIBILITY ISSUE IN IE
						'If TabIndex=2 Then
							Wait(5)	
							obj(TabIndex-1).Click
							'Browser("Portal").Page("Orders").WebTable("Order Date").ChildItem(TabIndex,5,"Link",0).Click
							Wait(5)
							'Verify MSPN from new screen
							MspnActual=Browser("Portal").Page("Order Lines").Link("mspn").GetROProperty("innertext")
							If Mspn=MspnActual Then
								AddNewCase strTCID,""&Scenario&" : -MSPN Verification","The MSPN number in the search result should be as per the value entered in the search field. Expected value : "&Mspn,"The ShipTO number in the search results is as per the value entered in the search field. Actual value : "&MspnActual,"Pass"
							Else
								AddNewCase strTCID,""&Scenario&" : -MSPN Number Verification","The MSPN number in the search result should be as per the value entered in the search field. Expected value : "&Mspn,"The ShipTO number in the search results is not as per the value entered in the search field. Actual value :"&MspnActual,"Fail"
							End If
													
							'Click on MSPN link
							Browser("Portal").Page("Order Lines").Link("mspn").Click
							'Wait for Order Schedules page
							Browser("Portal").Page("Order Schedules").Sync
							Wait(5)
							'Check shipping status
							If Browser("Portal").Page("Order Schedules").WebElement("Shipping Status").GetROProperty("innertext")<>"" And Browser("Portal").Page("Order Schedules").WebElement("Shipping Status").GetROProperty("innertext")<>"Cancelled" Then
								strDispatchDate=Browser("Portal").Page("Order Schedules").WebElement("Dispatch Date").GetROProperty("innertext")
								If strDispatchDate<>"" Then
									AddNewCase strTCID,""&Scenario&" : - Shipping status and Dispatch date verification","Shipping status and Dispatch date should be populated as expected","Shipping status and Dispatch date is populated as expected","Pass"
								Else
									AddNewCase strTCID,""&Scenario&" : - Shipping status and Dispatch date verification","Shipping status and Dispatch date should be populated as expected"," Dispatch date is not populated as expected","Fail"
								End If
							Else
									AddNewCase strTCID,""&Scenario&" : - Shipping status and Dispatch date verification","Shipping status and Dispatch date should be populated as expected"," Shipping status is not populated as expected","Fail"						
							End If
	
							'Click on back to order link
							Browser("Portal").Page("Order Schedules").Link("Back to Orders").Click
							'Click on back to order link
							'Browser("Portal").Page("Order Lines").Link("Back to Orders").Click
							Wait 5
					Next
				Else
					Exit Do
				End If
		End Select
		
    End If
Loop While False
Next

Function CheckIfSearchResultsAreDisplayed()

	'Wait for page to be loaded
	Browser("Portal").Page("Orders").Sync
	'Check if the no results found message appears in the search results
	If Browser("Portal").Page("Orders").WebElement("No results found").Exist(5) Then
		CheckIfSearchResultsAreDisplayed=1
		If ScenarioType="Positive" Then
			AddNewCase strTCID,""&Scenario&" : - Search result verification","Search results should be displayed as expected","Search results are not displayed for the input provided","Fail"
		Else
			AddNewCase strTCID,""&Scenario&" : - Search result verification with invalid search key","Search results should not be displayed","Search results are not displayed for the input provided","Pass"
		End If

	Else
		CheckIfSearchResultsAreDisplayed=0
	End If
End Function



