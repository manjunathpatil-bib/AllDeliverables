'Script Name     - AddProduct
'Description     - AddProduct script add a prodcut to a pricebook and attaches it with the opportunity as per the datasheet values
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
Dim intProductFound

intProductFound=0

environment.value("varpathLogin")=Environment("RootResourceDirectory")
Repositoriescollection.Add environment.value("varpathLogin")&"ObjectRepository\Oppurtunities\AddProduct.tsr"
varpath1=Environment("RootScriptDirectory")
environment.value("varpath1") = varpath1
Datatable.AddSheet "AddProduct"
Datatable.ImportSheet environment.value("varpath1")&"TestData\Oppurtunities\Addproduct.xlsx","Sheet1","AddProduct"
RowCount = Datatable.GetSheet("AddProduct").GetRowCount

For i = 1 To RowCount
	Do
		Datatable.SetCurrentRow(i)
		RunTest = Datatable.GetSheet("AddProduct").GetParameter("Run")
			
		'Fetch Values from the datasheet
		OpportunityName=DataTable.Value("OpportunityName","AddProduct")
		Scenario=DataTable.Value("Scenario","AddProduct")
		Product=DataTable.Value("Product","AddProduct")
		strCurrency=DataTable.Value("Currency","AddProduct")
		ListPrice=DataTable.Value("ListPrice","AddProduct")
		UseStandardPriceorNot=DataTable.Value("UseStandardPriceorNot","AddProduct")
		
		If RunTest = "Yes" Then
			'Wait for page to load properly
			Browser("Home | Salesforce").Page("Home | Salesforce").Sync
			Wait 3
			'Click on Oppurtunities link
			ClickonOpportunitiesLink
			Wait 5 
			'Click on the oppurtunityname mentioned in data table
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").Link("OppurtunityName").SetTOProperty "text",OpportunityName
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").Link("OppurtunityName").Click
			Wait 5 
			'Click on pricebook
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").Link("priceBookClick").Click
			Wait 5 	
			'Click on Related tab
			Browser("Opportunities | Salesforce").Page("PBook1 | Salesforce").WebElement("Related").Click
			Wait 5 	
			'Check if the product to add is already added
			Browser("Opportunities | Salesforce").Page("Products | Salesforce").Link("CheckProductExists").SetTOProperty "text",Product
			If Browser("Opportunities | Salesforce").Page("Products | Salesforce").Link("CheckProductExists").Exist(5) Then
				AddNewCase strTCID,""&Scenario,"User should be able to add the product to the pricebook","Product is already present in the pricebook","Pass"
			Else
				'Click on Add Product button
				Browser("Opportunities | Salesforce").Page("PBook1 | Salesforce").Link("Add Product").Click
				Wait 5 	
				 'Select Product
  	             Browser("Opportunities | Salesforce").Page("PBook1 | Salesforce").WebList("Product").Select Product
				'Select Currency 
				Browser("Opportunities | Salesforce").Page("PBook1 | Salesforce").WebList("Currency").Select strCurrency
				'Click on Next
				Browser("Opportunities | Salesforce").Page("PBook1 | Salesforce").WebButton("Next").Click
				Wait 5 
				'Set the List Price
				Browser("Opportunities | Salesforce").Page("PBook1 | Salesforce").WebEdit("List Price").Set ListPrice
				'Use standard price or not
				If UseStandardPriceorNot="Yes" Then
					Browser("Opportunities | Salesforce").Page("PBook1 | Salesforce").WebCheckBox("UseStandardPrice").Set "ON"
				End If
				'Click on Save
				Browser("Opportunities | Salesforce").Page("PBook1 | Salesforce").WebButton("Save").Click
			End If
			ClickonOpportunitiesLink
			Wait 5 
			'Click on the oppurtunityname mentioned in data table
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").Link("OppurtunityName").SetTOProperty "text",OpportunityName
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").Link("OppurtunityName").Click
			Wait 5 
			'Click on Sales detail
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").Link("Sales Detail").Click
			Wait 5 
			'Click on Add Product
			Browser("Opportunities | Salesforce").Page("Products Custom Screen").Frame("Products Custom Screen").WebButton("Add Products").Click
			Wait 5
			'Click on search button in the search and add products frame
			Browser("Opportunities | Salesforce").Page("Search and add products").Frame("Search and Add Product").WebButton("Search").Click
			Wait 10
			If Browser("Opportunities | Salesforce").Page("Search and add products").Frame("Search and Add Product").WebElement("ResultsNumber").Exist(0)=False Then
				'Click on the seaerch result checkbox
				Browser("Opportunities | Salesforce").Page("Search and add products").Frame("Search and Add Product").WebElement("SearchResultProductCheckbox").Click
				'Click on Add products button	
				Browser("Opportunities | Salesforce").Page("Search and add products").Frame("Search and Add Product").WebButton("Add Products").Click
				Wait 10
				'intTabCountProduct=Browser("Opportunities | Salesforce").Page("Search and add products").Frame("Search and Add Product").WebTable("ProductSearchResults").GetRowCount
				intTabCountProduct=Browser("Opportunities | Salesforce").Page("Search and add products").Frame("Search and Add Product").WebTable("ProductSearchResults").GetROProperty("rows")
				For intTabIter = 1 To intTabCountProduct Step 1
					strTabProductName=Browser("Opportunities | Salesforce").Page("Search and add products").Frame("Search and Add Product").WebTable("ProductSearchResults").GetCellData(intTabIter,2)
					If strTabProductName=Product Then
						AddNewCase strTCID,""&Scenario,"Product "&Product&" has been added to the opportunity "&OpportunityName,"Product "&Product&" has not been added to the opportunity "&OpportunityName,"Fail"
						Exit Do
					End If
				Next
			Else
				AddNewCase strTCID,""&Scenario,"Product "&Product&" has been added to the opportunity "&OpportunityName,"Product "&Product&" is not available in the PriceBook","Fail"			
			End If
			Wait 5
			AddNewCase strTCID,""&Scenario,"Product "&Product&" should be added to the opportunity "&OpportunityName,"Product "&Product&" has been added to the opportunity "&OpportunityName,"Pass"
	    End If
	Loop While False
Next


