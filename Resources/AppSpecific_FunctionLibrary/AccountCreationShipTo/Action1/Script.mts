'-----------------------------------------------------------------------------------------------------------------
'Script Name  - AccountCreationShipTo
'Description  - Account Creation ShipTo
'Created By   -
'Created On   -
'Modified By  -
'Modified On  -
'Authour      - CGI
'-----------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------
'Environment setup

scriptpathShipTo = environment("ScriptPath1")

'scriptpath = environment.value("TestDir")
environment.value("varpathShipTo") =Mid(scriptpathShipTo,1,Instrrev(Mid(scriptpathShipTo,1,instrrev(scriptpathShipTo,"\")-1),"\"))

'Describes the ObjectRepository path
Repositoriescollection.Add environment.value("varpathShipTo")&"\ObjectRepository\AccountCreation.tsr"


varpathShipTo = environment.value("varpathShipTo")
varpathShipTo = Mid(scriptpathShipTo,1,Instrrev(Mid(varpathShipTo,1,instrrev(varpathShipTo,"\")-1),"\"))
environment.value("varpathShipTo") = varpathShipTo

'Describes TestData Sheet path
Datatable.AddSheet "Sheet1"
Datatable.ImportSheet environment.value("varpathShipTo")&"\TestData\AccountCreationShipTo.xlsx","Sheet1","Sheet1"
'Datatable.getsheet("Sheet1").SetCurrentRow 1
RowCount = Datatable.GetSheet("Sheet1").GetRowCount


For k = 1 To RowCount
Datatable.SetCurrentRow(k)
RunTest = Datatable.GetSheet("Sheet1").GetParameter("Run")	
	
	If RunTest = "Yes" Then
		
			'Incrementing BilltoAccName
			BilltoAccName = datatable.Value("BilltoAccName","Sheet1")
			BilltoAccNamearr = split(BilltoAccName,"_")
			BilltoAccNamearr(ubound(BilltoAccNamearr)) =cstr( cint(BilltoAccNamearr(ubound(BilltoAccNamearr)))+1)
			BillToAccNM = BilltoAccNamearr(0)
			For i=1  to  Ubound(BilltoAccNamearr)
				BillToAccNM = BillToAccNM&"_"&BilltoAccNamearr(i)
			Next
			datatable.Value("BilltoAccName","Sheet1") = BillToAccNM
			
			'Incrementing ShiptoAccName
			ShiptoAccName = datatable.Value("ShiptoAccName","Sheet1")
			ShiptoAccNamearr = split(ShiptoAccName,"_")
			ShiptoAccNamearr(ubound(ShiptoAccNamearr)) =cstr( cint(ShiptoAccNamearr(ubound(ShiptoAccNamearr)))+1)
			ShipToAccNM = ShiptoAccNamearr(0)
			For m=1  to  Ubound(ShiptoAccNamearr)
				ShipToAccNM = ShipToAccNM&"_"&ShiptoAccNamearr(m)
			Next
			datatable.Value("ShiptoAccName","Sheet1") = ShipToAccNM
			
			'Intialyzing the values
			BusinessUnit  = datatable.Value("BusinessUnit","Sheet1")
			Channel		= datatable.Value("Channel","Sheet1")
			SubChannel  = datatable.Value("SubChannel","Sheet1")
			Phone		= datatable.Value("Phone","Sheet1")
			Fax			= datatable.Value("FAX","Sheet1")   
			CommercialRelation = datatable.Value("CommercialRelation","Sheet1")
			AccountCurrency = datatable.Value("AccountCurrency","Sheet1")
			Website				= datatable.Value("Website","Sheet1")
			SalesAgreementCodes = datatable.Value("SalesAgreementCodes","Sheet1")
			Street    = datatable.Value("Street","Sheet1")
			City = datatable.Value("City","Sheet1")
			Zip = datatable.Value("Zip","Sheet1")
			State = datatable.Value("State","Sheet1")
			Country = datatable.Value("Country","Sheet1")
			CreditRating = datatable.Value("CreditRating","Sheet1")
			
			'Incrementing the MMID
			MDMID = datatable.Value("MDMID","Sheet1")
			MDMID = Cstr(cdbl(MDMID)+1)
			MDMIDShipto=Cstr(cdbl(MDMID)+1)
			datatable.Value("MDMID","Sheet1") = MDMIDShipto
			'datatable.ExportSheet environment.value("varpath")&"TestData\AccountCreation.xlsx","Sheet1","Sheet1"
			datatable.ExportSheet environment.value("varpathShipTo")&"TestData\AccountCreationShipTo.xlsx","Sheet1","Sheet1"
			SrcCountry = datatable.Value("SrcCountry","Sheet1")
			
			'Account
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").Link("lnk_MainWinAccounts").Click
			wait(2)
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").Link("lnk_AccountNew").Click
			wait 2
			
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("Delivery Group").Click
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Next").Click @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Next")_;_script infofile_;_ZIP::ssf3.xml_;_
			
 @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Next")_;_script infofile_;_ZIP::ssf5.xml_;_
			wait(1)
			
			'ShipTo account creation
			
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_AccNameBillto").Set ShiptoAccName
			'
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").Link("lnk_Channel").Click
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebButton("btn_Channel").Click
			
			wait 3
			'Channel
			'set the user difine list like below 
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_ChannelOptionlist").Link("lnk_ChannerlOptions").SetTOProperty "title",Channel
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_ChannelOptionlist").Link("lnk_ChannerlOptions").Click
			
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebList("lst_BusinessUnit").Select BusinessUnit'"PLNA"
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").Link("Channel_2").Click remove
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").Link("ADVANTAGE").Click remove
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_Phone").Set Phone		'"111-111-2223"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_Fax").Set Fax			' "111-111-2224"
			
			''Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").Link("Commercial Relation_2").Click
			''Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").Link("Commercial Relation_2").Click
			wait 2
			
			'CommercialRelation
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebButton("btn_Commercial Relation").Click
			
			''Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").Link("lnk_CommercialRelation").Object.scrollIntoView
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebButton("btn_Commercial Relation").Object.scrollIntoView
			'
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebButton("btn_Commercial Relation").Click
			''Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").Link("lnk_CommercialRelation").Click
			'wait 2
			
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_CommercialRelationlistblank").Link("lnk_listoptions").SetTOProperty "title",CommercialRelation @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").Link("ADVANTAGE")_;_script infofile_;_ZIP::ssf11.xml_;_
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_CommercialRelationlistblank").Link("lnk_listoptions").Click
			
			'Website
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").Link("Direct").Click
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_WebSite").Set Website	
			wait 2
			
			''Account Currency
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebButton("btn_Account Currency").Click
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_AccountCurrencylist").Link("lnk_listoptions").SetTOProperty "title", AccountCurrency
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_AccountCurrencylist").Link("lnk_listoptions").Click
			
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_SalesAgreementCode1").Set SalesAgreementCodes'"R123489"
			
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_SalesAgreementCode").Set SalesAgreementCodes '"R123489"
			
			'passing address
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_BillingStreet").Object.scrollIntoView
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_BillingStreet").Set Street    '"1 Parkway South"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_BillingCity").Set City '"Greenville"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_BillingState").Set State ' "SC"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_Billingzip").Set Zip ' "29615"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_BillingCountry").Set Country '"US"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_ShippingStreet").Set Street' "1 Parkway South"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_ShippingCity").Set City'"GreenVille"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_Shippingstate").Set State'"SC"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_Shippingzip").Set Zip'"29615"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_ShippingCoutry").Set Country '"US"
			
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").Link("lnk_CreditRating").Object.scrollIntoView
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebButton("btn_Credit Rating").Object.scrollIntoView
			
			'Credit Rating
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebButton("btn_Credit Rating").Click
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").Link("lnk_CreditRating").Click
			
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_CreditRatingOptionlist").Link("lnk_listoptions").SetTOProperty "title","Great"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_CreditRatingOptionlist").Link("lnk_listoptions").Click
			
			'MMID
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebEdit("txt_Shippingzip").Set MDMIDShipto'"9999996"
			
			'SoucreCountry
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").Link("lnk_SourceCountry").Click
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountDelivery").WebButton("btn_Source Country").Click
			
			
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_SourceCountryOptionlist").Link("lnk_listoptions").SetTOProperty "title", "North America"
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_SourceCountryOptionlist").Link("lnk_listoptions").Click
			
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Save_2").Click
			''Browser("Accounts | Salesforce").Page("AccNAme_AutoPOC_Ship10_2").Link("name:="&BilltoAccName).Click
			
			wait 4

	End If

Next
