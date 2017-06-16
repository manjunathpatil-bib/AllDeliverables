'Script Name     - AccountCreationBillTo
'Description     - Account Creation BillTo
'Created By      -
'Created On      -
'Modified By     -
'Modified On     -
'Authour         - CGI
'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------

'Environment Setup

scriptpathBillTo = environment("ScriptPath1")

environment.value("varpathBillTo") =Mid(scriptpathBillTo,1,Instrrev(Mid(scriptpathBillTo,1,instrrev(scriptpathBillTo,"\")-1),"\"))

'Describes the path of Repository
Repositoriescollection.Add environment.value("varpathBillTo")&"\ObjectRepository\AccountCreation.tsr"
datatable.AddSheet "Sheet1"

varpathBillTo = environment.value("varpathBillTo")
varpathBillTo = Mid(scriptpathBillTo,1,Instrrev(Mid(varpathBillTo,1,instrrev(varpathBillTo,"\")-1),"\"))
environment.value("varpathBillTo") = varpathBillTo

'Descibes the path of TestData
Datatable.ImportSheet environment.value("varpathBillTo")&"TestData\AccountCreationBillTo.xlsx","Sheet1","Sheet1"
'datatable.getsheet("Sheet1").SetCurrentRow 1
RowCount = Datatable.GetSheet("Sheet1").GetRowCount

For k = 1 To RowCount
	Datatable.SetCurrentRow(k)
	RunTest = Datatable.GetSheet("Sheet1").GetParameter("Run")

	If RunTest = "Yes" Then
		
		'Incrementing the BilltoAccName 
		BilltoAccName = datatable.Value("BilltoAccName","Sheet1")
		BilltoAccNamearr = split(BilltoAccName,"_")
		BilltoAccNamearr(ubound(BilltoAccNamearr)) =cstr( cint(BilltoAccNamearr(ubound(BilltoAccNamearr)))+1)
		strBilltoAccNameVerify=BilltoAccName
		BillToAccNM = BilltoAccNamearr(0)
		For i=1  to  Ubound(BilltoAccNamearr)
			BillToAccNM = BillToAccNM&"_"&BilltoAccNamearr(i)
		Next
		
		datatable.Value("BilltoAccName","Sheet1") = BillToAccNM
		
		'Incrementing the ShiptoAccName
		ShiptoAccName = datatable.Value("ShiptoAccName","Sheet1")
		ShiptoAccNamearr = split(ShiptoAccName,"_")
		ShiptoAccNamearr(ubound(ShiptoAccNamearr)) =cstr( cint(ShiptoAccNamearr(ubound(ShiptoAccNamearr)))+1)
		ShipToAccNM = ShiptoAccNamearr(0)
		
		For m =1  to  Ubound(ShiptoAccNamearr)
			ShipToAccNM = ShipToAccNM&"_"&ShiptoAccNamearr(m)
		Next
		
		datatable.Value("ShiptoAccName","Sheet1") = ShipToAccNM


		'Intializing the Values
		BusinessUnit  = datatable.Value("BusinessUnit","Sheet1")
		Channel		= datatable.Value("Channel","Sheet1")
		SubChannel  = datatable.Value("SubChannel","Sheet1")
		Phone		= datatable.Value("Phone","Sheet1")
		Fax			= datatable.Value("FAX","Sheet1")   
		CommercialRelation = datatable.Value("CommercialRelation","Sheet1")
		Website				= datatable.Value("Website","Sheet1")
		AccountCurrency = datatable.Value("AccountCurrency","Sheet1")
		SalesAgreementCodes = datatable.Value("SalesAgreementCodes","Sheet1")
		Street    = datatable.Value("Street","Sheet1")
		City = datatable.Value("City","Sheet1")
		Zip = datatable.Value("Zip","Sheet1")
		State = datatable.Value("State","Sheet1")
		Country = datatable.Value("Country","Sheet1")
		CreditRating = datatable.Value("CreditRating","Sheet1")
		MDMID = datatable.Value("MDMID","Sheet1")
		MDMID = Cstr(cdbl(MDMID)+1)
		MDMIDShipto=Cstr(cdbl(MDMID)+1)
		datatable.Value("MDMID","Sheet1") = MDMIDShipto
		'datatable.ExportSheet environment.value("varpath")&"TestData\AccountCreation.xlsx","Sheet1","Sheet1"
		datatable.ExportSheet environment.value("varpathBillTo")&"TestData\AccountCreationBillTo.xlsx","Sheet1","Sheet1"
		SrcCountry = datatable.Value("SrcCountry","Sheet1")




			'Account
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").Sync
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").Link("lnk_MainWinAccounts").Click
			wait 2
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").Link("lnk_AccountNew").Click
			
			
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("Bill To").Click
			
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_BilltoBillto").Click
 @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebRadioGroup("changeRecordTypeRadio1:8392;a")_;_script infofile_;_ZIP::ssf4.xml_;_
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Next").Click @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Next")_;_script infofile_;_ZIP::ssf5.xml_;_
			wait 1
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_AccNameBillto").Set BilltoAccName
			wait 3
			
			'BusinessUnit
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebList("lst_BusinessUnit").Select BusinessUnit '"PLNA"
			
			'Channel
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebButton("btn_Channel").Click
			wait 3
			'set the user difine list like below -
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_ChannelOptionlist").Link("lnk_ChannerlOptions").SetTOProperty "title",Channel
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_ChannelOptionlist").Link("lnk_ChannerlOptions").Click
			
			wait 1
			
			'Phone
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_Phone").Set Phone'"111-111-2221"
			'Fax
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_Fax").Set Fax '"111-111-2222"
			'SubChannel @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebEdit("WebEdit 2")_;_script infofile_;_ZIP::ssf12.xml_;_
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").Link("lnk_SubChannel").Click
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebButton("btn_Sub Channel").Click
			
			'set the user difine list like below -subchannel
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_SubChannelOptionlist").Link("lnk_listoptions").SetTOProperty "title",SubChannel
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_SubChannelOptionlist").Link("lnk_listoptions").Click
			
			'CommercialRelation
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").Link("lnk_CommercialRelation").Click
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebButton("btn_Commercial Relation").Click
			
			wait 3
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_CommercialRelationlistblank").Link("lnk_listoptions").SetTOProperty "title",CommercialRelation @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").Link("ADVANTAGE")_;_script infofile_;_ZIP::ssf11.xml_;_
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_CommercialRelationlistblank").Link("lnk_listoptions").Click
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_WebSite").Set Website 
			wait 1
			
			'Accountcurrency
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").Link("lnk_AccountCurrency").Click
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebButton("btn_Account Currency").Click
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_AccountCurrencylist").Link("lnk_listoptions").SetTOProperty "title", AccountCurrency
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_AccountCurrencylist").Link("lnk_listoptions").Click
			
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_SalesAgreementCode").Set SalesAgreementCodes'"R123489"
			
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_SalesAgreementCode1").Set SalesAgreementCodes'"R123489"
 @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").Link("Direct")_;_script infofile_;_ZIP::ssf15.xml_;_
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_BillingStreet").Object.scrollIntoView
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_BillingStreet").Set Street
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_BillingCity").Set City
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_BillingState").Set State
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_Billingzip").Set Zip
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_BillingCountry").Set Country
			
			'shipping address
			
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_ShippingStreet").Set Street'"1 Parkway South" @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebEdit("WebEdit 11")_;_script infofile_;_ZIP::ssf23.xml_;_
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_ShippingCity").Set City'"Greenville" @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebEdit("WebEdit 12")_;_script infofile_;_ZIP::ssf24.xml_;_
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_Shippingstate").Set State'"SC" @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebEdit("WebEdit 13")_;_script infofile_;_ZIP::ssf25.xml_;_
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_Shippingzip").Set Zip' "29615" @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebEdit("WebEdit 14")_;_script infofile_;_ZIP::ssf26.xml_;_
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_ShippingCoutry").Set Country'"US" @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebEdit("WebEdit 15")_;_script infofile_;_ZIP::ssf27.xml_;_
			
			wait 2
			
			'credit rating
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").Link("lnk_CreditRating").hightlight
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").Link("lnk_CreditRating").Object.scrollIntoView
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebButton("btn_Credit Rating").Object.scrollIntoView
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebButton("btn_Credit Rating").Click
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").Link("lnk_CreditRating").Click
			
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_CreditRatingOptionlist").Link("lnk_listoptions").SetTOProperty "title",CreditRating
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_CreditRatingOptionlist").Link("lnk_listoptions").Click
			
			'MMID
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_CreditRatingOptionlist").WebEdit("txt_MDMID").Set MDMID' "9999995" @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebEdit("WebEdit 16")_;_script infofile_;_ZIP::ssf30.xml_;_
			
			'source country
			'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").Link("lnk_SourceCountry").Click
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebButton("btn_Source Country").Click
			
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_SourceCountryOptionlist").Link("lnk_listoptions").SetTOProperty "title", SrcCountry
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_SourceCountryOptionlist").Link("lnk_listoptions").Click
			Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Save").Click @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Save")_;_script infofile_;_ZIP::ssf33.xml_;_
			
			wait 5 ' 

				'Verifications
		
		
		'Account Name verification
		strAccName=Browser("Accounts | Salesforce").Page("Verifications").WebElement("AccountName").GetROProperty("innertext")
		If Trim(strAccName)=Trim(strBilltoAccNameVerify) Then
			AddNewCase 2,"Account Name Verification","Account name should be as expected","Account name:"&strBilltoAccNameVerify&" is as expected","Pass"
			'Reporter.ReportEvent micPass,"Account Name Verification", "Account: "&strBilltoAccNameVerify&" Verification Successful"	
		Else
			AddNewCase 2,"Account Name Verification","Account name should be as expected","Account name:"&strBilltoAccNameVerify&" is not as expected","Fail"
			'Reporter.ReportEvent micFail, "Account Name Verification", "Account: "&strBilltoAccNameVerify&" Verification Failure"	
		End If
		
		'Sub Channel verification
		strAccName=Browser("Accounts | Salesforce").Page("Verifications").WebElement("SubChannel").GetROProperty("innertext")
		If Trim(strAccName)=Trim(SubChannel) Then
			 AddNewCase 3,"SubChannel Verification","SubChannel should be as expected","SubChannel:"&SubChannel&" is as expected","Pass"
			'Reporter.ReportEvent micPass,"SubChannel Verification", "SubChannel: "&SubChannel&" Verification Successful"	
		Else
		  AddNewCase 3,"SubChannel Verification","SubChannel should be as expected","SubChannel:"&SubChannel&" is not as expected","Fail"
			'Reporter.ReportEvent micFail, "SubChannel Verification", "SubChannel: "&SubChannel&" Verification Failure"	
		End If
		
		'Channel verification
		strAccName=Browser("Accounts | Salesforce").Page("Verifications").WebElement("Channel").GetROProperty("innertext")
		If Trim(strAccName)=Trim(Channel) Then
		AddNewCase 4,"Channel Verification","Channel should be as expected","Channel:"&Channel&" is as expected","Pass"
			'Reporter.ReportEvent micPass,"Channel Verification", "Channel: "&Channel&" Verification Successful"	
		Else
		AddNewCase 4,"Channel Verification","Channel should be as expected","Channel:"&Channel&" is not as expected","Fail"
			'Reporter.ReportEvent micFail, "Channel Verification", "Channel: "&Channel&" Verification Failure"	
		End If
		
		'Streetname verification
		strAccName=Browser("Accounts | Salesforce").Page("Verifications").WebElement("Street Name").GetROProperty("innertext")
		strAccName=Replace(strAccName, ",","")
		If Trim(strAccName)=Trim(Street) Then
		AddNewCase 5,"Streetname Verification","Streetname should be as expected","Streetname:"&Street&" is as expected","Pass"
		'Reporter.ReportEvent micPass,"Streetname Verification", "Street: "&Street&" Verification Successful"	
		Else
		AddNewCase 5,"Streetname Verification","Streetname should be as expected","Streetname:"&Street&" is not as expected","Fail"
			'Reporter.ReportEvent micFail, "Streetname Verification", "Street: "&Street&" Verification Failure"	
		End If

		End If
		
Next
