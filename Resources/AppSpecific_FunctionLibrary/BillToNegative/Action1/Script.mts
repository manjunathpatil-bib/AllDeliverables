'Script Name     - BillToNegative
'Description     - BillToNegative
'Created By      -
'Created On      -
'Modified By     -
'Modified On     -
'Authour         - CGI
'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------

'Environment Setup

scriptpathBillToNG = environment("ScriptPath1")

'scriptpathBillToNG = environment.value("TestDir")
environment.value("varpathBillToNG") =Mid(scriptpathBillToNG,1,Instrrev(Mid(scriptpathBillToNG,1,instrrev(scriptpathBillToNG,"\")-1),"\"))
'msgbox scriptpathBillToNG 

'Describes the path of Repository
Repositoriescollection.Add environment.value("varpathBillToNG")&"\ObjectRepository\AccountCreation.tsr"
Repositoriescollection.Add environment.value("varpathBillToNG")&"\ObjectRepository\BillToNegative.tsr"

datatable.AddSheet "Sheet1"

varpathBillToNG = environment.value("varpathBillToNG")
varpathBillToNG = Mid(scriptpathBillToNG,1,Instrrev(Mid(varpathBillToNG,1,instrrev(varpathBillToNG,"\")-1),"\"))
environment.value("varpathBillToNG") = varpathBillToNG

'Descibes the path of TestData
Datatable.ImportSheet environment.value("varpathBillToNG")&"TestData\BillToNegative.xlsx","Sheet1","Sheet1"

RowCount = Datatable.GetSheet("Sheet1").GetRowCount
'
For k = 1 To RowCount
Datatable.SetCurrentRow(k)
RunTest = Datatable.GetSheet("Sheet1").GetParameter("Run")

	If RunTest = "Yes" Then
	   'Incrementing the BilltoAccName 
        BilltoAccName = datatable.Value("BilltoAccName","Sheet1")
    
	    If BilltoAccName <> Null Then

             BilltoAccNamearr = split(BilltoAccName,"_")
             BilltoAccNamearr(ubound(BilltoAccNamearr)) =cstr( cint(BilltoAccNamearr(ubound(BilltoAccNamearr)))+1)
             BillToAccNM = BilltoAccNamearr(0)
	             
	             For i=1  to  Ubound(BilltoAccNamearr)
	             BillToAccNM = BillToAccNM&"_"&BilltoAccNamearr(i)
	             Next

             datatable.Value("BilltoAccName","Sheet1") = BillToAccNM

        End If


            'Intializing the Values
             BusinessUnit  = datatable.Value("BusinessUnit","Sheet1")
             AccountCurrency = datatable.Value("AccountCurrency","Sheet1")
			 datatable.ExportSheet environment.value("varpathBillToNG")&"TestData\BillToNegative.xlsx","Sheet1","Sheet1"
             SrcCountry = datatable.Value("SrcCountry","Sheet1")

             'Account

              Browser("Accounts | Salesforce").Page("Accounts | Salesforce").Link("lnk_MainWinAccounts").Click
              wait 2
              Browser("Accounts | Salesforce").Page("Accounts | Salesforce").Link("lnk_AccountNew").Click


              Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("Bill To").Click

			  'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_BilltoBillto").Click
				
			  Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Next").Click
			  wait 1
				
			  Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebEdit("txt_AccNameBillto").Set BilltoAccName 
			  wait 3
				
				'BusinessUnit
			    If BusinessUnit <> ""  Then
					Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebList("lst_BusinessUnit").Select BusinessUnit '"PLNA"
					
				End If
				
				
				
				''Accountcurrency
				''Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").Link("lnk_AccountCurrency").Click
				'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebButton("btn_Account Currency").Click
				'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_AccountCurrencylist").Link("lnk_listoptions").SetTOProperty "title", AccountCurrency
				'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_AccountCurrencylist").Link("lnk_listoptions").MakeObjVisible
				'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_AccountCurrencylist").Link("lnk_listoptions").Click
				'
				
				'SourceCountry
				Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("wnd_CreateAccountBillTo").WebButton("btn_Source Country").Click
				
				Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_SourceCountryOptionlist").Link("lnk_listoptions").SetTOProperty "title", SrcCountry
				Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("elm_SourceCountryOptionlist").Link("lnk_listoptions").Click
				Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Save").Click
				
			  wait 5 ' 

          'Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("These required fields").Check CheckPoint("These required fields must be completed: Account Name") @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("These required fields")_;_script infofile_;_ZIP::ssf1.xml_;_
          ExpMessage = Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebElement("These required fields").GetROProperty("innertext")
          wait 2
          
           	ConditonText = datatable.Value("ConditionText","Sheet1")
          	If ConditonText = ExpMessage  Then
          	Reporter.ReportEvent micPass, "Mandatory Field Validation", "Mandatory Field validation is success"
           	Else   
            Reporter.ReportEvent micFail, "Mandatory Field Validation", "Mandatory Field validation failed"
           	End If
           
           Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Cancel").Click @@ hightlight id_;_Browser("Accounts | Salesforce").Page("Accounts | Salesforce").WebButton("Cancel")_;_script infofile_;_ZIP::ssf2.xml_;_
           wait 2
           
    End if

Next
