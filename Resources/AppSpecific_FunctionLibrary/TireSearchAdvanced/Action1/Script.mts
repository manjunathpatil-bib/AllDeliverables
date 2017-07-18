'Script Name     - TireSearchAdvanced
'Description     - Script performs advanced search on Tire search portal
'Created By      -
'Created On      -
'Modified By     -
'Modified On     -
'Authour         - CGI
'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------

'Environment Setup

' Variable Declaration
Dim RowCount
Dim ProductNo
Dim Qty
Dim Brand
Dim AspectRatio
Dim RimSize
Dim Season
Dim SpeedRating
Dim ActBrand
Dim ActProductNo
Dim ActSeasonText
Dim Flag

scriptpathLogin = environment("ScriptPath1")
environment.value("varpathLogin") = Mid(scriptpathLogin,1,Instrrev(Mid(scriptpathLogin,1,instrrev(scriptpathLogin,"\")-1),"\"))
Repositoriescollection.Add environment.value("varpathLogin")&"ObjectRepository\TireSearchAdvanced.tsr"
varpath1 = environment.value("varpathLogin")
varpath1 = Mid(scriptpathLogin,1,Instrrev(Mid(varpath1,1,instrrev(varpath1,"\")-1),"\"))
environment.value("varpath1") = varpath1
Datatable.AddSheet "Sheet1"
Datatable.ImportSheet environment.value("varpath1")&"TestData\TireSearchAdvanced.xlsx","Sheet1","Sheet1"
RowCount = Datatable.GetSheet("Sheet1").GetRowCount

For i = 1 To RowCount
	Datatable.SetCurrentRow(i)
	RunTest = Datatable.GetSheet("Sheet1").GetParameter("Run")
	
	If RunTest = "Yes" Then
			
		VehicleType = Datatable.Value("Type","Sheet1")	
		BillTo = Datatable.Value("BillToNumber","Sheet1")
		ShipTo = Datatable.Value("ShipToNumber","Sheet1")
		ProductNo = Datatable.Value("SearchProductNo","Sheet1")
		ProductNoEx = Mid(ProductNo, 1, 3)	
		Qty = Datatable.Value("SearchQty","Sheet1")
		' Passenger Car values
		Brand = Datatable.Value("Brand","Sheet1")
		AspectRatio = Datatable.Value("AspectRatio","Sheet1")
		RimSize = Datatable.Value("RimSize","Sheet1")
		Season = Datatable.Value("Season","Sheet1") 
		SpeedRating = Datatable.Value("SpeedRating","Sheet1")
		Browser("Products").Page("Products").Link("Tires").Click
		wait 2				
		' Select BillTo value
		If Browser("Products").Page("Products").WebElement("SelectBillTo").Exist Then
			Browser("Products").Page("Products").WebElement("SelectBillTo").Click
			wait 2
			Browser("Products").Page("Products").WebEdit("BillToWebEdit").Set BillTo
			wait 2
			Browser("Products").Page("Products").WebElement("BillToSearchResult").Click
			wait 2
		End If
		' Select ShipTo value
		If Browser("Products").Page("Products").WebElement("SelectShipTo").Exist Then
			Browser("Products").Page("Products").WebElement("SelectShipTo").Click
			wait 2
			Browser("Products").Page("Products").WebEdit("ShipToWebEdit").Set ShipTo
			wait 2
			Browser("Products").Page("Products").WebElement("ShipToSearchResult").Click
            wait 2	
		End If
		' Select Vehicle Type
		If VehicleType = "PC" Then
			Browser("Products").Page("Products").Link("Passenger car & light").FireEvent "onmouseover"
			wait 1
			Browser("Products").Page("Products").Link("Passenger car & light").Click
			wait 1
		ElseIf VehicleType = "CT" Then
			Browser("Products").Page("Products").Link("Commercial truck").FireEvent "onmouseover"
			wait 1
			Browser("Products").Page("Products").Link("Commercial truck").Click
			wait 1
		End If
		Browser("Products").Page("Products").WebEdit("Search Products").Set ProductNo
		Browser("Products").Page("Products").WebNumber("Quantity").Set Qty
		Browser("Products").Page("Products").WebElement("Advance Search").Click
		wait 2
		Browser("Products").Page("Products").WebList("selectBrand").Select Brand
		wait 1
		Browser("Products").Page("Products").WebEdit("AspectRatio").Set AspectRatio
		wait 1
		Browser("Products").Page("Products").WebEdit("RimSize").Set RimSize
		wait 1
		Browser("Products").Page("Products").WebList("selectSeason").Select Season
		wait 1
		Browser("Products").Page("Products").WebList("selectSpeedRating").Select SpeedRating
		wait 1
		' Uncheck the Show Zero Quantity check box
		Browser("Products").Page("Products").WebCheckBox("WebCheckBoxZero").Set "OFF"
		Browser("Products").Page("Products").Link("SEARCH").Click
		Browser("Products").Page("Products").Sync
		wait 10
		' Check for the Error message
		If  Browser("Products").Page("Products").WebElement("ErrorMessage").Exist Then
			ErrorMessageApp = Browser("Products").Page("Products").WebElement("ErrorMessage").GetROProperty("innertext")
			AddNewCase strTCID, "Advance Search","Search results should be shown.", "Search results are not as expected. Actual value : "&ErrorMessageApp, "Fail"
			Exit For
		End If
		If Browser("Products").Page("Products").WebElement("Brand").Exist Then
			ActBrand = Browser("Products").Page("Products").WebElement("Brand").GetROProperty("innertext")
		End If
		If Browser("Products").Page("Products").WebElement("ProductNo").Exist Then
			ActProductNo = Browser("Products").Page("Products").WebElement("ProductNo").GetROProperty("innertext")	
		End If
		If Browser("Products").Page("Products").WebElement("SeasonText").Exist Then
			ActSeasonText = Browser("Products").Page("Products").WebElement("SeasonText").GetROProperty("innertext")
		End If		
		' Verify Brand of Tire
		If Brand<>"" Then
			If  Brand = ActBrand Then
				AddNewCase strTCID, "Advance Search - Brand Verification","Brand Name should be shown. Expected value : "&Brand, "Brand Name is as expected. Actual value : "&ActBrand, "Pass"
			Else
				AddNewCase strTCID, "Advance Search - Brand Verification","Brand Name should be shown. Expected value : "&Brand, "Brand Name is Not as expected. Actual value : "&ActBrand, "Fail"
			End If
		End If
		' Verify Season of Tire
		If  Season<>"" Then
		Season = Replace(Season,"-"," ")
			If Season = ActSeasonText Then
				AddNewCase strTCID, "Advance Search - Season Verification", "Season should be shown. Expected value : "&Season, "Season is as expected. Actual value : "&ActSeasonText, "Pass"
			Else
				AddNewCase strTCID, "Advance Search - Season Verification", "Season should be shown. Expected value : "&Season, "Season is Not as expected. Actual value : "&ActSeasonText, "Fail"
			End If
		End If
		' Verify Width of Tire
		strWidth = Mid(ActProductNo, 1, 3)
		If ProductNoEx = strWidth Then
			AddNewCase strTCID, "Advance Search - Width Verification", "Width of the tire should be shown. Expected value : "&ProductNoEx, "Width of the tire is as expected. Actual value : "&strWidth, "Pass"
		Else
			AddNewCase strTCID, "Advance Search - Width Verification", "Width of the tire should be shown. Expected value : "&ProductNoEx, "Width of the tire is Not as expected. Actual value : "&strWidth, "Fail"		
		End If
		' Verify Speed Rating of Tire
		strSpeedRating=Right(ActProductNo,1)
		If SpeedRating<>"" Then
			If SpeedRating = strSpeedRating  Then
				AddNewCase strTCID, "Advance Search - Speed Rating Verification", "Speed Rating of the tire should match. Expected value : "&SpeedRating, "Speed Rating of the tire is as expected. Actual value : "&strSpeedRating, "Pass"
			Else
				AddNewCase strTCID, "Advance Search - Speed Rating Verification", "Speed Rating of the tire should match. Expected value : "&SpeedRating, "Speed Rating of the tire is not as expected. Actual value : "&strSpeedRating, "Fail"
			End If
		End If
		' Verify Aspect Ratio of Tire
		strAspectRatio=Split(Split(ActProductNo,"/")(1),"R")(0)
		If AspectRatio<>"" Then
			If AspectRatio = strAspectRatio Then
				AddNewCase strTCID, "Advance Search - Aspect Ratio Verification", "Aspect Ratio of the tire should match. Expected value : "&AspectRatio, "Aspect Ratio of the tire is as expected. Actual value : "&strAspectRatio, "Pass"
			Else	
				AddNewCase strTCID, "Advance Search - Aspect Ratio Verification", "Aspect Ratio of the tire should match. Expected value : "&AspectRatio, "Aspect Ratio of the tire is not as expected. Actual value : "&strAspectRatio, "Fail"
			End If	
		End If
		' Verify Rim Size of Tire
		strRimSize=Split(Split(Split(ActProductNo,"/")(1),"R")(1)," ")(0)
		If RimSize<>"" Then
			If  RimSize = strRimSize Then
				AddNewCase strTCID, "Advance Search - Rim Size Verification", "Rim Size of the tire should match. Expected value : "&RimSize, "Rim Size of the tire is as expected. Actual value : "&strRimSize, "Pass"
			Else
				AddNewCase strTCID, "Advance Search - Rim Size Verification", "Rim Size of the tire should match. Expected value : "&RimSize, "Rim Size of the tire is not as expected. Actual value : "&strRimSize, "Fail"
			End If
		End If
		' Verify Weight of Tire
		If Browser("Products").Page("Products").WebElement("TyreWeight").Exist Then
			TireWeight  = Browser("Products").Page("Products").WebElement("TyreWeight").GetROProperty("innertext")	
			AddNewCase strTCID, "Advance Search - Weight Verification", "Tire weight should be shown.","Tire weight is as Expected. Actual value : "&TireWeight, "Pass"
		Else
		   	AddNewCase strTCID, "Advance Search - Weight Verification", "Tire weight should be shown.","Tire weight is Not as Expected. Actual value : "&TireWeight, "Fail"
		End If
		' Verify Quantity of Tire
		If Browser("Products").Page("Products").WebElement("QtyAvailable").Exist Then
			TireQtyAct = Browser("Products").Page("Products").WebElement("QtyAvailable").GetROProperty("innertext")
			If TireQtyAct >= 0 Then
				AddNewCase strTCID, "Advance Search - Quantity Verification", "Tire quantity should be shown. Expected value : "&Qty, "Tire quanity is as Expected. Actual value : "&TireQtyAct, "Pass"
			Else	
				AddNewCase strTCID, "Advance Search - Quantity Verification", "Tire quantity should be shown. Expected value : "&Qty, "Tire quanity is Not shown. Actual value : "&TireQtyAct, "Fail"
			End If
			
		End If	
		' Verify the Brand Name for all search records
		If Brand<>"" Then
		Set BrandObjDesc=Description.Create
		BrandObjDesc("micclass").value="WebElement"
		BrandObjDesc("class").value="mdp-tires__result-brand"
		BrandObjDesc("html tag").value="DIV"
		BrandObjDesc("visible").value="True"
		Set BrandObj= Browser("Products").Page("Products").ChildObjects(BrandObjDesc)
			For Iterator = 0 To BrandObj.count-1 Step 1
				ActualBrand=BrandObj(Iterator).GetROProperty("innertext")
				If Trim(ActualBrand)=Trim(Brand) Then
					Flag = 1
				Else
					Flag = 0
					Exit For
				End If
			Next
			If Flag = 1 Then
			 	   AddNewCase strTCID, "Advance Search - Brand Verification for All Records", "Brand Name should match for all Records. Expected value : "&Brand, "Brand Name matches for all Records. Actual value : "&ActualBrand, "Pass"
				Else
					AddNewCase strTCID, "Advance Search - Brand Verification for All Records", "Brand Name should match for all Records. Expected value : "&Brand, "Brand Name doesn't matches for all Records. Actual value : "&ActualBrand, "Fail"
				End If
			End If
    	End If
Next










