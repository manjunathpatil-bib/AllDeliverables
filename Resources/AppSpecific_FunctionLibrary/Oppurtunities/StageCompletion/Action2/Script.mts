'Script Name     - StageCompletion
'Description     - StageCompletion script complets all the stages in the opportunity workflow
'Created By      -
'Created On      -
'Modified By     -
'Modified On     -
'Authour         - CGI
'-------------------------------------------------------------------------------------------------------------
'Entry point - Any page. The script clicks on Salesforce Opportunities link and picks the mentioned opportunity from datatable
'Exit point - Script completes all the stages and verifies if final stage is Close or not
'-------------------------------------------------------------------------------------------------------------

'Environment Setup

Dim RowCount
Dim RunTest

environment.value("varpathLogin")=Environment("RootResourceDirectory")
Repositoriescollection.Add environment.value("varpathLogin")&"ObjectRepository\Oppurtunities\StageCompletion.tsr"

'scriptpathLogin = environment("ScriptPath1")
'
'environment.value("varpathLogin") = Mid(scriptpathLogin,1,Instrrev(Mid(scriptpathLogin,1,instrrev(scriptpathLogin,"\")-1),"\"))
'
'varpath1 = environment.value("varpathLogin")
'
'varpath1 = Mid(scriptpathLogin,1,Instrrev(Mid(varpath1,1,instrrev(varpath1,"\")-1),"\"))
'
'environment.value("varpath1") = varpath1

varpath1=Environment("RootScriptDirectory")

environment.value("varpath1") = varpath1

Datatable.AddSheet "StageCompletion"
Datatable.ImportSheet environment.value("varpath1")&"TestData\Oppurtunities\StageCompletion.xlsx","Sheet1","StageCompletion"

RowCount = Datatable.GetSheet("StageCompletion").GetRowCount


For i = 1 To RowCount
	Do
		Datatable.SetCurrentRow(i)
		RunTest = Datatable.GetSheet("StageCompletion").GetParameter("Run")
			
		'Fetch Values from the datasheet
		OpportunityName=DataTable.Value("OpportunityName","StageCompletion")
		Scenario=DataTable.Value("Scenario","StageCompletion")
		Probability=DataTable.Value("Probability","StageCompletion")
		
		If RunTest = "Yes" Then
			'Wait for page to load properly
			Browser("Home | Salesforce").Page("Home | Salesforce").Sync
			Browser("OpportunityNew | Salesforce").Page("App Launcher | Salesforce").WebButton("App Launcher").Click
			Wait 3
			Browser("App Launcher | Salesforce").Page("App Launcher | Salesforce").Link("Opportunities").Click
			'Click on Oppurtunities link
			If Browser("Home | Salesforce").Page("Home | Salesforce").Link("Opportunities").Exist(conExistTimeout) Then
				Browser("Home | Salesforce").Page("Home | Salesforce").Link("Opportunities").Click
			End If
			Wait 5 
'			'Click on the oppurtunityname mentioned in data table
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").Link("OppurtunityName").SetTOProperty "text",OpportunityName
			Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").Link("OppurtunityName").Click
			Wait 5 
			CurrentStage=Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("CurrentStage").GetROProperty("innertext")
			Select Case CurrentStage
				Case "Open"
					strTimes=3
				Case "Proposal"
					strTimes=2
				Case "Negotiation"
					strTimes=1
				Case "Closed"
					strTimes=0
			End Select
			For Iterator = 1 To strTimes Step 1
				'Add a new task
				AddNewTask()
				'Change probability value
				Probability=Probability+"%"
				Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebButton("OpportunityEditButton").Click
				Wait 3
				Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebEdit("ProbabilityEdit").Click
				Wait 2
				Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebEdit("ProbabilityEdit").Set Probability
				Wait 3
				Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("EditSave").Click
				Wait 3 
				Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("Mark Stage as Complete").Click
				If Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("StageClosingOverlayText").Exist(5) Then
					Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebList("select").Select 1
					Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebButton("Save").Click
				End If
				Probability=Replace(Probability,"%","")
				Probability=cint(Probability)+10
			Next
		End If
	Loop While False
	CurrentStage=Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("CurrentStage").GetROProperty("innertext")
	If CurrentStage="Closed" Then
		AddNewCase strTCID,""&Scenario,"All stages should be completed","All stages are completed","Pass"
	Else
		AddNewCase strTCID,""&Scenario,"All stages should be completed","Error in stage completion","Fail"	
	End If
Next

Function AddNewTask()
	'Generate Task subject name
	Subject="TestSubjectName"&Int ((999 - 100 + 1) * Rnd + 100)
	'generate due date
	DueDate=Date
	'Click on Log a call first
	Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("Log a Call").Click
	Wait 5
	'Click on the New Task button
	Browser("Opportunities | Salesforce").Page("Opportunities | Salesforce").WebElement("New Task").Click
	Wait 5 
	'Enter Task Subject
	Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebEdit("TaskSubject").Set Subject
	'Enter Task Due Date
	Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebEdit("TaskDuedate").Set DueDate
	'Enter TaskStatus
	Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebButton("TaskStatus").Click
	Wait 2
	Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebMenu("StatusMenu").Select "Completed"
	'Click on SAve button
	Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebButton("TaskSave").Click
End Function


