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
Dim intStageCompleted
intStageCompleted=0

environment.value("varpathLogin")=Environment("RootResourceDirectory")
Repositoriescollection.Add environment.value("varpathLogin")&"ObjectRepository\Oppurtunities\StageCompletion.tsr"
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
		TaskCreation=DataTable.Value("TaskCreation","StageCompletion")
		
		If RunTest = "Yes" Then
			'Wait for page to load properly
			Browser("Home | Salesforce").Page("Home | Salesforce").Sync
			Wait 3
			ClickonOpportunitiesLink
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
				If TaskCreation="Yes" Then
					'Add a new task
					AddNewTask()
				End If
				'Change probability value
				Probability=cstr(Probability)
				Probability=Probability+"%"
				If Iterator=2 Then
					If Browser("OpportunityNew | Salesforce").Page("App Launcher | Salesforce").WebElement("Change Opportunity Owner").Exist(5) Then
						Browser("OpportunityNew | Salesforce").Page("App Launcher | Salesforce").WebButton("Cancel").Click
					End If
				End If
				Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebButton("OpportunityEditButton").Click
				Wait 3
				If Iterator=2 Then
					If Browser("OpportunityNew | Salesforce").Page("App Launcher | Salesforce").WebElement("Change Opportunity Owner").Exist(5) Then
						Browser("OpportunityNew | Salesforce").Page("App Launcher | Salesforce").WebButton("Cancel").Click
					End If
				End If
				Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebEdit("ProbabilityEdit").Click
				Wait 2
				Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebEdit("ProbabilityEdit").Set Probability
				Wait 3
				Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("EditSave").Click
				Wait 3 
				CurrentStage=Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("CurrentStage").GetROProperty("innertext")
				Wait 3
				If Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("Mark Stage as Complete").Exist(5) Then
					Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("Mark Stage as Complete").Click				
				End If
				If Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("Stage change popup").Exist(5) Then
					AddNewCase strTCID,"Stage completion : "&CurrentStage,""&CurrentStage&" stage should be successfully completed",""&CurrentStage&" stage is successfully completed","Pass"
					intStageCompleted=1
				End If
				If Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("StageClosingOverlayText").Exist(5) Then
					Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebList("select").Select 1
					Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebButton("Save").Click
				End If
					If Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("Stage change popup").Exist(5) Then
					AddNewCase strTCID,"Stage completion : "&CurrentStage,"Opportunity : "&OpportunityName&" - Stage "&CurrentStage&" should be successfully completed","Opportunity : "&OpportunityName&" - Stage "&CurrentStage&" stage is successfully completed","Pass"
					intStageCompleted=1
				End If
				If intStageCompleted=0 Then
					AddNewCase strTCID,"Stage completion : "&CurrentStage,"Opportunity : "&OpportunityName&" - Stage "&CurrentStage&" should be successfully completed","Opportunity : "&OpportunityName&" - Stage "&CurrentStage&" stage is not successfully completed","Fail"
				End If
				Probability=Replace(Probability,"%","")
				Probability=cint(Probability)+10
				If Err.Number<>0 Then
					Exit Do
				End If
			Next
		Wait 10
		CurrentStage=Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("CurrentStage").GetROProperty("innertext")
		If CurrentStage="Closed" Then
			AddNewCase strTCID,""&Scenario,"Opportunity : "&OpportunityName&" - All stages should be completed","All stages are completed","Pass"
		Else
			AddNewCase strTCID,""&Scenario,"Opportunity : "&OpportunityName&" - All stages should be completed","Error in stage completion. Failed stage : "&CurrentStage,"Fail"	
		End If
	End If
	Loop While False
Next

Function AddNewTask()
	'Check current stage
	CurrentStage=Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("CurrentStage").GetROProperty("innertext")
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
	Wait 3
	Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").Link("Completed").Click
	'Click on SAve button
	Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebButton("TaskSave").Click

	If Browser("OpportunityNew | Salesforce").Page("OpportunityNew | Salesforce").WebElement("Task Creation Confirmation").Exist(conExistTimeout) Then
		AddNewCase strTCID,"Opportunity : "&OpportunityName&" -Task Creation in stage "&CurrentStage,"Task should be successfully created in the "&CurrentStage&" stage","Task is successfully created in the "&CurrentStage&" stage","Pass"
	Else	
		AddNewCase strTCID,"Opportunity : "&OpportunityName&" -Task Creation in stage "&CurrentStage,"Task should be successfully created in the "&CurrentStage&" stage","Task creation in the "&CurrentStage&" stage has failed","Fail"
	End If
End Function


