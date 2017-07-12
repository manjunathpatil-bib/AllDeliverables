'-----------------------------------------------------------------------------------------------------------------
'Script Name  - PortalLogout
'Description  - Logout of the Portal
'Created By   -
'Created On   -
'Modified By  -
'Modified On  -
'Authour      - CGI
'-----------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------
'Environment setup

' Variable Declartion

scriptpathLogin = environment("ScriptPath1")

environment.value("varpathLogin") = Mid(scriptpathLogin,1,Instrrev(Mid(scriptpathLogin,1,instrrev(scriptpathLogin,"\")-1),"\"))

Repositoriescollection.Add environment.value("varpathLogin")&"ObjectRepository\PortalLogout.tsr"

' Logout of Portal
Browser("Products").Page("Products").WebElement("UserLogged").FireEvent "onmouseover"
wait 2
'Browser("Products").Page("Products").WebElement("UserLogged").Click
'wait 2
Browser("Products").Page("Favorites").Link("Log Out").Click
wait 5


' Verification of Logout
CheckURL = Browser("Products").Page("Login").GetROProperty("URL")

If CheckURL = "https://uftpoc-partner-portal.cs83.force.com/s/dealer-login/?language=en_US" Then
	AddNewCase strTCID,"Portal Logout","User should be Logged out of Portal", "User is Logged out of Portal", "Pass"   
Else	
	AddNewCase strTCID,"Portal Logout","User should be Logged out of Portal", "User is not Logged out of Portal", "Fail"   
End If



