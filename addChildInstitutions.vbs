'****************************************************************************************************************
' *Script Name:	institution Relations
'  ActionName: addChildInstitution
'
'  Date		  Name					  Description
'* ------		--------					---------------
'****************************************************************************************************************** 
Option Explicit
Dim i, env, parentInstitution, childInstitution, trackerlogin, trackerpass, instPrefix, instPerson, instPersonPrefix, vertical, rc
Dim tmp_pwd, userpwd



env = DataTable.Value("env", dtGlobalSheet)
trackerLogin = "qtptester"
trackerpass = "4d7571d24bba0a9a9e310777afdca6e031bb9a4b449d9addda4a4f32"
'instPrefix = "Mikayeel NG agency "
instPrefix = "Elena NG school "
instPersonPrefix = "elenang_school"


userpwd = "Qtppass1"

vertical = DataTable.Value("vertical", dtGlobalSheet)
Environment.Value("veritcal") = vertical
rc = DataTable.GetSheet("driver").GetRowCount

If tracker_login(env, trackerLogin, trackerPass) <> "successful" Then
	Reporter.ReportEvent micFail, "login", "Tracker Login failed"
	ExitAction
End If

'first create the first parent institution
DataTable.GetSheet("driver").SetCurrentRow(1)
parentInstitution = instPrefix & DataTable.GetSheet("driver").GetParameter("parentInstitution")
If search_institution_by_name(parentInstitution, vertical) <> "successful" Then
	Browser("Tracker").Page("Tracker").Frame("Header").Link("Tools").Click
	Browser("Tracker").Page("Tracker").Sync
	Browser("Tracker").Page("Tracker - Tools").Frame("Tools").Link("Add Institution").Click
	Browser("Tracker Popups").Page("Institution - Information").Sync
	Browser("Tracker Popups").Page("Institution - Information").WebList("Version").Select "Connect 5.0"
	create_institution parentInstitution
	If search_institution_by_name(parentInstitution, vertical) <> "successful" Then
		Reporter.ReportEvent micFail, "search", "search parent institution just created failed"
		ExitAction
	End If
	'and add institution person to that child institution.
	instPerson = instPersonPrefix & DataTable.GetSheet("driver").GetParameter("parentInstitution")
	tmp_pwd = add_inst_person (instPerson)
End If



For i = 1 To rc 
	
	DataTable.GetSheet("driver").SetCurrentRow(i)
	parentInstitution = instPrefix & DataTable.GetSheet("driver").GetParameter("parentInstitution")
	childInstitution = instPrefix & DataTable.GetSheet("driver").GetParameter("childInstitution")
	If search_institution_by_name(parentInstitution, vertical) <> "successful" Then
		Reporter.ReportEvent micFail, "search", "search parent institution failed"
		ExitAction
	End If

	'add child institution.
	Browser("Tracker Popups").Page("Institution - Information").Sync
	Browser("Tracker Popups").Page("Institution - Information").Link("Child Agencies").Click
	Browser("Tracker Popups").Page("Add Child Institution").Sync
	Browser("Tracker Popups").Page("Add Child Institution").Link("Add child Institution").Click
	While Browser("Tracker Popups").GetROProperty("title") <> "Institution - Information" 
		Browser("Tracker Popups").Close
		Wait 3
	Wend
	create_institution childInstitution
	Print parentInstitution & " -> " & childInstitution



'/****************************************************************************************************************************
  'following code commented out for now since UI for NG is not ready yet.
	'now search for that child institution 
	If search_institution_by_name(childInstitution, vertical) <> "successful" Then
		Reporter.ReportEvent micFail, "search", "search child institution just created failed"
		ExitAction
	End If
	'and add institution person to that child institution.
	instPerson = instPersonPrefix & DataTable.GetSheet("driver").GetParameter("childInstitution")
	tmp_pwd = add_inst_person (instPerson)
	If tmp_pwd = "unsuccessful" Then
		Reporter.ReportEvent micFail, "user creation", "institution user creation failed. username : " & instPerson
'		Else 
'		    temp_login env, instPerson, tmp_pwd
'			setup_security_profile "", userpwd
	End If

	'tracker_login env, trackerLogin, trackerPass
'*********************************************************************************************************************************
Next










Sub create_institution(inst_name)
  Dim Arr, last, siscode
  Arr = Split(inst_name, " ") 'split by space
  last = Ubound(Arr)
  siscode = Arr(last)
  Browser("Tracker Popups").Page("Institution - Information").Sync
  If Browser("Tracker Popups").Page("Institution - Information").WebList("Application").Exist(0) Then
	 Browser("Tracker Popups").Page("Institution - Information").WebList("Application").Select Environment.Value("veritcal") 
	 Browser("Tracker Popups").Page("Institution - Information").Sync
  End If
  If Browser("Tracker Popups").Page("Institution - Information").WebList("Version").GetROProperty("disabled") <> 1 Then
	  Browser("Tracker Popups").Page("Institution - Information").WebList("Version").Select "Connect 5.0"
	  Browser("Tracker Popups").Page("Institution - Information").Sync
  End If
  Browser("Tracker Popups").Page("Institution - Information").WebEdit("InstitutionName").Set inst_name
  Browser("Tracker Popups").Page("Institution - Information").WebEdit("M_phone").Set "333-333-3333"
  Browser("Tracker Popups").Page("Institution - Information").WebEdit("Address").Set "300 Taller Street"
  Browser("Tracker Popups").Page("Institution - Information").WebEdit("City").Set "Los Angeles"
  Browser("Tracker Popups").Page("Institution - Information").WebList("State").Select "California"
  Browser("Tracker Popups").Page("Institution - Information").WebEdit("Zip").Set "91403"
  Browser("Tracker Popups").Page("Institution - Information").WebEdit("SIS_cd").Set siscode
  Browser("Tracker Popups").Page("Institution - Information").WebList("TimeZone").Select "(GMT-08:00) Pacific Time (US & Canada)"
  Browser("Tracker Popups").Page("Institution - Information").WebButton("Submit").Click
  Browser("Tracker Popups").Dialog("Windows Internet Explorer").WinButton("OK").Click
  Browser("Tracker Popups").Page("Institution - Information").WebButton("Close").Click
End Sub



 




Function add_inst_person(instperson)
    'pre-condition: needs to be on institution information page.
	'This function will create an institution person username instperson and 
	' reset the password and return the temp password.
	Dim retval, temppass
	On Error Resume Next
	retval = "unsuccessful"
    Browser("Tracker Popups").Page("Institution - Information").Sync
	Browser("Tracker Popups").Page("Institution - Information").Link("People").Click
	Browser("Tracker Popups").Page("Institution - People").Sync
	Browser("Tracker Popups").Page("Institution - People").Link("Add Person").Click
    Browser("Tracker Popups").Page("Institution - Person").Sync
	Browser("Tracker Popups").Page("Institution - Person").WebEdit("FirstName").Set "Lisa"
	Browser("Tracker Popups").Page("Institution - Person").WebEdit("LastName").Set "L-" & instperson
	Browser("Tracker Popups").Page("Institution - Person").WebList("Title").Select "#1"
	Browser("Tracker Popups").Page("Institution - Person").WebEdit("EmailAddress").Set "qatest@blackboardconnect.com"
	Browser("Tracker Popups").Page("Institution - Person").WebEdit("PrimaryPhone").Set "333-333-3333"
	Browser("Tracker Popups").Page("Institution - Person").WebEdit("UserName").Set instperson
	Browser("Tracker Popups").Page("Institution - Person").WebButton("Submit").Click
	Browser("Tracker Popups").Dialog("Windows Internet Explorer").WinButton("OK").Click
	Browser("Tracker Popups").Page("Institution - Person").Sync
	Browser("Tracker Popups").Page("Institution - Person").Link("Click here").Click
	Browser("Tracker Popups").Dialog("Windows Internet Explorer").WinButton("OK").Click
	temppass = get_temp_pwd()
	Browser("Tracker Popups").Page("Institution - Person").Link("Information").Click
	Browser("Tracker Popups").Page("Institution - Information").Sync
	Browser("Tracker Popups").Page("Institution - Information").WebButton("Close").Click
	If temppass <> "" Then
		retval = temppass
	End If
	add_inst_person = retval
End Function






