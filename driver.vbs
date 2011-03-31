'****************************************************************************************************************
' *Script Name:	institution Relations
'  ActionName: addChildInstitution
'
'  Date		  Name					  Description
'* ------		--------					---------------
'****************************************************************************************************************** 
Option Explicit
Dim i, env, parentInstitution, childInstitution, trackerlogin, trackerpass, instPrefix
Dim tmp_pwd, userpwd, instPerson, instPersonPrefix, vertical, rc

env = DataTable.Value("env", dtGlobalSheet)
trackerLogin = "qtptester"
trackerpass = "4d7571d24bba0a9a9e310777afdca6e031bb9a4b449d9addda4a4f32"
instPrefix = "CTY_YAN_2_"
instPersonPrefix = "ctyyan2_"
userpwd = "Qtppass1"
vertical = DataTable.Value("vertical", dtGlobalSheet)
rc = DataTable.GetSheet("driver").GetRowCount

If tracker_login(env, trackerLogin, trackerPass) <> "successful" Then
	Reporter.ReportEvent micFail, "login", "Tracker Login failed"
	ExitAction
End If

'first create the first parent institution
DataTable.GetSheet("driver").SetCurrentRow(1)
parentInstitution = instPrefix & DataTable.GetSheet("driver").GetParameter("parentInstitution")
If search_institution_by_name(parentInstitution, vertical) <> "successful" Then
	Browser("Tracker").Page("Tracker").Frame("Add Institution").Link("Add Institution").Click
	With	Browser("Tracker Popups").Page("Institution - Information")
		.Sync
		.WebList("Version").Select "NG"
	End With
	create_institution parentInstitution
End If

With	Browser("Tracker Popups")
	For i = 1 To rc
		DataTable.GetSheet("driver").SetCurrentRow(i)
		parentInstitution = instPrefix & DataTable.GetSheet("driver").GetParameter("parentInstitution")
		childInstitution = instPrefix & DataTable.GetSheet("driver").GetParameter("childInstitution")
		If search_institution_by_name(parentInstitution, vertical) <> "successful" Then
			Reporter.ReportEvent micFail, "search", "search parent institution failed"
			ExitAction
		End If
		'add child institution.
		With	.Page("Institution - Information")
			.Sync
			.Link("Child Agencies").Click
		End With
		With	.Page("Add Child Institution")
			.Sync
			.Link("Add child Institution").Click
		End With
		create_institution childInstitution
		Print parentInstitution & " -> " & childInstitution
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
			Else
				temp_login env, instPerson, tmp_pwd
				setup_security_profile "", userpwd
		End If
		tracker_login env, trackerLogin, trackerPass
	Next
End With

Sub create_institution(inst_name)
	With	Browser("Tracker Popups")
		With	.Page("Institution - Information")
			.Sync
			.WebEdit("InstitutionName").Set inst_name
			.WebEdit("M_phone").Set "333-333-3333"
			.WebEdit("Address").Set "454 ROYAL LANE"
			.WebEdit("City").Set "TAUNGGYI"
			.WebList("State").Select "New York"
			.WebEdit("Zip").Set "11366"
			.WebEdit("SIS_cd").Set Right(inst_name, 2)
			.WebList("TimeZone").Select "(GMT-05:00) Eastern Time (US & Canada)"
			.WebButton("Submit").Click
		End With
		.Dialog("Windows Internet Explorer").WinButton("OK").Click
		.Page("Institution - Information").WebButton("Close").Click
	End With
End Sub

Function add_inst_person(instperson)
    'pre-condition: needs to be on institution information page.
	'This function will create an institution person username instperson and 
	' reset the password and return the temp password.
	Dim retval, temppass
	On Error Resume Next
	retval = "unsuccessful"
	With	Browser("Tracker Popups")
		With	.Page("Institution - Information")
			.Sync
			.Link("People").Click
		End With
		With	.Page("Institution - People")
			.Sync
			.Link("Add Person").Click
		End With
		With	.Page("Institution - Person")
			.Sync
			.WebEdit("FirstName").Set "Rikki"
			.WebEdit("LastName").Set "LN" & instperson
			.WebList("Title").Select "#1"
			.WebEdit("EmailAddress").Set "qatest@blackboardconnect.com"
			.WebEdit("PrimaryPhone").Set "333-333-3333"
			.WebEdit("UserName").Set instperson
			.WebButton("Submit").Click
		End With
		.Dialog("Windows Internet Explorer").WinButton("OK").Click
		With	.Page("Institution - Person")
			.Sync
			.Link("Click here").Click
		End With
		.Dialog("Windows Internet Explorer").WinButton("OK").Click
	End With
	temppass = get_temp_pwd()
	If temppass <> "" Then
		retval = temppass
	End If
	add_inst_person = retval
End Function




