'****************************************************************************************************************
' *Script Name:	institution Relations
'  ActionName: deleteChainInstitutions
'         This action will
'             - delete child institutions from reverse order (ie, from tree's leafs)
'
'Date		Name					Description
'* ------		--------					---------------
'****************************************************************************************************************
Option Explicit 

Dim i, env, institution, trackerlogin, trackerpass, instPrefix, vertical, rc, row, result

env = DataTable.Value("env", dtGlobalSheet)
trackerLogin = "qtptester"
trackerpass = "4d7571d24bba0a9a9e310777afdca6e031bb9a4b449d9addda4a4f32"
instPrefix = "CTY_YAN_2_"
vertical = DataTable.Value("vertical", dtGlobalSheet)

If tracker_login(env, trackerLogin, trackerPass) <> "successful" Then
	Reporter.ReportEvent micFail, "login", "Tracker Login failed"
	ExitAction
End If

rc = DataTable.GetSheet("driver").GetRowCount
row = rc
For i = 1 To rc
    DataTable.GetSheet("driver").SetCurrentRow(row)
    institution = instPrefix & DataTable.GetSheet("driver").GetParameter("childInstitution")  
	result = delete_institution_vertical(institution, vertical)
	Reporter.ReportEvent micDone, "deletion", result
	row = row -1  
Next



