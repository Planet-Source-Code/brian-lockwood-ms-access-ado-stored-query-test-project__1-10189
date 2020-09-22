<!-- METADATA TYPE="typelib" UUID="00000200-0000-0010-8000-00AA006D2EA4" NAME="ADO Type Library"-->

<%

Function Exec_qry_del_Categories(CategoryID)
 Dim strSQL 
 Dim objCmd

	Set objCmd = Server.CreateObject("ADODB.Command")

	On Error Resume Next

	strSQL = "qry_del_Categories"
	With objCmd
		.Commandtext = strSQL
		.Commandtype = adCmdStoredProc
		Set .ActiveConnection = g_objCn

		.Parameters.Append .CreateParameter("pCategoryID", adInteger, adParamInput, 4, CategoryID)
	
		.Execute Options=adExecuteNoRecords
	End With

	Set objCmd = Nothing

	If Err <> 0 Then
		Exec_qry_del_Categories = Err.Number
	Else
		Exec_qry_del_Categories = 0
	End If
End Function

Function Exec_qry_sel_Categories(CategoryID)
 Dim strSQL 
 Dim objCmd
 Dim objRs

	Set objCmd = Server.CreateObject("ADODB.Command")
	Set objRs = Server.CreateObject("ADODB.Recordset")

	On Error Resume Next

	strSQL = "qry_sel_Categories"
	With objCmd
		.Commandtext = strSQL
		.Commandtype = adCmdStoredProc
		Set .ActiveConnection = g_objCn

		.Parameters.Append .CreateParameter("pCategoryID", adInteger, adParamInput, 4, CategoryID)
		objRs.Open objCmd
	End With

	Set objCmd = Nothing

	If Err <> 0 Then
		Set qry_sel_Categories = objRs
	End If
End Function

Function Exec_qry_ins_Categories(CategoryName, Description, Picture)
 Dim strSQL 
 Dim objCmd

	Set objCmd = Server.CreateObject("ADODB.Command")

	On Error Resume Next

	strSQL = "qry_ins_Categories"
	With objCmd
		.Commandtext = strSQL
		.Commandtype = adCmdStoredProc
		Set .ActiveConnection = g_objCn

		.Parameters.Append .CreateParameter("pCategoryName", adVarWChar, adParamInput, 15, CategoryName)
		.Parameters.Append .CreateParameter("pDescription", adLongVarWChar, adParamInput, 2147483647, Description)
		.Parameters.Append .CreateParameter("pPicture", adLongVarBinary, adParamInput, 10737418, Picture)
	
		.Execute Options=adExecuteNoRecords
	End With

	Set objCmd = Nothing

	If Err <> 0 Then
		Exec_qry_ins_Categories = Err.Number
	Else
		Exec_qry_ins_Categories = 0
	End If
End Function

Function Exec_qry_upd_Categories(CategoryID, CategoryName, Description, Picture)
 Dim strSQL 
 Dim objCmd

	Set objCmd = Server.CreateObject("ADODB.Command")

	On Error Resume Next

	strSQL = "qry_upd_Categories"
	With objCmd
		.Commandtext = strSQL
		.Commandtype = adCmdStoredProc
		Set .ActiveConnection = g_objCn

		.Parameters.Append .CreateParameter("pCategoryName", adVarWChar, adParamInput, 15, CategoryName)
		.Parameters.Append .CreateParameter("pDescription", adLongVarWChar, adParamInput, 2147483647, Description)
		.Parameters.Append .CreateParameter("pPicture", adLongVarBinary, adParamInput, 10737418, Picture)
		.Parameters.Append .CreateParameter("pCategoryID", adInteger, adParamInput, 4, CategoryID)
	
		.Execute Options=adExecuteNoRecords
	End With

	Set objCmd = Nothing

	If Err <> 0 Then
		Exec_qry_upd_Categories = Err.Number
	Else
		Exec_qry_upd_Categories = 0
	End If
End Function

%>
