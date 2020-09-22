Attribute VB_Name = "basCalls2Queries"
Option Explicit

Public Function Exec_qry_del_Customers(ByVal varCustomerID As Variant) As Long
 Dim strSQL As String
 Dim objCmd As New ADODB.Command

	On Error GoTo PROC_ERR

	strSQL = "qry_del_Customers"
	With objCmd
		.Commandtext = strSQL
		.Commandtype = adCmdStoredProc
		Set .ActiveConnection = g_objCn

		.Parameters.Append .CreateParameter("pCustomerID", adVarWChar, adParamInput, 5, varCustomerID)
	
		.Execute Options:=adExecuteNoRecords
	End With

	Set objCmd = Nothing

	Exec_qry_del_Customers = 0
	Exit Function
PROC_ERR:
	Exec_qry_del_Customers = Err.Number
End Function

Public Function Exec_qry_sel_Customers(ByVal varCustomerID As Variant, ByRef objRs As ADODB.Recordset) As Long
 Dim strSQL As String
 Dim objCmd As New ADODB.Command

	On Error GoTo PROC_ERR

	strSQL = "qry_sel_Customers"
	With objCmd
		.Commandtext = strSQL
		.Commandtype = adCmdStoredProc
		Set .ActiveConnection = g_objCn

		.Parameters.Append .CreateParameter("pCustomerID", adVarWChar, adParamInput, 5, varCustomerID)
		objRs.Open objCmd
	End With

	Set objCmd = Nothing

	Exec_qry_sel_Customers = 0
	Exit Function
PROC_ERR:
	Exec_qry_sel_Customers = Err.Number
End Function

Public Function Exec_qry_ins_Customers(ByVal varCustomerID As Variant, ByVal varCompanyName As Variant, ByVal varContactName As Variant, ByVal varContactTitle As Variant, ByVal varAddress As Variant, ByVal varCity As Variant, ByVal varRegion As Variant, ByVal varPostalCode As Variant, ByVal varCountry As Variant, ByVal varPhone As Variant, ByVal varFax As Variant) As Long
 Dim strSQL As String
 Dim objCmd As New ADODB.Command

	On Error GoTo PROC_ERR

	strSQL = "qry_ins_Customers"
	With objCmd
		.Commandtext = strSQL
		.Commandtype = adCmdStoredProc
		Set .ActiveConnection = g_objCn

		.Parameters.Append .CreateParameter("pCustomerID", adVarWChar, adParamInput, 5, varCustomerID)
		.Parameters.Append .CreateParameter("pCompanyName", adVarWChar, adParamInput, 40, varCompanyName)
		.Parameters.Append .CreateParameter("pContactName", adVarWChar, adParamInput, 30, varContactName)
		.Parameters.Append .CreateParameter("pContactTitle", adVarWChar, adParamInput, 30, varContactTitle)
		.Parameters.Append .CreateParameter("pAddress", adVarWChar, adParamInput, 60, varAddress)
		.Parameters.Append .CreateParameter("pCity", adVarWChar, adParamInput, 15, varCity)
		.Parameters.Append .CreateParameter("pRegion", adVarWChar, adParamInput, 15, varRegion)
		.Parameters.Append .CreateParameter("pPostalCode", adVarWChar, adParamInput, 10, varPostalCode)
		.Parameters.Append .CreateParameter("pCountry", adVarWChar, adParamInput, 15, varCountry)
		.Parameters.Append .CreateParameter("pPhone", adVarWChar, adParamInput, 24, varPhone)
		.Parameters.Append .CreateParameter("pFax", adVarWChar, adParamInput, 24, varFax)
	
		.Execute Options:=adExecuteNoRecords
	End With

	Set objCmd = Nothing

	Exec_qry_ins_Customers = 0
	Exit Function
PROC_ERR:
	Exec_qry_ins_Customers = Err.Number
End Function

Public Function Exec_qry_upd_Customers(ByVal varCustomerID As Variant, ByVal varCompanyName As Variant, ByVal varContactName As Variant, ByVal varContactTitle As Variant, ByVal varAddress As Variant, ByVal varCity As Variant, ByVal varRegion As Variant, ByVal varPostalCode As Variant, ByVal varCountry As Variant, ByVal varPhone As Variant, ByVal varFax As Variant) As Long
 Dim strSQL As String
 Dim objCmd As New ADODB.Command

	On Error GoTo PROC_ERR

	strSQL = "qry_upd_Customers"
	With objCmd
		.Commandtext = strSQL
		.Commandtype = adCmdStoredProc
		Set .ActiveConnection = g_objCn

		.Parameters.Append .CreateParameter("pCustomerID", adVarWChar, adParamInput, 5, varCustomerID)
		.Parameters.Append .CreateParameter("pCompanyName", adVarWChar, adParamInput, 40, varCompanyName)
		.Parameters.Append .CreateParameter("pContactName", adVarWChar, adParamInput, 30, varContactName)
		.Parameters.Append .CreateParameter("pContactTitle", adVarWChar, adParamInput, 30, varContactTitle)
		.Parameters.Append .CreateParameter("pAddress", adVarWChar, adParamInput, 60, varAddress)
		.Parameters.Append .CreateParameter("pCity", adVarWChar, adParamInput, 15, varCity)
		.Parameters.Append .CreateParameter("pRegion", adVarWChar, adParamInput, 15, varRegion)
		.Parameters.Append .CreateParameter("pPostalCode", adVarWChar, adParamInput, 10, varPostalCode)
		.Parameters.Append .CreateParameter("pCountry", adVarWChar, adParamInput, 15, varCountry)
		.Parameters.Append .CreateParameter("pPhone", adVarWChar, adParamInput, 24, varPhone)
		.Parameters.Append .CreateParameter("pFax", adVarWChar, adParamInput, 24, varFax)
	
		.Execute Options:=adExecuteNoRecords
	End With

	Set objCmd = Nothing

	Exec_qry_upd_Customers = 0
	Exit Function
PROC_ERR:
	Exec_qry_upd_Customers = Err.Number
End Function

