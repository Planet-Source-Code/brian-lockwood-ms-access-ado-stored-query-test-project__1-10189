VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Query-Blaster Test Project"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCustomerID 
      Height          =   285
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   25
      Top             =   840
      Width           =   2475
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   1380
      TabIndex        =   24
      Top             =   4980
      Width           =   915
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   2340
      TabIndex        =   23
      Top             =   4980
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   3300
      TabIndex        =   22
      Top             =   4980
      Width           =   915
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Top             =   4440
      Width           =   2475
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   1680
      TabIndex        =   18
      Top             =   4080
      Width           =   2475
   End
   Begin VB.TextBox txtCountry 
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      Top             =   3720
      Width           =   2475
   End
   Begin VB.TextBox txtPostalCode 
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Top             =   3360
      Width           =   2475
   End
   Begin VB.TextBox txtRegion 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   3000
      Width           =   2475
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   2640
      Width           =   2475
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   2475
   End
   Begin VB.TextBox txtContactTitle 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   1920
      Width           =   2475
   End
   Begin VB.TextBox txtContactName 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   2475
   End
   Begin VB.TextBox txtCompanyName 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   2475
   End
   Begin VB.ComboBox cboCompanyName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2475
   End
   Begin VB.Label Label10 
      Caption         =   "Customer ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Fax 
      Caption         =   "Postal Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Phone"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Country:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Postal Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label txtRegionx 
      Caption         =   "Region"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "City:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Contact Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Company Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Contact Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Company Filter:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   *************************************************
'   Project:  Code technique Demonstration for Access, ADO and stored queries.
'
'   Author:   LockwoodTech    7/28/00  www.lockwoodtech.com
'
'   Credits:  All stored queries and ADO code generated with Lockwoodtech Query-Blaster
'
'   Legal:    Code is free to use and distribute
'   *************************************************

Option Explicit


Private Sub cboCompanyName_Click()
 Dim strSQL     As String
 Dim rs         As New ADODB.Recordset
 Dim lngRetVal  As Long

    If cboCompanyName.ListIndex = -1 Then Exit Sub
    
    '   Pass the recordset in ByRef
    lngRetVal = Exec_qry_sel_Customers(cboCompanyName, rs)
    
    If lngRetVal <> 0 Then
        Exit Sub
    End If
    
    txtCompanyName = rs("CompanyName").Value
    txtContactName = IIf(IsNull(rs("ContactName")), "", rs("ContactName"))
    txtContactTitle = IIf(IsNull(rs("ContactTitle")), "", rs("ContactTitle"))
    txtAddress = IIf(IsNull(rs("Address")), "", rs("Address"))
    txtCity = IIf(IsNull(rs("City")), "", rs("City"))
    txtRegion = IIf(IsNull(rs("Region")), "", rs("region"))
    txtPostalCode = IIf(IsNull(rs("PostalCode")), "", rs("PostalCode"))
    txtCountry = IIf(IsNull(rs("Country")), "", rs("Country"))
    txtPhone = IIf(IsNull(rs("Phone")), "", rs("Phone"))
    txtFax = IIf(IsNull(rs("Fax")), "", rs("Fax"))
    txtCustomerID = rs("CustomerID")
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdDelete_Click()
    If Exec_qry_del_Customers(txtCustomerID) = 0 Then
        Call Clear_Controls
        Call requery_list
        MsgBox "Success", vbInformation, "Results"
    Else
        MsgBox "Failure", vbCritical, "Results"
    End If
End Sub

Private Sub cmdNew_Click()
    Call Clear_Controls
    cboCompanyName.ListIndex = -1
End Sub

Private Sub cmdSave_Click()
    If Me.cboCompanyName.ListIndex = -1 Then
    
        If Exec_qry_ins_Customers(txtCustomerID, txtCompanyName, NullIt(txtContactName), NullIt(txtContactTitle), NullIt(txtAddress), NullIt(txtCity), NullIt(txtRegion), NullIt(txtPostalCode), NullIt(txtCountry), NullIt(txtPhone), NullIt(txtFax)) = 0 Then
                Call requery_list
                MsgBox "Success", vbInformation, "Results"
            Else
                MsgBox "Failure", vbCritical, "Results"
        End If
    
    Else

        If Exec_qry_upd_Customers(txtCustomerID, txtCompanyName, NullIt(txtContactName), NullIt(txtContactTitle), NullIt(txtAddress), NullIt(txtCity), NullIt(txtRegion), NullIt(txtPostalCode), NullIt(txtCountry), NullIt(txtPhone), NullIt(txtFax)) = 0 Then
                MsgBox "Success", vbInformation, "Results"
            Else
                MsgBox "Failure", vbCritical, "Results"
        End If
        
    End If
End Sub

Public Sub Form_Load()
 Dim strSQL     As String
 Dim rs As New ADODB.Recordset
 
    strSQL = "SELECT CustomerID FROM customers"
    rs.Open strSQL, g_objCn
    Do While Not rs.EOF
        cboCompanyName.AddItem rs(0)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Function Clear_Controls()
 Dim ctl As Control
 
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then ctl = ""
    Next ctl
End Function

Private Function requery_list()
 Dim rs As New ADODB.Recordset
 Dim strSQL As String
 
    cboCompanyName.Clear
    
    strSQL = "SELECT CustomerID FROM customers Order by CustomerID"
    rs.Open strSQL, g_objCn
    Do While Not rs.EOF
        cboCompanyName.AddItem rs(0)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Function


Public Function Exec_qry_del_Customers(ByVal varCustomerID As Variant) As Long
 Dim strSQL As String
 Dim objCmd As New ADODB.Command

    On Error GoTo PROC_ERR

    strSQL = "qry_del_Customers"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
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
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
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
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
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
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
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



Private Function NullIt(ctl As Control) As Variant
    If TypeOf ctl Is ListBox Or TypeOf ctl Is ComboBox Then
        If ctl.ListIndex = -1 Then
            NullIt = Null
        Else
            NullIt = ctl.ItemData(ctl.ListIndex)
        End If
    ElseIf TypeOf ctl Is TextBox Then
        If ctl = "" Then
            NullIt = Null
        Else
            NullIt = ctl
        End If
    'Elseif ADD OTHER CONTROLS AS NECESSARY
    End If
End Function































