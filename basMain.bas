Attribute VB_Name = "basMain"
Public g_objCn As New ADODB.Connection

Public g_strConnectionString As String

Sub main()
    g_strConnectionString = App.Path & "\Northwind.mdb"

    g_objCn.Provider = "Microsoft.Jet.OLEDB.4.0"
    g_objCn.Open g_strConnectionString

    frmTest.Show
End Sub



