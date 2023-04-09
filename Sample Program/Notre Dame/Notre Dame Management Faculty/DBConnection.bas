Attribute VB_Name = "DBConnection"
Global cnMS As New ADODB.Connection
Global cnMSSQL As New ADODB.Connection
Global cnMSSQLtest As New ADODB.Connection
Global strIP1 As String
Global strIP2 As String
Global dtMRem As Date
Global lngDiff As Double
Sub Main()
    dtMRem = Now
    Set cnMS = New ADODB.Connection
    cnMS.CursorLocation = 1
    cnMS.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\School.mdb" & ";Persist Security Info =false"
    
    Dim RsS As New ADODB.Recordset
    Set RsS = cnMS.Execute("select * from MSServer")
    If Not RsS.EOF Then
        On Error GoTo errMSSQL
        Set cnMSSQL = New ADODB.Connection
        cnMSSQL.CursorLocation = 1
        cnMSSQL.Open "Driver={SQL Server};" & _
               "Server=" & RsS!servername & ";" & _
               "Database=" & RsS!serverdatabase & ";" & _
               "Uid=" & RsS!serveruid & ";" & _
               "Pwd=" & RsS!serverpwd & ";"
    
    End If
    MainForm.Show
    Exit Sub
errMSSQL:
    
    MsgBox "Failed to connect to MSSQL Server.", vbExclamation, "LGVHA"
    frmCamera.Show vbModal
    
    End
End Sub
Function testSQL(strServer As String, strDatabase As String, strUID As String, strPWD As String) As Boolean
        testSQL = True
        On Error GoTo err_Failed
        Set cnMSSQLtest = New ADODB.Connection
        cnMSSQLtest.CursorLocation = 1
        cnMSSQLtest.Open "Driver={SQL Server};" & _
               "Server=" & strServer & ";" & _
               "Database=" & strDatabase & ";" & _
               "Uid=" & strUID & ";" & _
               "Pwd=" & strPWD & ";"
        cnMSSQLtest.Close
            Exit Function
err_Failed:
        testSQL = False
               

End Function
Function ConnectSQL() As Boolean
    Dim RsX As New ADODB.Recordset
    Set RsX = cnMS.Execute("select * from MSServer")
    If Not RsX.EOF Then
        On Error GoTo errMSSQl1
        Set cnMSSQL = New ADODB.Connection
        cnMSSQL.CursorLocation = 1
        cnMSSQL.Open "Driver={SQL Server};" & _
               "Server=" & RsX!servername & ";" & _
               "Database=" & RsX!serverdatabase & ";" & _
               "Uid=" & RsX!serveruid & ";" & _
               "Pwd=" & RsX!serverpwd & ";"
    End If
    ConnectSQL = True
    Exit Function
errMSSQl1:
    ConnectSQL = False
End Function

