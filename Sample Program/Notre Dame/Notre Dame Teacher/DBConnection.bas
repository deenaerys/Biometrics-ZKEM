Attribute VB_Name = "DBConnection"
Global cnMS As New ADODB.Connection
Global cnMSSQL As New ADODB.Connection
Global cnMSSQLtest As New ADODB.Connection
Global strIP1 As String
Global strIP2 As String
Global dtMRem As Date
Global dtSRem1 As Date
Global dtSRem2 As Date
Global lngDiff As Double
Global lngDiffPIC1 As Double
Global lngDiffPIC2 As Double
Sub Main()
    dtMRem = Now
    dtSrem = Now
    Set cnMS = New ADODB.Connection
    cnMS.CursorLocation = 1
    cnMS.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\School.mdb" & ";Persist Security Info =false"
    
    Dim RsS As New ADODB.Recordset
    Set RsS = cnMS.Execute("select * from MSServer")
    If Not RsS.EOF Then
        On Error GoTo errMSSQL
        If CheckConnection(RsS!servername) Then
        Set cnMSSQL = New ADODB.Connection
        cnMSSQL.CursorLocation = 1
        cnMSSQL.Open "Driver={SQL Server};" & _
               "Server=" & RsS!servername & ";" & _
               "Database=" & RsS!serverdatabase & ";" & _
               "Uid=" & RsS!serveruid & ";" & _
               "Pwd=" & RsS!serverpwd & ";"
        Else
            GoTo errMSSQL
        End If
    End If
    MainForm.Show
    Exit Sub
errMSSQL:
    
    MsgBox "Failed to connect to MSSQL Server.", vbExclamation, "LGVHA"
    'frmCamera.Show vbModal
    MainForm.Show
    'End
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
        If CheckConnection(CStr(RsX!servername)) Then
            Set cnMSSQL = New ADODB.Connection
            cnMSSQL.CursorLocation = 1
            cnMSSQL.Open "Driver={SQL Server};" & _
                   "Server=" & RsX!servername & ";" & _
                   "Database=" & RsX!serverdatabase & ";" & _
                   "Uid=" & RsX!serveruid & ";" & _
                   "Pwd=" & RsX!serverpwd & ";"
        Else
            GoTo errMSSQl1
        End If
    End If
    ConnectSQL = True
    Exit Function
errMSSQl1:
    ConnectSQL = False
End Function

