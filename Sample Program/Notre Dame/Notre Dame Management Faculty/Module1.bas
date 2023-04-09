Attribute VB_Name = "Module1"
Function iSACExist(strAC As String) As Boolean
    Dim Rsv As New ADODB.Recordset
    iSACExist = False
    Set Rsv = cnMSSQL.Execute("select acnumber from userst where acnumber = '" & strAC & "'")
    If Not Rsv.EOF Then
        iSACExist = True
    End If
End Function
Function iSEditACExist(strNewAC As String, strOldAC As String) As Boolean
    Dim Rsv As New ADODB.Recordset
    iSEditACExist = False
    Set Rsv = cnMSSQL.Execute("select acnumber from userst where acnumber = '" & strNewAC & "' and  acnumber <> '" & strOldAC & "'")
    If Not Rsv.EOF Then
        iSEditACExist = True
    End If
End Function

Function iSStudNoExist(strStudNo As String) As Boolean
    Dim Rsv As New ADODB.Recordset
    iSStudNoExist = False
    Set Rsv = cnMSSQL.Execute("select studentnumber from userst where studentnumber = '" & strStudNo & "'")
    If Not Rsv.EOF Then
        iSStudNoExist = True
    End If
End Function

Function iSEditStudNoExist(strNewStudNo As String, strOldStudNo As String) As Boolean
    Dim Rsv As New ADODB.Recordset
    iSEditStudNoExist = False
    Set Rsv = cnMSSQL.Execute("select studentnumber from userst where studentnumber = '" & strNewStudNo & "' and  studentnumber <> '" & strOldStudNo & "'")
    If Not Rsv.EOF Then
        iSEditStudNoExist = True
    End If
End Function

