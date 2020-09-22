Attribute VB_Name = "modRSCustPay"
Option Explicit


Public Type tCustPay

    CustPayID As Long
    FK_CustID As Long
    
    FP As String
    CustPayDate As Date
    AccountName As String
    CheckNo As String
    
    DateDue As Date
    DateIssued As Date
    
    AccountNo As String
    BankName As String
    Amount As Double
    Remarks As String
    
    Cleared As Boolean
    
    RC As Date
    RM As Date
    RCU As String
    RMU As String
    
End Type


Public Function AddCustPay(vCustPay As tCustPay) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'default
    AddCustPay = False
    
    sSQL = "SELECT * FROM tblCustPay WHERE CustPayID=" & vCustPay.CustPayID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCustPay", "AddCustPay", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        GoTo RAE
    End If

    'add new record
    vRS.AddNew
    
    If WriteCustPay(vRS, vCustPay) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    AddCustPay = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditCustPay(vCustPay As tCustPay) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditCustPay = False
    
    sSQL = "SELECT * FROM tblCustPay WHERE CustPayID=" & vCustPay.CustPayID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCustPay", "EditCustPay", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        WriteErrorLog "modRSCustPay", "EditCustPay", "CustPayID does not exist. CustPayID= " & vCustPay.CustPayID
        GoTo RAE
    End If
    
    'edit
    If WriteCustPay(vRS, vCustPay) = False Then
        GoTo RAE
    End If
    
    vRS.Update
    
    EditCustPay = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteCustPay(ByVal iCustPayID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo RAE
    'default
    DeleteCustPay = False
    
    sSQL = "DELETE * FROM tblCustPay WHERE CustPayID=" & iCustPayID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCustPay", "DeleteCustPay", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     

    DeleteCustPay = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetCustPayByID(ByVal iCustPayID As Long, ByRef vCustPay As tCustPay) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetCustPayByID = False
    
    sSQL = "SELECT * FROM tblCustPay WHERE CustPayID=" & iCustPayID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCustPay", "GetCustPayByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadCustPay(vRS, vCustPay) = False Then
        GoTo RAE
    End If
    
    GetCustPayByID = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetCustPayByCheckNo(ByVal sCheckNo As String, ByVal sBankName As String, ByRef vCustPay As tCustPay) As Boolean

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetCustPayByCheckNo = False
    
    sSQL = "SELECT * FROM tblCustPay WHERE FP='check' AND BankName='" & sBankName & "' AND CheckNo='" & sCheckNo & "'"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCustPay", "GetCustPayByCheckNo", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadCustPay(vRS, vCustPay) = False Then
        GoTo RAE
    End If
    
    GetCustPayByCheckNo = True
    
RAE:
    Set vRS = Nothing
    
End Function


Public Function AnyCustPayExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyCustPayExist = False
    
    sSQL = "SELECT * FROM tblCustPay"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCustPay", "AnyCustPayExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyCustPayExist = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetNewCustPayID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewCustPayID = -1
    
    sSQL = "SELECT Max(tblCustPay.CustPayID)+1 AS MaxOfCustPayID" & _
            " From tblCustPay"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCustPay", "GetNewCustPayID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewCustPayID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewCustPayID = ReadField(vRS.Fields("MaxOfCustPayID"))
    
    If GetNewCustPayID < 1 Then
        GetNewCustPayID = 1
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Function


Public Function ReadCustPay(ByRef vRS As ADODB.Recordset, ByRef vCustPay As tCustPay) As Boolean
    
    'default
    ReadCustPay = False
    
    On Error GoTo RAE
    
    With vCustPay
        
        .CustPayID = ReadField(vRS.Fields("CustPayID"))
        .FK_CustID = ReadField(vRS.Fields("FK_CustID"))
        
        .FP = ReadField(vRS.Fields("FP"))
        
        
        .CustPayDate = ReadField(vRS.Fields("CustPayDate"))
        .AccountName = ReadField(vRS.Fields("AccountName"))
        .CheckNo = ReadField(vRS.Fields("CheckNo"))
        .DateDue = ReadField(vRS.Fields("DateDue"))
        .DateIssued = ReadField(vRS.Fields("DateIssued"))
        .AccountNo = ReadField(vRS.Fields("AccountNo"))
        .BankName = ReadField(vRS.Fields("BankName"))
        .Amount = ReadField(vRS.Fields("Amount"))
        .Remarks = ReadField(vRS.Fields("Remarks"))
        .Cleared = ReadField(vRS.Fields("Cleared"))
        .RC = ReadField(vRS.Fields("RC"))
        .RM = ReadField(vRS.Fields("RM"))
        .RCU = ReadField(vRS.Fields("RCU"))
        .RMU = ReadField(vRS.Fields("RMU"))

    End With
    
    ReadCustPay = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteCustPay(ByRef vRS As ADODB.Recordset, ByRef vCustPay As tCustPay) As Boolean
    
    'default
    WriteCustPay = False
    
    On Error GoTo RAE

    With vCustPay
    
        vRS.Fields("CustPayID") = .CustPayID
        vRS.Fields("FK_CustID") = .FK_CustID
        
        vRS.Fields("FP") = .FP
        
        vRS.Fields("CustPayDate") = .CustPayDate
        
        vRS.Fields("AccountName") = .AccountName
        vRS.Fields("CheckNo") = .CheckNo
        vRS.Fields("DateDue") = .DateDue
        vRS.Fields("DateIssued") = .DateIssued
        vRS.Fields("AccountNo") = .AccountNo
        vRS.Fields("BankName") = .BankName
        vRS.Fields("Amount") = .Amount
        vRS.Fields("Remarks") = .Remarks
        vRS.Fields("Cleared") = .Cleared
        vRS.Fields("RC") = .RC
        vRS.Fields("RM") = .RM
        vRS.Fields("RCU") = .RCU
        vRS.Fields("RMU") = .RMU

    End With

    WriteCustPay = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function


