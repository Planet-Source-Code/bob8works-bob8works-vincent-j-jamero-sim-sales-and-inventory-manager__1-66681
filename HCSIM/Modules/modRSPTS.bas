Attribute VB_Name = "modRSPTS"
Option Explicit


Public Type tPTS

    PTSID As Long
    FK_SupID As Long
    
    FP As String
    PTSDate As Date
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


Public Function AddPTS(vPTS As tPTS) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'default
    AddPTS = False
    
    sSQL = "SELECT * FROM tblPTS WHERE PTSID=" & vPTS.PTSID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPTS", "AddPTS", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        GoTo RAE
    End If

    'add new record
    vRS.AddNew
    
    If WritePTS(vRS, vPTS) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    AddPTS = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditPTS(vPTS As tPTS) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditPTS = False
    
    sSQL = "SELECT * FROM tblPTS WHERE PTSID=" & vPTS.PTSID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPTS", "EditPTS", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        WriteErrorLog "modRSPTS", "EditPTS", "PTSID does not exist. PTSID= " & vPTS.PTSID
        GoTo RAE
    End If
    
    'edit
    If WritePTS(vRS, vPTS) = False Then
        GoTo RAE
    End If
    
    vRS.Update
    
    EditPTS = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeletePTS(ByVal iPTSID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo RAE
    'default
    DeletePTS = False
    
    sSQL = "DELETE * FROM tblPTS WHERE PTSID=" & iPTSID
    
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPTS", "DeletePTS", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     

    DeletePTS = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetPTSByID(ByVal iPTSID As Long, ByRef vPTS As tPTS) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetPTSByID = False
    
    sSQL = "SELECT * FROM tblPTS WHERE PTSID=" & iPTSID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPTS", "GetPTSByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadPTS(vRS, vPTS) = False Then
        GoTo RAE
    End If
    
    GetPTSByID = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetPTSByCheckNo(ByVal sCheckNo As String, ByVal sBankName As String, ByRef vPTS As tPTS) As Boolean

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetPTSByCheckNo = False
    
    sSQL = "SELECT * FROM tblPTS WHERE FP='check' AND BankName='" & sBankName & "' AND CheckNo='" & sCheckNo & "'"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPTS", "GetPTSByCheckNo", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadPTS(vRS, vPTS) = False Then
        GoTo RAE
    End If
    
    GetPTSByCheckNo = True
    
RAE:
    Set vRS = Nothing
    
End Function


Public Function AnyPTSExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyPTSExist = False
    
    sSQL = "SELECT * FROM tblPTS"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPTS", "AnyPTSExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyPTSExist = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetNewPTSID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewPTSID = -1
    
    sSQL = "SELECT Max(tblPTS.PTSID)+1 AS MaxOfPTSID" & _
            " From tblPTS"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPTS", "GetNewPTSID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewPTSID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewPTSID = ReadField(vRS.Fields("MaxOfPTSID"))
    
    If GetNewPTSID < 1 Then
        GetNewPTSID = 1
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Function


Public Function ReadPTS(ByRef vRS As ADODB.Recordset, ByRef vPTS As tPTS) As Boolean
    
    'default
    ReadPTS = False
    
    On Error GoTo RAE
    
    With vPTS
        
        .PTSID = ReadField(vRS.Fields("PTSID"))
        .FK_SupID = ReadField(vRS.Fields("FK_SupID"))
        
        .FP = ReadField(vRS.Fields("FP"))
        
        
        .PTSDate = ReadField(vRS.Fields("PTSDate"))
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
    
    ReadPTS = True
    Exit Function
    
RAE:
    
End Function

Public Function WritePTS(ByRef vRS As ADODB.Recordset, ByRef vPTS As tPTS) As Boolean
    
    'default
    WritePTS = False
    
    On Error GoTo RAE

    With vPTS
    
        vRS.Fields("PTSID") = .PTSID
        vRS.Fields("FK_SupID") = .FK_SupID
        
        vRS.Fields("FP") = .FP
        
        vRS.Fields("PTSDate") = .PTSDate
        
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

    WritePTS = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function





