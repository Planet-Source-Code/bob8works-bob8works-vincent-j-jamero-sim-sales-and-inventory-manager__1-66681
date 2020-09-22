Attribute VB_Name = "modRSSup"
Option Explicit


Public Type tSup
    
    
    SupID As Long
    SupName As String
    CPName As String
    CPPosition As String
    Address As String
    ContactNumber As String
    
    BegAP As Double
    
    Active As Boolean
    RC As Date
    RM As Date
    RCU As String
    RMU As String
    
End Type


Public Function AddSup(vSup As tSup) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddSup = False
    
    sSQL = "SELECT * FROM tblSup WHERE SupName='" & vSup.SupName & "' OR SupID=" & vSup.SupID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSup", "AddSup", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        WriteErrorLog "modRSSup", "AddSup", "Adding Failed. Reaseon: Duplication of SupName or SupID"
        GoTo RAE
    End If
    
    
    'add new record
    vRS.AddNew
    
    If WriteSup(vRS, vSup) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    AddSup = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditSup(vSup As tSup) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditSup = False
    
    sSQL = "SELECT * FROM tblSup WHERE SupID=" & vSup.SupID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSup", "EditSup", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        WriteErrorLog "modRSSup", "EditSup", "SupID does not exist. SupID= " & vSup.SupID
        GoTo RAE
    End If
    
    'edit
    If WriteSup(vRS, vSup) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditSup = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteSup(ByVal iSupID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sErrD As String
    Dim iErrN As Long
    
    On Error GoTo RAE
    'default
    DeleteSup = False
    
    sSQL = "DELETE * FROM tblSup WHERE SupID=" & iSupID
    
    If ConnectRS(PrimeDB, vRS, sSQL, False, iErrN, sErrD) = False Then
        If iErrN = -2147467259 Then
            'it includes releted data
            MsgBox "Unable to delete entry. It includes other related record." & vbNewLine & vbNewLine & _
                    "Details: " & sErrD, vbExclamation
        Else
            WriteErrorLog "modRSSup", "DeleteSup", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
            GoTo RAE
        End If
    End If

    DeleteSup = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function GetSupByName(sSupName As String, vSup As tSup) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSupByName = False
    
    sSQL = "SELECT * FROM tblSup WHERE SupName='" & sSupName & "'"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSup", "GetSupByName", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadSup(vRS, vSup) = False Then
        GoTo RAE
    End If
    
    GetSupByName = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function GetSupByID(ByVal iSupID As Long, ByRef vSup As tSup) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSupByID = False
    
    sSQL = "SELECT * FROM tblSup WHERE SupID=" & iSupID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSup", "GetSupByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadSup(vRS, vSup) = False Then
        GoTo RAE
    End If
    
    GetSupByID = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function AnySupExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnySupExist = False
    
    sSQL = "SELECT * FROM tblSup"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSup", "AnySupExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnySupExist = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetNewSupID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewSupID = -1
    
    sSQL = "SELECT Max(tblSup.SupID)+1 AS MaxOfSupID" & _
            " From tblSup"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSup", "GetNewSupID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewSupID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewSupID = ReadField(vRS.Fields("MaxOfSupID"))
    
    If GetNewSupID < 1 Then
        GetNewSupID = 1
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Function


Public Function GetSupBegAP(ByVal iSupID As Long) As Double
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSupBegAP = -1
    
    sSQL = "SELECT tblSup.BegAP" & _
            " From tblSup" & _
            " WHERE SupID=" & iSupID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSup", "GetSupBegAP", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetSupBegAP = 0
        GoTo RAE
    End If

    GetSupBegAP = ReadField(vRS.Fields("BegAP"))
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetAllSupBegAP() As Double
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetAllSupBegAP = -1
    
    sSQL = "SELECT Sum(tblSup.BegAP) AS SumOfBegAP" & _
            " FROM tblSup"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSup", "GetAllSupBegAP", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetAllSupBegAP = 0
        GoTo RAE
    End If

    GetAllSupBegAP = ReadField(vRS.Fields("SumOfBegAP"))
    
RAE:
    Set vRS = Nothing
End Function



Public Function SetSupBegAP(ByVal iSupID As Long, ByVal dNewBegAP As Double) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    SetSupBegAP = False
    
    sSQL = "SELECT *" & _
            " From tblSup" & _
            " WHERE SupID=" & iSupID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSup", "SetSupBegAP", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    On Error GoTo RAE
    
    vRS.MoveFirst
    vRS.Fields("BegAP").Value = dNewBegAP
    vRS.Update
    
    SetSupBegAP = True
    
    
RAE:
    Set vRS = Nothing
End Function





Public Function ReadSup(ByRef vRS As ADODB.Recordset, ByRef vSup As tSup) As Boolean
    
    'default
    ReadSup = False
    
    On Error GoTo RAE
    
    With vSup
        
        .SupID = ReadField(vRS.Fields("SupID"))
        .SupName = ReadField(vRS.Fields("SupName"))
        .CPName = ReadField(vRS.Fields("CPName"))
        .CPPosition = ReadField(vRS.Fields("CPPosition"))
        .Address = ReadField(vRS.Fields("Address"))
        .ContactNumber = ReadField(vRS.Fields("ContactNumber"))
        
        .BegAP = ReadField(vRS.Fields("BegAP"))
        
        .Active = ReadField(vRS.Fields("Active"))
        
        .RC = ReadField(vRS.Fields("RC"))
        .RM = ReadField(vRS.Fields("RM"))
        .RCU = ReadField(vRS.Fields("RCU"))
        .RMU = ReadField(vRS.Fields("RMU"))
        
    End With
    
    ReadSup = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteSup(ByRef vRS As ADODB.Recordset, ByRef vSup As tSup) As Boolean
    
    'default
    WriteSup = False
    
    On Error GoTo RAE

    With vSup
    
        vRS.Fields("SupID") = .SupID
        vRS.Fields("SupName") = .SupName
        vRS.Fields("CPName") = .CPName
        vRS.Fields("CPPosition") = .CPPosition
        vRS.Fields("Address") = .Address
        vRS.Fields("ContactNumber") = .ContactNumber
        
        vRS.Fields("BegAP") = .BegAP
        
        vRS.Fields("Active") = .Active
        
        vRS.Fields("RC") = .RC
        vRS.Fields("RM") = .RM
        vRS.Fields("RCU") = .RCU
        vRS.Fields("RMU") = .RMU

    End With

    WriteSup = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function






