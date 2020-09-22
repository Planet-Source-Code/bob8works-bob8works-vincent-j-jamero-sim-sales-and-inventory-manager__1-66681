Attribute VB_Name = "modRSPO"
Option Explicit


Public Type tPO

    POID As Long
    RefNum As String
    FK_SupID As Long
    
    PODate As Date
    
    CA As String
    FP As String
    OptFK_PTSID As Long

    TotalAmt As Double
    PayAmtOnDate As Double
    
    POBalance As Double
    
    Remarks As String
    
    RC As Date
    RM As Date
    RCU As String
    RMU As String
    
End Type


Public Function AddPO(vPO As tPO) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddPO = False
    
    sSQL = "SELECT * FROM tblPO WHERE POID=" & vPO.POID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPO", "AddPO", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        WriteErrorLog "modRSPO", "AddPO", "Adding Failed. Reaseon: Duplication of RefNum or POID"
        GoTo RAE
    End If
    
    
    'add new record
    vRS.AddNew
    
    If WritePO(vRS, vPO) = False Then
        GoTo RAE
    End If
    
    vRS.Update
    
 
    AddPO = True
    
RAE:
    Set vRS = Nothing
    
End Function

Public Function EditPO(vPO As tPO) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim tmpPO As tPO
    
    'default
    EditPO = False
    
    sSQL = "SELECT * FROM tblPO WHERE POID=" & vPO.POID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPO", "EditPO", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If

    If GetPOByID(vPO.POID, tmpPO) = False Then
        WriteErrorLog "modRSPO", "EditPO", "Failed on: 'GetPOByID(vPO.POID, tmpPO) = False'"
        GoTo RAE
    End If
    
    
    'edit
    If WritePO(vRS, vPO) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditPO = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeletePO(ByVal iPOID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    On Error GoTo RAE
    'default
    DeletePO = False
    
    sSQL = "DELETE * FROM tblPO WHERE POID=" & iPOID
    
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPO", "DeletePO", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     
    DeletePO = True
    
RAE:
    Set vRS = Nothing
End Function



Public Function GetPOByID(ByVal iPOID As Long, ByRef vPO As tPO) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetPOByID = False
    
    sSQL = "SELECT * FROM tblPO WHERE POID=" & iPOID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPO", "GetPOByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadPO(vRS, vPO) = False Then
        GoTo RAE
    End If
    
    GetPOByID = True
    
RAE:
    Set vRS = Nothing
End Function




Public Function GetNewPOID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewPOID = -1
    
    sSQL = "SELECT Max(tblPO.POID)+1 AS MaxOfPOID" & _
            " From tblPO"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPO", "GetNewPOID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewPOID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewPOID = ReadField(vRS.Fields("MaxOfPOID"))
    
    If GetNewPOID < 1 Then
        GetNewPOID = 1
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Function


Public Function ReadPO(ByRef vRS As ADODB.Recordset, ByRef vPO As tPO) As Boolean
    
    'default
    ReadPO = False
    
    On Error GoTo RAE
    
    With vPO
        
        .POID = ReadField(vRS.Fields("POID"))
        .RefNum = ReadField(vRS.Fields("RefNum"))
        
        .FK_SupID = ReadField(vRS.Fields("FK_SupID"))
        
        .PODate = ReadField(vRS.Fields("PODate"))
        
        .FP = ReadField(vRS.Fields("FP"))
        .OptFK_PTSID = ReadField(vRS.Fields("OptFK_PTSID"))
        .CA = ReadField(vRS.Fields("CA"))

        .TotalAmt = vRS.Fields("TotalAmt")
        .PayAmtOnDate = ReadField(vRS.Fields("PayAmtOnDate"))

        .POBalance = ReadField(vRS.Fields("POBalance"))

        .Remarks = ReadField(vRS.Fields("Remarks"))
        
        .RC = ReadField(vRS.Fields("RC"))
        .RM = ReadField(vRS.Fields("RM"))
        .RCU = ReadField(vRS.Fields("RCU"))
        .RMU = ReadField(vRS.Fields("RMU"))
        
    End With
    
    ReadPO = True
    Exit Function
    
RAE:
    
End Function

Public Function WritePO(ByRef vRS As ADODB.Recordset, ByRef vPO As tPO) As Boolean
    
    'default
    WritePO = False
    
    On Error GoTo RAE

    With vPO
    
        vRS.Fields("POID") = .POID
        vRS.Fields("RefNum") = .RefNum
        
        vRS.Fields("FK_SupID") = .FK_SupID
        
        vRS.Fields("PODate") = .PODate
        
        vRS.Fields("CA") = .CA
        vRS.Fields("FP") = .FP
        vRS.Fields("OptFK_PTSID") = .OptFK_PTSID
        
        vRS.Fields("TotalAmt") = .TotalAmt
        vRS.Fields("PayAmtOnDate") = .PayAmtOnDate

        vRS.Fields("POBalance") = .POBalance
        vRS.Fields("Remarks") = .Remarks
        
        vRS.Fields("RC") = .RC
        vRS.Fields("RM") = .RM
        vRS.Fields("RCU") = .RCU
        vRS.Fields("RMU") = .RMU
    
    End With

    WritePO = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function


