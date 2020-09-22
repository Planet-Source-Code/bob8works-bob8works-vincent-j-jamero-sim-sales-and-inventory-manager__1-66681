Attribute VB_Name = "modRSSI"
Option Explicit


Public Type tSI

    SIID As Long
    RefNum As String
    FK_CustID As Long
    
    SIDate As Date

    OptFK_CustPayID As Long

    TotalAmt As Double
    
    SIBalance As Double
    
    Remarks As String
    
    RC As Date
    RM As Date
    RCU As String
    RMU As String
    
End Type


Public Function AddSI(vSI As tSI) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddSI = False
    
    sSQL = "SELECT * FROM tblSI WHERE SIID=" & vSI.SIID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSI", "AddSI", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        WriteErrorLog "modRSSI", "AddSI", "Adding Failed. Reaseon: Duplication of RefNum or SIID"
        GoTo RAE
    End If
    
    
    'add new record
    vRS.AddNew
    
    If WriteSI(vRS, vSI) = False Then
        GoTo RAE
    End If
    
    vRS.Update
    
 
    AddSI = True
    
RAE:
    Set vRS = Nothing
    
End Function

Public Function EditSI(vSI As tSI) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim tmpSI As tSI
    
    'default
    EditSI = False
    
    sSQL = "SELECT * FROM tblSI WHERE SIID=" & vSI.SIID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSI", "EditSI", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If

    If GetSIByID(vSI.SIID, tmpSI) = False Then
        WriteErrorLog "modRSSI", "EditSI", "Failed on: 'GetSIByID(vSI.SIID, tmpSI) = False'"
        GoTo RAE
    End If
    
    
    'edit
    If WriteSI(vRS, vSI) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditSI = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteSI(ByVal iSIID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    On Error GoTo RAE
    'default
    DeleteSI = False
    
    sSQL = "DELETE * FROM tblSI WHERE SIID=" & iSIID
    
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSI", "DeleteSI", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     
    DeleteSI = True
    
RAE:
    Set vRS = Nothing
End Function



Public Function GetSIByID(ByVal iSIID As Long, ByRef vSI As tSI) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSIByID = False
    
    sSQL = "SELECT * FROM tblSI WHERE SIID=" & iSIID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSI", "GetSIByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadSI(vRS, vSI) = False Then
        GoTo RAE
    End If
    
    GetSIByID = True
    
RAE:
    Set vRS = Nothing
End Function




Public Function GetNewSIID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewSIID = -1
    
    sSQL = "SELECT Max(tblSI.SIID)+1 AS MaxOfSIID" & _
            " From tblSI"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSI", "GetNewSIID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewSIID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewSIID = ReadField(vRS.Fields("MaxOfSIID"))
    
    If GetNewSIID < 1 Then
        GetNewSIID = 1
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Function


Public Function ReadSI(ByRef vRS As ADODB.Recordset, ByRef vSI As tSI) As Boolean
    
    'default
    ReadSI = False
    
    On Error GoTo RAE
    
    With vSI
        
        .SIID = ReadField(vRS.Fields("SIID"))
        .RefNum = ReadField(vRS.Fields("RefNum"))
        .FK_CustID = ReadField(vRS.Fields("FK_CustID"))
        
        .SIDate = ReadField(vRS.Fields("SIDate"))
        
        .OptFK_CustPayID = ReadField(vRS.Fields("OptFK_CustPayID"))

        .TotalAmt = vRS.Fields("TotalAmt")

        .Remarks = ReadField(vRS.Fields("Remarks"))
        
        .RC = ReadField(vRS.Fields("RC"))
        .RM = ReadField(vRS.Fields("RM"))
        .RCU = ReadField(vRS.Fields("RCU"))
        .RMU = ReadField(vRS.Fields("RMU"))
        
    End With
    
    ReadSI = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteSI(ByRef vRS As ADODB.Recordset, ByRef vSI As tSI) As Boolean
    
    'default
    WriteSI = False
    
    On Error GoTo RAE

    With vSI

        vRS.Fields("SIID") = .SIID
        vRS.Fields("RefNum") = .RefNum
        vRS.Fields("FK_CustID") = .FK_CustID
        
        vRS.Fields("SIDate") = .SIDate
        
        vRS.Fields("OptFK_CustPayID") = .OptFK_CustPayID
        
        vRS.Fields("TotalAmt") = .TotalAmt

        vRS.Fields("Remarks") = .Remarks
        
        vRS.Fields("RC") = .RC
        vRS.Fields("RM") = .RM
        vRS.Fields("RCU") = .RCU
        vRS.Fields("RMU") = .RMU
    
    End With

    WriteSI = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function




