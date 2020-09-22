Attribute VB_Name = "modRSCust"
Option Explicit


Public Type tCust

    CustID As Long
    CustName As String
    CPName As String
    CPPosition As String
    ContactNumber As String
    
    AddrProvince As String
    AddrCity As String
    AddrBrgy As String
    AddrStreet As String

    BegAR As Double
    
    Active As Boolean
    RC As Date
    RM As Date
    RCU As String
    RMU As String
    
End Type


Public Function AddCust(vCust As tCust) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddCust = False
    
    sSQL = "SELECT * FROM tblCust WHERE CustName='" & vCust.CustName & "' OR CustID=" & vCust.CustID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCust", "AddCust", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        WriteErrorLog "modRSCust", "AddCust", "Adding Failed. Reaseon: Duplication of CustName or CustID"
        GoTo RAE
    End If
    
    
    'add new record
    vRS.AddNew
    
    If WriteCust(vRS, vCust) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    AddCust = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditCust(vCust As tCust) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditCust = False
    
    sSQL = "SELECT * FROM tblCust WHERE CustID=" & vCust.CustID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCust", "EditCust", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        WriteErrorLog "modRSCust", "EditCust", "CustID does not exist. CustID= " & vCust.CustID
        GoTo RAE
    End If
    
    'edit
    If WriteCust(vRS, vCust) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditCust = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteCust(ByVal iCustID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sErrD As String
    Dim iErrN As Long
    
    
    On Error GoTo RAE
    'default
    DeleteCust = False
    
    sSQL = "DELETE * FROM tblCust WHERE CustID=" & iCustID
    
    If ConnectRS(PrimeDB, vRS, sSQL, False, iErrN, sErrD) = False Then
        If iErrN = -2147467259 Then
            'it includes releted data
            MsgBox "Unable to delete entry. It includes other related record." & vbNewLine & vbNewLine & _
                    "Details: " & sErrD, vbExclamation
        Else
            WriteErrorLog "modRSCust", "DeleteCust", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
            GoTo RAE
        End If
    End If
    
    
    
     
    DeleteCust = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function GetCustByName(sCustName As String, vCust As tCust) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetCustByName = False
    
    sSQL = "SELECT * FROM tblCust WHERE CustName='" & sCustName & "'"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCust", "GetCustByName", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadCust(vRS, vCust) = False Then
        GoTo RAE
    End If
    
    GetCustByName = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function GetCustByID(ByVal iCustID As Long, ByRef vCust As tCust) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetCustByID = False
    
    sSQL = "SELECT * FROM tblCust WHERE CustID=" & iCustID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCust", "GetCustByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadCust(vRS, vCust) = False Then
        GoTo RAE
    End If
    
    GetCustByID = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function AnyCustExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyCustExist = False
    
    sSQL = "SELECT * FROM tblCust"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCust", "AnyCustExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyCustExist = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetNewCustID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewCustID = -1
    
    sSQL = "SELECT Max(tblCust.CustID)+1 AS MaxOfCustID" & _
            " From tblCust"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCust", "GetNewCustID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewCustID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewCustID = ReadField(vRS.Fields("MaxOfCustID"))
    
    If GetNewCustID < 1 Then
        GetNewCustID = 1
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Function


Public Function GetCustBegAR(ByVal iCustID As Long) As Double
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetCustBegAR = -1
    
    sSQL = "SELECT tblCust.BegAR" & _
            " From tblCust" & _
            " WHERE CustID=" & iCustID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCust", "GetCustBegAR", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetCustBegAR = 0
        GoTo RAE
    End If

    GetCustBegAR = ReadField(vRS.Fields("BegAR"))
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetAllBegAR() As Double
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetAllBegAR = -1
    
    sSQL = "SELECT Sum(tblCust.BegAR) AS SumOfBegAR" & _
            " FROM tblCust"



    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCust", "GetAllBegAR", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetAllBegAR = 0
        GoTo RAE
    End If

    GetAllBegAR = ReadField(vRS.Fields("SumOfBegAR"))
    
RAE:
    Set vRS = Nothing
End Function

Public Function SetCustBegAR(ByVal iCustID As Long, ByVal dNewBegAR As Double) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    SetCustBegAR = False
    
    sSQL = "SELECT *" & _
            " From tblCust" & _
            " WHERE CustID=" & iCustID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCust", "SetCustBegAR", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    On Error GoTo RAE
    
    vRS.MoveFirst
    vRS.Fields("BegAR").Value = dNewBegAR
    vRS.Update
    
    SetCustBegAR = True
    
    
RAE:
    Set vRS = Nothing
End Function





Public Function ReadCust(ByRef vRS As ADODB.Recordset, ByRef vCust As tCust) As Boolean
    
    'default
    ReadCust = False
    
    On Error GoTo RAE
    
    With vCust
        
        .CustID = ReadField(vRS.Fields("CustID"))
        .CustName = ReadField(vRS.Fields("CustName"))
        .CPName = ReadField(vRS.Fields("CPName"))
        .CPPosition = ReadField(vRS.Fields("CPPosition"))
        .ContactNumber = ReadField(vRS.Fields("ContactNumber"))
        .AddrProvince = ReadField(vRS.Fields("AddrProvince"))
        .AddrCity = ReadField(vRS.Fields("AddrCity"))
        .AddrBrgy = ReadField(vRS.Fields("AddrBrgy"))
        .AddrStreet = ReadField(vRS.Fields("AddrStreet"))
        
        .BegAR = ReadField(vRS.Fields("BegAR"))
        
        .Active = ReadField(vRS.Fields("Active"))
        
        .RC = ReadField(vRS.Fields("RC"))
        .RM = ReadField(vRS.Fields("RM"))
        .RCU = ReadField(vRS.Fields("RCU"))
        .RMU = ReadField(vRS.Fields("RMU"))
        
    End With
    
    ReadCust = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteCust(ByRef vRS As ADODB.Recordset, ByRef vCust As tCust) As Boolean
    
    'default
    WriteCust = False
    
    On Error GoTo RAE

    With vCust
    
        vRS.Fields("CustID") = .CustID
        vRS.Fields("CustName") = .CustName
        vRS.Fields("CPName") = .CPName
        vRS.Fields("CPPosition") = .CPPosition
        vRS.Fields("ContactNumber") = .ContactNumber
        vRS.Fields("AddrProvince") = .AddrProvince
        vRS.Fields("AddrCity") = .AddrCity
        vRS.Fields("AddrBrgy") = .AddrBrgy
        vRS.Fields("AddrStreet") = .AddrStreet

        vRS.Fields("BegAR") = .BegAR
        
        vRS.Fields("Active") = .Active
        vRS.Fields("RC") = .RC
        vRS.Fields("RM") = .RM
        vRS.Fields("RCU") = .RCU
        vRS.Fields("RMU") = .RMU

    End With

    WriteCust = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function








