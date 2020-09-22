Attribute VB_Name = "modRSBank"
Option Explicit


Public Type tBank

    BankID As Long
    BankName As String
    Address As String
    
End Type


Public Function AddBank(ByVal sBankName As String, Optional sAddress As String = "") As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim vBank As tBank
    
    
    'default
    AddBank = False
    
    sSQL = "SELECT * FROM tblBank WHERE BankName='" & sBankName & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSBank", "AddBank", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddBank = True
        GoTo RAE
    End If
    
    'set new Bank
    vBank.BankName = sBankName
    vBank.Address = sAddress
    'get newBank ID
    vBank.BankID = GetNewBankID
    
    'add new record
    vRS.AddNew
    
    If WriteBank(vRS, vBank) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    AddBank = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditBank(vBank As tBank) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditBank = False
    
    sSQL = "SELECT * FROM tblBank WHERE BankID=" & vBank.BankID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSBank", "EditBank", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        WriteErrorLog "modRSBank", "EditBank", "BankID does not exist. BankID= " & vBank.BankID
        GoTo RAE
    End If
    
    'edit
    If WriteBank(vRS, vBank) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditBank = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteBank(ByVal iBankID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    On Error GoTo RAE
    'default
    DeleteBank = False
    
    sSQL = "DELETE * FROM tblBank WHERE BankID=" & iBankID

    Dim sErrD As String
    Dim iErrN As Long
    If ConnectRS(PrimeDB, vRS, sSQL, False, iErrN, sErrD) = False Then
        If iErrN = -2147467259 Then
            'it includes releted data
            MsgBox "Unable to delete entry. It includes other related record." & vbNewLine & vbNewLine & _
                    "Details: " & sErrD, vbExclamation
        Else
            WriteErrorLog "modRSBank", "DeleteBank", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
            GoTo RAE
        End If
    End If
     
    DeleteBank = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function GetBankByName(sBankName As String, vBank As tBank) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetBankByName = False
    
    sSQL = "SELECT * FROM tblBank WHERE BankName='" & sBankName & "'"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSBank", "GetBankByName", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadBank(vRS, vBank) = False Then
        GoTo RAE
    End If
    
    GetBankByName = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function GetBankByID(ByVal iBankID As Long, ByRef vBank As tBank) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetBankByID = False
    
    sSQL = "SELECT * FROM tblBank WHERE BankID=" & iBankID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSBank", "GetBankByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadBank(vRS, vBank) = False Then
        GoTo RAE
    End If
    
    GetBankByID = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function AnyBankExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyBankExist = False
    
    sSQL = "SELECT * FROM tblBank"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSBank", "AnyBankExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyBankExist = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetNewBankID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewBankID = -1
    
    sSQL = "SELECT Max(tblBank.BankID)+1 AS MaxOfBankID" & _
            " From tblBank"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSBank", "GetNewBankID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewBankID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewBankID = ReadField(vRS.Fields("MaxOfBankID"))
    
    If GetNewBankID < 1 Then
        GetNewBankID = 1
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Function


Public Sub FillBankToCMB(ByRef cmb As ComboBox)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    

    sSQL = "SELECT tblBank.BankName" & _
            " From tblBank" & _
            " ORDER BY tblBank.BankName"


    cmb.Clear
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAddress", "FillBankToCMB", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    vRS.MoveFirst
    While vRS.EOF = False
        cmb.AddItem ReadField(vRS.Fields("BankName"))
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    
End Sub


Public Function ReadBank(ByRef vRS As ADODB.Recordset, ByRef vBank As tBank) As Boolean
    
    'default
    ReadBank = False
    
    On Error GoTo RAE
    
    With vBank
        
        .BankID = ReadField(vRS.Fields("BankID"))
        .BankName = ReadField(vRS.Fields("BankName"))
        .Address = ReadField(vRS.Fields("Address"))
        
    End With
    
    ReadBank = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteBank(ByRef vRS As ADODB.Recordset, ByRef vBank As tBank) As Boolean
    
    'default
    WriteBank = False
    
    On Error GoTo RAE

    With vBank
    
        vRS.Fields("BankID") = .BankID
        vRS.Fields("BankName") = .BankName
        vRS.Fields("Address") = .Address

    End With

    WriteBank = True
    Exit Function
    
RAE:
    MsgBox Err.Address
End Function





