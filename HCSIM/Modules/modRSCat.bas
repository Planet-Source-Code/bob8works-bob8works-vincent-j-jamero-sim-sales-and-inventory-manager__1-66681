Attribute VB_Name = "modRSCat"
Option Explicit


Public Type tCat

    CatID As Long
    CatTitle As String
    Description As String
    
End Type


Public Function AddCat(ByVal sCatTitle As String, Optional sDescription As String = "") As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim vCat As tCat
    
    
    'default
    AddCat = False
    
    sSQL = "SELECT * FROM tblCat WHERE CatTitle='" & sCatTitle & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCat", "AddCat", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddCat = True
        GoTo RAE
    End If
    
    'set new Category
    vCat.CatTitle = sCatTitle
    vCat.Description = sDescription
    'get newCategory ID
    vCat.CatID = GetNewCatID
    
    'add new record
    vRS.AddNew
    
    If WriteCat(vRS, vCat) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    AddCat = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditCat(vCat As tCat) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditCat = False
    
    sSQL = "SELECT * FROM tblCat WHERE CatID=" & vCat.CatID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCat", "EditCat", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        WriteErrorLog "modRSCat", "EditCat", "CatID does not exist. CatID= " & vCat.CatID
        GoTo RAE
    End If
    
    'edit
    If WriteCat(vRS, vCat) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditCat = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteCat(ByVal iCatID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    On Error GoTo RAE
    'default
    DeleteCat = False
    
    sSQL = "DELETE * FROM tblCat WHERE CatID=" & iCatID

    Dim sErrD As String
    Dim iErrN As Long
    If ConnectRS(PrimeDB, vRS, sSQL, False, iErrN, sErrD) = False Then
        If iErrN = -2147467259 Then
            'it includes releted data
            MsgBox "Unable to delete entry. It includes other related record." & vbNewLine & vbNewLine & _
                    "Details: " & sErrD, vbExclamation
        Else
            WriteErrorLog "modRSCat", "DeleteCat", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
            GoTo RAE
        End If
    End If
     
    DeleteCat = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function GetCatByTitle(sCatTitle As String, vCat As tCat) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetCatByTitle = False
    
    sSQL = "SELECT * FROM tblCat WHERE CatTitle='" & sCatTitle & "'"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCat", "GetCatByTitle", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadCat(vRS, vCat) = False Then
        GoTo RAE
    End If
    
    GetCatByTitle = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function GetCatByID(ByVal iCatID As Long, ByRef vCat As tCat) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetCatByID = False
    
    sSQL = "SELECT * FROM tblCat WHERE CatID=" & iCatID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCat", "GetCatByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadCat(vRS, vCat) = False Then
        GoTo RAE
    End If
    
    GetCatByID = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function AnyCatExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyCatExist = False
    
    sSQL = "SELECT * FROM tblCat"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCat", "AnyCatExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyCatExist = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetNewCatID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewCatID = -1
    
    sSQL = "SELECT Max(tblCat.CatID)+1 AS MaxOfCatID" & _
            " From tblCat"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSCat", "GetNewCatID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewCatID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewCatID = ReadField(vRS.Fields("MaxOfCatID"))
    
    If GetNewCatID < 1 Then
        GetNewCatID = 1
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Function


Public Sub FillCatToCMB(ByRef cmb As ComboBox)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    

    sSQL = "SELECT tblCat.CatTitle" & _
            " From tblCat" & _
            " ORDER BY tblCat.CatTitle"


    cmb.Clear
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAddress", "FillCatToCMB", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    vRS.MoveFirst
    While vRS.EOF = False
        cmb.AddItem ReadField(vRS.Fields("CatTitle"))
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    
End Sub


Public Function ReadCat(ByRef vRS As ADODB.Recordset, ByRef vCat As tCat) As Boolean
    
    'default
    ReadCat = False
    
    On Error GoTo RAE
    
    With vCat
        
        .CatID = ReadField(vRS.Fields("CatID"))
        .CatTitle = ReadField(vRS.Fields("CatTitle"))
        .Description = ReadField(vRS.Fields("Description"))
        
    End With
    
    ReadCat = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteCat(ByRef vRS As ADODB.Recordset, ByRef vCat As tCat) As Boolean
    
    'default
    WriteCat = False
    
    On Error GoTo RAE

    With vCat
    
        vRS.Fields("CatID") = .CatID
        vRS.Fields("CatTitle") = .CatTitle
        vRS.Fields("Description") = .Description

    End With

    WriteCat = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function



