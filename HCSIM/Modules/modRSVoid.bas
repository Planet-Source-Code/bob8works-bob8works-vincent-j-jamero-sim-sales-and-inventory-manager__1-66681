Attribute VB_Name = "modRSVoid"
Option Explicit


Public Type tVoid

    VoidID As Long
    VoidDate As Date
    FK_ProdID As Long
    InvQty As Double
    FK_PackID As Long
    Qty As Double
    
End Type


Public Function AddVoid(ByRef vVoid As tVoid) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'default
    AddVoid = False
    
    sSQL = "SELECT * FROM tblVoid WHERE VoidID=" & vVoid.VoidID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSVoid", "AddVoid", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddVoid = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteVoid(vRS, vVoid) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    AddVoid = True
    
    'Update Inventory
    Call modRSStockInv.ClearStockInvByProd(vVoid.FK_ProdID, vVoid.VoidDate)
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditVoid(vVoid As tVoid) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditVoid = False
    
    sSQL = "SELECT * FROM tblVoid WHERE VoidID=" & vVoid.VoidID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSVoid", "EditVoid", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        WriteErrorLog "modRSVoid", "EditVoid", "VoidID does not exist. VoidID= " & vVoid.VoidID
        GoTo RAE
    End If
    
    'edit
    If WriteVoid(vRS, vVoid) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditVoid = True
    
    'Update Inventory
    Call modRSStockInv.ClearStockInvByProd(vVoid.FK_ProdID, vVoid.VoidDate)
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteVoid(ByVal lVoidID As Long, dVoidDate As Date) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo RAE
    'default
    DeleteVoid = False
    
    sSQL = "DELETE * FROM tblVoid WHERE VoidID=" & lVoidID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSVoid", "DeleteVoid", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     
    DeleteVoid = True
    
    'Update Inventory
    Call modRSStockInv.ClearStockInvByProd(lVoidID, dVoidDate)

    
RAE:
    Set vRS = Nothing
End Function

Public Function GetVoidByID(ByVal iVoidID As Long, ByRef vVoid As tVoid) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetVoidByID = False
    
    sSQL = "SELECT * FROM tblVoid WHERE VoidID=" & iVoidID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSVoid", "GetVoidByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadVoid(vRS, vVoid) = False Then
        GoTo RAE
    End If
    
    GetVoidByID = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function AnyVoidExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyVoidExist = False
    
    sSQL = "SELECT * FROM tblVoid"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSVoid", "AnyVoidExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyVoidExist = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetNewVoidID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewVoidID = -1
    
    sSQL = "SELECT Max(tblVoid.VoidID)+1 AS MaxOfVoidID" & _
            " From tblVoid"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSVoid", "GetNewVoidID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewVoidID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewVoidID = ReadField(vRS.Fields("MaxOfVoidID"))
    
    If GetNewVoidID < 1 Then
        GetNewVoidID = 1
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Function


Public Function ReadVoid(ByRef vRS As ADODB.Recordset, ByRef vVoid As tVoid) As Boolean
    
    'default
    ReadVoid = False
    
    On Error GoTo RAE
    
    With vVoid
        
        .VoidID = ReadField(vRS.Fields("VoidID"))
        .VoidDate = ReadField(vRS.Fields("VoidDate"))
        .FK_ProdID = ReadField(vRS.Fields("FK_ProdID"))
        .InvQty = ReadField(vRS.Fields("InvQty"))
        .FK_PackID = ReadField(vRS.Fields("FK_PackID"))
        .Qty = ReadField(vRS.Fields("Qty"))

    End With
    
    ReadVoid = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteVoid(ByRef vRS As ADODB.Recordset, ByRef vVoid As tVoid) As Boolean
    
    'default
    WriteVoid = False
    
    On Error GoTo RAE

    With vVoid
    
        vRS.Fields("VoidID") = .VoidID
        vRS.Fields("VoidDate") = .VoidDate
        vRS.Fields("FK_ProdID") = .FK_ProdID
        vRS.Fields("InvQty") = .InvQty
        vRS.Fields("FK_PackID") = .FK_PackID
        vRS.Fields("Qty") = .Qty

    
    End With

    WriteVoid = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function





