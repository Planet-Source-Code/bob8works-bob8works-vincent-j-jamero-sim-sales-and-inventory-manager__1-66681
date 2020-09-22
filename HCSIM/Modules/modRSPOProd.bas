Attribute VB_Name = "modRSPOProd"
Option Explicit


Public Type tPOProd

    FK_POID As Long
    FK_ProdID As Long
    FK_PackID As Long
    
    InvQty As Double
    Qty As Double
    UnitPrice As Double
    Amount As Double
    
End Type


Public Function AddPOProd(vPOProd As tPOProd, vPO As tPO) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddPOProd = False
    
    sSQL = "SELECT * FROM tblPOProd WHERE FK_POID=" & vPOProd.FK_POID & " AND FK_ProdID=" & vPOProd.FK_ProdID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPOProd", "AddPOProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        WriteErrorLog "modRSPOProd", "AddPOProd", "Adding Failed. Reaseon: Duplication of POProdCode or FK_ProdID"
        GoTo RAE
    End If
    
    
    'add new record
    vRS.AddNew
    
    If WritePOProd(vRS, vPOProd) = False Then
        GoTo RAE
    End If

    vRS.Update
   
    
    AddPOProd = True
    
    'Update Inventory
    Call modRSStockInv.ClearStockInvByProd(vPOProd.FK_ProdID, vPO.PODate)
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditPOProd(vPOProd As tPOProd, vPO As tPO) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim tmpProduct As tPOProd
    
    'default
    EditPOProd = False
    
    sSQL = "SELECT * FROM tblPOProd WHERE FK_POID=" & vPOProd.FK_POID & " AND FK_ProdID=" & vPOProd.FK_ProdID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPOProd", "EditPOProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If

    If AnyRecordExisted(vRS) = False Then
        WriteErrorLog "modRSPOProd", "EditPOProd", "Product Not Found  |  Failed On: 'AnyRecordExisted(vRS) = False'"
        GoTo RAE
    End If

    'edit
    If WritePOProd(vRS, vPOProd) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditPOProd = True
    
    'Update Inventory
    Call modRSStockInv.ClearStockInvByProd(vPOProd.FK_ProdID, vPO.PODate)


RAE:
    Set vRS = Nothing
End Function


Public Function DeletePOProd(ByVal lFK_ProdID As Long, ByVal lFK_POID As Long, vPO As tPO) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo RAE
    'default
    DeletePOProd = False
    
    sSQL = "DELETE * FROM tblPOProd WHERE FK_POID=" & lFK_POID & " AND FK_ProdID=" & lFK_ProdID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPOProd", "DeletePOProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     
    
    DeletePOProd = True
    
    'Update Inventory
    Call modRSStockInv.ClearStockInvByProd(lFK_ProdID, vPO.PODate)

RAE:
    Set vRS = Nothing
End Function



Public Function DeleteAllPOProd(ByVal lFK_POID As Long, vPO As tPO) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo RAE
    'default
    DeleteAllPOProd = False
    
    sSQL = "DELETE * FROM tblPOProd WHERE FK_POID=" & lFK_POID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPOProd", "DeleteAllPOProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     
    'Update Inventory
    ClearStockInv vPO.PODate
    
    DeleteAllPOProd = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetPOProdByID(ByVal lFK_ProdID As Long, ByVal lFK_POID As Long, ByRef vPOProd As tPOProd) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetPOProdByID = False
    
    sSQL = "SELECT * FROM tblPOProd WHERE FK_POID=" & lFK_POID & " AND FK_ProdID=" & lFK_ProdID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPOProd", "GetPOProdByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadPOProd(vRS, vPOProd) = False Then
        GoTo RAE
    End If
    
    GetPOProdByID = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function AnyPOProdExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyPOProdExist = False
    
    sSQL = "SELECT * FROM tblPOProd"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSPOProd", "AnyPOProdExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyPOProdExist = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function ReadPOProd(ByRef vRS As ADODB.Recordset, ByRef vPOProd As tPOProd) As Boolean
    
    'default
    ReadPOProd = False
    
    On Error GoTo RAE
    
    With vPOProd
        
        .FK_POID = ReadField(vRS.Fields("FK_POID"))
        .FK_ProdID = ReadField(vRS.Fields("FK_ProdID"))
        .FK_PackID = ReadField(vRS.Fields("FK_PackID"))
        .InvQty = ReadField(vRS.Fields("InvQty"))
        .UnitPrice = ReadField(vRS.Fields("UnitPrice"))
        .Amount = ReadField(vRS.Fields("Amount"))
        .Qty = ReadField(vRS.Fields("Qty"))

    End With
    
    ReadPOProd = True
    Exit Function
    
RAE:
    
End Function

Public Function WritePOProd(ByRef vRS As ADODB.Recordset, ByRef vPOProd As tPOProd) As Boolean
    
    'default
    WritePOProd = False
    
    On Error GoTo RAE

    With vPOProd
        
        vRS.Fields("FK_POID") = .FK_POID
        vRS.Fields("FK_ProdID") = .FK_ProdID
        vRS.Fields("FK_PackID") = .FK_PackID
        vRS.Fields("InvQty") = .InvQty
        vRS.Fields("UnitPrice") = .UnitPrice
        vRS.Fields("Amount") = .Amount
        vRS.Fields("Qty") = .Qty

    End With

    WritePOProd = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function


