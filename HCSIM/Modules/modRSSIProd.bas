Attribute VB_Name = "modRSSIProd"
Option Explicit


Public Type tSIProd

    FK_SIID As Long
    FK_ProdID As Long
    FK_PackID As Long
    
    InvQty As Double
    Qty As Double
    UnitPrice As Double
    Amount As Double
    
End Type


Public Function AddSIProd(vSIProd As tSIProd, vSI As tSI) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddSIProd = False
    
    sSQL = "SELECT * FROM tblSIProd WHERE FK_SIID=" & vSIProd.FK_SIID & " AND FK_ProdID=" & vSIProd.FK_ProdID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSIProd", "AddSIProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        WriteErrorLog "modRSSIProd", "AddSIProd", "Adding Failed. Reaseon: Duplication of SIProdCode or FK_ProdID"
        GoTo RAE
    End If
    
    
    'add new record
    vRS.AddNew
    
    If WriteSIProd(vRS, vSIProd) = False Then
        GoTo RAE
    End If

    vRS.Update
   
    
    AddSIProd = True
    
    'Update Inventory
    Call modRSStockInv.ClearStockInvByProd(vSIProd.FK_ProdID, vSI.SIDate)
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditSIProd(vSIProd As tSIProd, vSI As tSI) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim tmpProduct As tSIProd
    
    'default
    EditSIProd = False
    
    sSQL = "SELECT * FROM tblSIProd WHERE FK_SIID=" & vSIProd.FK_SIID & " AND FK_ProdID=" & vSIProd.FK_ProdID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSIProd", "EditSIProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If

    If AnyRecordExisted(vRS) = False Then
        WriteErrorLog "modRSSIProd", "EditSIProd", "Product Not Found  |  Failed On: 'AnyRecordExisted(vRS) = False'"
        GoTo RAE
    End If

    'edit
    If WriteSIProd(vRS, vSIProd) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditSIProd = True
    
    'Update Inventory
    Call modRSStockInv.ClearStockInvByProd(vSIProd.FK_ProdID, vSI.SIDate)
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteSIProd(ByVal lFK_ProdID As Long, ByVal lFK_SIID As Long, vSI As tSI) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo RAE
    'default
    DeleteSIProd = False
    
    sSQL = "DELETE * FROM tblSIProd WHERE FK_SIID=" & lFK_SIID & " AND FK_ProdID=" & lFK_ProdID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSIProd", "DeleteSIProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     
    
    DeleteSIProd = True
    
    'Update Inventory
    Call modRSStockInv.ClearStockInvByProd(lFK_ProdID, vSI.SIDate)

    
RAE:
    Set vRS = Nothing
End Function



Public Function DeleteAllSIProd(ByVal lFK_SIID As Long, vSI As tSI) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo RAE
    'default
    DeleteAllSIProd = False
    
    sSQL = "DELETE * FROM tblSIProd WHERE FK_SIID=" & lFK_SIID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSIProd", "DeleteAllSIProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     
    
    DeleteAllSIProd = True
    
    'update Stock Inventory
    modRSStockInv.ClearStockInv vSI.SIDate
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetSIProdByID(ByVal lFK_ProdID As Long, ByVal lFK_SIID As Long, ByRef vSIProd As tSIProd) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSIProdByID = False
    
    sSQL = "SELECT * FROM tblSIProd WHERE FK_SIID=" & lFK_SIID & " AND FK_ProdID=" & lFK_ProdID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSIProd", "GetSIProdByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadSIProd(vRS, vSIProd) = False Then
        GoTo RAE
    End If
    
    GetSIProdByID = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function AnySIProdExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnySIProdExist = False
    
    sSQL = "SELECT * FROM tblSIProd"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSSIProd", "AnySIProdExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnySIProdExist = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function ReadSIProd(ByRef vRS As ADODB.Recordset, ByRef vSIProd As tSIProd) As Boolean
    
    'default
    ReadSIProd = False
    
    On Error GoTo RAE
    
    With vSIProd
        
        .FK_SIID = ReadField(vRS.Fields("FK_SIID"))
        .FK_ProdID = ReadField(vRS.Fields("FK_ProdID"))
        .FK_PackID = ReadField(vRS.Fields("FK_PackID"))
        .InvQty = ReadField(vRS.Fields("InvQty"))
        .UnitPrice = ReadField(vRS.Fields("UnitPrice"))
        .Amount = ReadField(vRS.Fields("Amount"))
        .Qty = ReadField(vRS.Fields("Qty"))

    End With
    
    ReadSIProd = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteSIProd(ByRef vRS As ADODB.Recordset, ByRef vSIProd As tSIProd) As Boolean
    
    'default
    WriteSIProd = False
    
    On Error GoTo RAE

    With vSIProd
        
        vRS.Fields("FK_SIID") = .FK_SIID
        vRS.Fields("FK_ProdID") = .FK_ProdID
        vRS.Fields("FK_PackID") = .FK_PackID
        vRS.Fields("InvQty") = .InvQty
        vRS.Fields("UnitPrice") = .UnitPrice
        vRS.Fields("Amount") = .Amount
        vRS.Fields("Qty") = .Qty

    End With

    WriteSIProd = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function



