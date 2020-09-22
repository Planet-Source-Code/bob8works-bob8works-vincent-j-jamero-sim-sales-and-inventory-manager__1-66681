Attribute VB_Name = "modRSProdPack"
Option Explicit


Public Type tProdPack

    FK_PackID As Long
    FK_ProdID As Long
        
    Qty As Double
    
    SupPrice As Double
    SRPrice As Double
    
    'optional membes
    PackTitle As String
End Type


Public Function AddProdPack(vProdPack As tProdPack) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddProdPack = False
    
    sSQL = "SELECT * FROM tblProdPack WHERE FK_PackID=" & vProdPack.FK_PackID & " AND FK_ProdID=" & vProdPack.FK_ProdID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSProdPack", "AddProdPack", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        WriteErrorLog "modRSProdPack", "AddProdPack", "Adding Failed. Reaseon: Duplication of ProdPackCode or FK_ProdID"
        GoTo RAE
    End If
    
    
    'add new record
    vRS.AddNew
    
    If WriteProdPack(vRS, vProdPack) = False Then
        GoTo RAE
    End If

    vRS.Update
   
    
    AddProdPack = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditProdPack(vProdPack As tProdPack) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim tmpProduct As tProdPack
    
    'default
    EditProdPack = False
    
    sSQL = "SELECT * FROM tblProdPack WHERE FK_PackID=" & vProdPack.FK_PackID & " AND FK_ProdID=" & vProdPack.FK_ProdID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSProdPack", "EditProdPack", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If

    If AnyRecordExisted(vRS) = False Then
        WriteErrorLog "modRSProdPack", "EditProdPack", "Product Not Found  |  Failed On: 'AnyRecordExisted(vRS) = False'"
        GoTo RAE
    End If

    'edit
    If WriteProdPack(vRS, vProdPack) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditProdPack = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteProdPack(ByVal lFK_ProdID As Long, ByVal lFK_PackID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo RAE
    'default
    DeleteProdPack = False
    
    sSQL = "DELETE * FROM tblProdPack WHERE FK_PackID=" & lFK_PackID & " AND FK_ProdID=" & lFK_ProdID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSProdPack", "DeleteProdPack", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     
    
    DeleteProdPack = True
    
RAE:
    Set vRS = Nothing
End Function



Public Function DeleteAllProdPack(ByVal lFK_ProdID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo RAE
    'default
    DeleteAllProdPack = False
    
    sSQL = "DELETE * FROM tblProdPack WHERE FK_ProdID=" & lFK_ProdID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSProdPack", "DeleteAllProdPack", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     
    
    DeleteAllProdPack = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetProdPackByID(ByVal lFK_ProdID As Long, ByVal lFK_PackID As Long, ByRef vProdPack As tProdPack) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetProdPackByID = False
    
    sSQL = "SELECT * FROM tblProdPack WHERE FK_PackID=" & lFK_PackID & " AND FK_ProdID=" & lFK_ProdID

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSProdPack", "GetProdPackByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadProdPack(vRS, vProdPack) = False Then
        GoTo RAE
    End If
    
    GetProdPackByID = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function AnyProdPackExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyProdPackExist = False
    
    sSQL = "SELECT * FROM tblProdPack"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSProdPack", "AnyProdPackExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyProdPackExist = True
    
RAE:
    Set vRS = Nothing
End Function






Public Function FillProdPackToTypeArray(ByVal lProdID As Long, ByRef vProdPack() As tProdPack) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    Dim vProd As tProd
    Dim i As Integer
    
    'default
    FillProdPackToTypeArray = False
    
    sSQL = "SELECT tblProdPack.FK_PackID, tblPack.PackTitle, tblProdPack.Qty, tblProdPack.SupPrice, tblProdPack.SRPrice" & _
            " FROM tblPack INNER JOIN tblProdPack ON tblPack.PackID = tblProdPack.FK_PackID" & _
            " WHERE tblProdPack.FK_ProdID=" & lProdID

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSProdPack", "FillProdPackToTypeArray", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If

    If modRSProd.GetProdByID(lProdID, vProd) = False Then
        ReDim vProdPack(0)
    Else
        'resize array
        ReDim vProdPack(modDBMain.getRecordCount(vRS))
    End If
    
    
    
    'add main package
    vProdPack(0).FK_PackID = vProd.FK_PackID
    vProdPack(0).PackTitle = modRSPack.GetPackTitleByID(vProd.FK_PackID)
    vProdPack(0).Qty = 1
    vProdPack(0).SupPrice = vProd.SupPrice
    vProdPack(0).SRPrice = vProd.SRPrice
    
    'add other package/s
    If AnyRecordExisted(vRS) = False Then
        FillProdPackToTypeArray = True
        GoTo RAE
    End If
    
    i = 1
    vRS.MoveFirst
    While vRS.EOF = False
        
        With vProdPack(i)
        
            .FK_ProdID = lProdID
            .FK_PackID = ReadField(vRS.Fields("FK_PackID"))
            .PackTitle = ReadField(vRS.Fields("PackTitle"))
            .Qty = ReadField(vRS.Fields("Qty"))
            .SupPrice = ReadField(vRS.Fields("SupPrice"))
            .SRPrice = ReadField(vRS.Fields("SRPrice"))
            
        End With
        
        i = i + 1
        vRS.MoveNext
    Wend
    
    FillProdPackToTypeArray = True
    
RAE:
    Set vRS = Nothing
End Function



Public Function ReadProdPack(ByRef vRS As ADODB.Recordset, ByRef vProdPack As tProdPack) As Boolean
    
    'default
    ReadProdPack = False
    
    On Error GoTo RAE
    
    With vProdPack
        
        .FK_PackID = ReadField(vRS.Fields("FK_PackID"))
        .FK_ProdID = ReadField(vRS.Fields("FK_ProdID"))

        .SupPrice = ReadField(vRS.Fields("SupPrice"))
        .SRPrice = ReadField(vRS.Fields("SRPrice"))
        .Qty = ReadField(vRS.Fields("Qty"))

    End With
    
    ReadProdPack = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteProdPack(ByRef vRS As ADODB.Recordset, ByRef vProdPack As tProdPack) As Boolean
    
    'default
    WriteProdPack = False
    
    On Error GoTo RAE

    With vProdPack
        
        vRS.Fields("FK_PackID") = .FK_PackID
        vRS.Fields("FK_ProdID") = .FK_ProdID

        vRS.Fields("SupPrice") = .SupPrice
        vRS.Fields("SRPrice") = .SRPrice
        vRS.Fields("Qty") = .Qty

    End With

    WriteProdPack = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function
