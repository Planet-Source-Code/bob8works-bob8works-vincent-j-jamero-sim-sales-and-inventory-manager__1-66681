Attribute VB_Name = "modRSStockInv"
Option Explicit

Public Type tStockInv
    StockInvDate As Date
    FK_ProdID As Long
    RunQty As Double
    Compacted As Boolean
End Type



Public Function UpdateProdStock(ByVal lFK_ProdID As Long, ByVal dDate As Date) As Boolean

    Dim dRunQty As Long
    Dim newStockInv As tStockInv
    
    
    'default
    UpdateProdStock = False
    
    
    'delete Old Stock Info
    DeleteStockInv lFK_ProdID, dDate
    
    'get Stock
    'Beg. Inv.  +  Purchase - Sold - Void
    


    dRunQty = modRSProd.GetProdBegInvStock(lFK_ProdID) _
                + GetSumPOInvQtyByDate(lFK_ProdID, CDate(0), dDate) _
                - (GetSumSIInvQtyByDate(lFK_ProdID, CDate(0), dDate) _
                    + GetSumVoidInvQtyByDate(lFK_ProdID, CDate(0), dDate))
                
                
    'set new Stock Inv Info
    With newStockInv
        .StockInvDate = dDate
        .FK_ProdID = lFK_ProdID
        .RunQty = dRunQty
        .Compacted = True
    End With
    
    'add stock
    If AddStockInv(newStockInv) = False Then
        Exit Function
    End If
    
    'update ahead records
    UpdateAheadStockInv lFK_ProdID, dDate
                
    'success
    UpdateProdStock = True
    
    
End Function

Public Sub ClearStockInvByProd(ByVal lFK_ProdID As Long, ByVal dStartDate As Date)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "DELETE * FROM tblStockInv" & _
            " WHERE tblStockInv.FK_ProdID=" & lFK_ProdID & " AND DateValue(tblStockInv.StockInvDate)>=DateValue(#" & dStartDate & "#)"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAP", "ClearStockInvByProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Sub

Public Sub ClearStockInv(ByVal dStartDate As Date)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "DELETE * FROM tblStockInv" & _
            " WHERE DateValue(tblStockInv.StockInvDate)>=DateValue(#" & dStartDate & "#)"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAP", "ClearStockInv", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Sub

Private Sub UpdateAheadStockInv(ByVal lFK_ProdID As Long, ByVal dDate As Date)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "DELETE * FROM tblStockInv" & _
            " WHERE tblStockInv.FK_ProdID=" & lFK_ProdID & " AND DateValue(tblStockInv.StockInvDate)>DateValue(#" & dDate & "#)"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAP", "UpdateAheadStockInv", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Sub





Public Function GetProdStock(ByVal lFK_ProdID As Long, ByVal dDate As Date) As Double

    Dim curStockInv As tStockInv
    Dim iFailedCounter As Integer
        
    'default
    GetProdStock = -1
    iFailedCounter = 0
    
    
GetRetry:
    If iFailedCounter > 5 Then
        GetProdStock = -1
        Exit Function
    End If
    iFailedCounter = iFailedCounter + 1
    
    
    'get old stock inv info
    If GetStockInvByID(lFK_ProdID, dDate, curStockInv) = True Then

        If curStockInv.Compacted = True Then
        
            'success
            GetProdStock = curStockInv.RunQty
            'exit
            
            Exit Function
        End If
    End If
    
    'there is no record yet
    'add record
    If UpdateProdStock(lFK_ProdID, dDate) = True Then
        GoTo GetRetry
    End If
    
End Function


Public Function GetSumPOInvQtyByDate(ByVal lFK_ProdID As Long, ByVal dMinDate As Date, dMaxDate As Date) As Double

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSumPOInvQtyByDate = -1
    
    sSQL = "SELECT Sum(tblPOProd.InvQty) AS SumOfInvQty" & _
            " FROM tblPO INNER JOIN tblPOProd ON tblPO.POID = tblPOProd.FK_POID" & _
            " WHERE tblPOProd.FK_ProdID=" & lFK_ProdID & " AND DateValue(tblPO.PODate)>=DateValue(#" & dMinDate & "#) AND DateValue(tblPO.PODate)<=DateValue(#" & dMaxDate & "#)"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAP", "GetSumPOInvQtyByDate", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetSumPOInvQtyByDate = 0
        GoTo RAE
    End If
    
    vRS.MoveFirst
        
    GetSumPOInvQtyByDate = ReadField(vRS.Fields("SumOfInvQty"))
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Function


Public Function GetSumSIInvQtyByDate(ByVal lFK_ProdID As Long, ByVal dMinDate As Date, dMaxDate As Date) As Double

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSumSIInvQtyByDate = -1
    
    sSQL = "SELECT Sum(tblSIProd.InvQty) AS SumOfInvQty" & _
            " FROM tblSI INNER JOIN tblSIProd ON tblSI.SIID = tblSIProd.FK_SIID" & _
            " WHERE tblSIProd.FK_ProdID=" & lFK_ProdID & " AND DateValue(tblSI.SIDate)>=DateValue(#" & dMinDate & "#) AND DateValue(tblSI.SIDate)<=DateValue(#" & dMaxDate & "#)"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAP", "GetSumSIInvQtyByDate", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetSumSIInvQtyByDate = 0
        GoTo RAE
    End If
    
    vRS.MoveFirst
        
    GetSumSIInvQtyByDate = ReadField(vRS.Fields("SumOfInvQty"))
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Function


Public Function GetSumVoidInvQtyByDate(ByVal lFK_ProdID As Long, ByVal dMinDate As Date, dMaxDate As Date) As Double

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSumVoidInvQtyByDate = -1
   
    sSQL = "SELECT Sum(tblVoid.InvQty) AS SumOfInvQty" & _
            " From tblVoid" & _
            " WHERE tblVoid.FK_ProdID=" & lFK_ProdID & " AND DateValue(tblVoid.VoidDate)>=DateValue(#" & dMinDate & "#) AND DateValue(tblVoid.VoidDate)<=DateValue(#" & dMaxDate & "#)"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAP", "GetSumVoidInvQtyByDate", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetSumVoidInvQtyByDate = 0
        GoTo RAE
    End If
    
    vRS.MoveFirst
        
    GetSumVoidInvQtyByDate = ReadField(vRS.Fields("SumOfInvQty"))
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Function





Private Function AddStockInv(vStockInv As tStockInv) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'default
    AddStockInv = False
    
    sSQL = "SELECT * FROM tblStockInv WHERE FK_ProdID=" & vStockInv.FK_ProdID & " AND DateValue(StockInvDate)=DateValue(#" & DateValue(vStockInv.StockInvDate) & "#)"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSStockInv", "AddStockInv", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        GoTo RAE
    End If
    'add new record
    vRS.AddNew
    
    If WriteStockInv(vRS, vStockInv) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    AddStockInv = True
    
RAE:
    Set vRS = Nothing
End Function



Private Function EditStockInv(vStockInv As tStockInv) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditStockInv = False
    
    sSQL = "SELECT * FROM tblStockInv WHERE FK_ProdID=" & vStockInv.FK_ProdID & " AND DateValue(StockInvDate)=DateValue(#" & vStockInv.StockInvDate & "#)"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSStockInv", "EditStockInv", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        WriteErrorLog "modRSStockInv", "EditStockInv", "StockInvDate does not exist. StockInvDate= " & vStockInv.StockInvDate
        GoTo RAE
    End If
    
    'edit
    If WriteStockInv(vRS, vStockInv) = False Then
        GoTo RAE
    End If
    
    vRS.Update
    
    EditStockInv = True
    
RAE:
    Set vRS = Nothing
End Function



Private Function GetStockInvByID(ByVal lFK_ProdID As Long, ByVal dStockInvDate As Date, ByRef vStockInv As tStockInv) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetStockInvByID = False
    
    sSQL = "SELECT * FROM tblStockInv WHERE FK_ProdID=" & lFK_ProdID & " AND DateValue(StockInvDate)=DateValue(#" & DateValue(dStockInvDate) & "#)"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSStockInv", "GetStockInvByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst

    If ReadStockInv(vRS, vStockInv) = False Then
        GoTo RAE
    End If
    
    GetStockInvByID = True
    
RAE:
    Set vRS = Nothing
End Function


Private Function GetSumCustPayByDate(ByVal lCustID As Long, ByVal dMinDate As Date, dMaxDate As Date) As Double

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSumCustPayByDate = -1

    sSQL = "SELECT Sum(tblCustPay.Amount) AS SumOfAmount" & _
            " From tblCustPay" & _
            " WHERE tblCustPay.FK_CustID=" & lCustID & " AND DateValue(tblCustPay.CustPayDate)>=DateValue(#" & DateValue(dMinDate) & "#) AND DateValue(tblCustPay.CustPayDate)<=DateValue(#" & DateValue(dMaxDate) & "#) AND tblCustPay.Cleared=True"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAR", "GetSumCustPayByDate", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetSumCustPayByDate = 0
        GoTo RAE
    End If
    
    vRS.MoveFirst
        
    GetSumCustPayByDate = ReadField(vRS.Fields("SumOfAmount"))
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Function



Private Function DeleteStockInv(ByVal lFK_ProdID As Long, ByVal dStockInvDate As Date) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo RAE
    'default
    DeleteStockInv = False
    
    sSQL = "DELETE * FROM tblStockInv WHERE FK_ProdID=" & lFK_ProdID & " AND DateValue(StockInvDate)=DateValue(#" & dStockInvDate & "#)"
    
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSStockInv", "DeleteStockInv", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
     

    DeleteStockInv = True
    
RAE:
    Set vRS = Nothing
End Function



Private Function ReadStockInv(ByRef vRS As ADODB.Recordset, ByRef vStockInv As tStockInv) As Boolean
    
    'default
    ReadStockInv = False
    
    'On Error GoTo RAE
    
    With vStockInv

        .StockInvDate = DateValue(ReadField(vRS.Fields("StockInvDate")))
        .FK_ProdID = ReadField(vRS.Fields("FK_ProdID"))
        .RunQty = ReadField(vRS.Fields("RunQty"))
        .Compacted = ReadField(vRS.Fields("Compacted"))

    End With
    
    ReadStockInv = True
    Exit Function
    
RAE:
    
End Function

Private Function WriteStockInv(ByRef vRS As ADODB.Recordset, ByRef vStockInv As tStockInv) As Boolean
    
    'default
    WriteStockInv = False
    
    On Error GoTo RAE

    With vStockInv
    
        vRS.Fields("StockInvDate") = DateValue(.StockInvDate)
        vRS.Fields("FK_ProdID") = .FK_ProdID
        vRS.Fields("RunQty") = .RunQty
        vRS.Fields("Compacted") = .Compacted

    End With

    WriteStockInv = True
    Exit Function
    
RAE:
    MsgBox Err.Description
End Function


