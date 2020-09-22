Attribute VB_Name = "modRSAP"
Option Explicit


Public Function GetAPBySup(ByVal lSupID As Long, ByVal dMinDate As Date, dMaxDate As Date) As Double

    GetAPBySup = (modRSSup.GetSupBegAP(lSupID) + GetSumPOByDate(lSupID, dMinDate, dMaxDate)) - GetSumPTSByDate(lSupID, dMinDate, dMaxDate)
        
End Function

Private Function GetSumPOByDate(ByVal lSupID As Long, ByVal dMinDate As Date, dMaxDate As Date) As Double

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSumPOByDate = -1
    
    sSQL = "SELECT Sum(tblPO.TotalAmt) AS SumOfTotalAmt" & _
            " From tblPO" & _
            " WHERE tblPO.FK_SupID=" & lSupID & " AND DateValue(tblPO.PODate)>=DateValue(#" & DateValue(dMinDate) & "#) AND DateValue(tblPO.PODate)<=DateValue(#" & DateValue(dMaxDate) & "#)"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAP", "GetSumPOByDate", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetSumPOByDate = 0
        GoTo RAE
    End If
    
    vRS.MoveFirst
        
    GetSumPOByDate = ReadField(vRS.Fields("SumOfTotalAmt"))
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Function


Private Function GetSumPTSByDate(ByVal lSupID As Long, ByVal dMinDate As Date, dMaxDate As Date) As Double

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSumPTSByDate = -1

    sSQL = "SELECT Sum(tblPTS.Amount) AS SumOfAmount" & _
            " From tblPTS" & _
            " WHERE tblPTS.FK_SupID=" & lSupID & " AND DateValue(tblPTS.PTSDate)>=DateValue(#" & DateValue(dMinDate) & "#) AND DateValue(tblPTS.PTSDate)<=DateValue(#" & DateValue(dMaxDate) & "#) AND tblPTS.Cleared=True"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAP", "GetSumPTSByDate", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetSumPTSByDate = 0
        GoTo RAE
    End If
    
    vRS.MoveFirst
        
    GetSumPTSByDate = ReadField(vRS.Fields("SumOfAmount"))
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Function

















Public Function GetAllAP(ByVal dMinDate As Date, dMaxDate As Date) As Double

    GetAllAP = (modRSSup.GetAllSupBegAP + GetAllSumPOByDate(dMinDate, dMaxDate)) - GetAllSumPTSByDate(dMinDate, dMaxDate)
        
End Function

Private Function GetAllSumPOByDate(ByVal dMinDate As Date, dMaxDate As Date) As Double

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetAllSumPOByDate = -1
    
    sSQL = "SELECT Sum(tblPO.TotalAmt) AS SumOfTotalAmt" & _
            " From tblPO" & _
            " WHERE DateValue(tblPO.PODate)>=DateValue(#" & DateValue(dMinDate) & "#) AND DateValue(tblPO.PODate)<=DateValue(#" & DateValue(dMaxDate) & "#)"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAP", "GetAllSumPOByDate", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetAllSumPOByDate = 0
        GoTo RAE
    End If
    
    vRS.MoveFirst
        
    GetAllSumPOByDate = ReadField(vRS.Fields("SumOfTotalAmt"))
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Function


Private Function GetAllSumPTSByDate(ByVal dMinDate As Date, dMaxDate As Date) As Double

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetAllSumPTSByDate = -1

    sSQL = "SELECT Sum(tblPTS.Amount) AS SumOfAmount" & _
            " From tblPTS" & _
            " WHERE DateValue(tblPTS.PTSDate)>=DateValue(#" & DateValue(dMinDate) & "#) AND DateValue(tblPTS.PTSDate)<=DateValue(#" & DateValue(dMaxDate) & "#) AND tblPTS.Cleared=True"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAP", "GetAllSumPTSByDate", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetAllSumPTSByDate = 0
        GoTo RAE
    End If
    
    vRS.MoveFirst
        
    GetAllSumPTSByDate = ReadField(vRS.Fields("SumOfAmount"))
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Function


