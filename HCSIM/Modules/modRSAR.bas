Attribute VB_Name = "modRSAR"
Option Explicit


Public Function GetARByCust(ByVal lCustID As Long, ByVal dMinDate As Date, dMaxDate As Date) As Double

    GetARByCust = (modRSCust.GetCustBegAR(lCustID) + GetSumSIByDate(lCustID, dMinDate, dMaxDate)) - GetSumCustPayByDate(lCustID, dMinDate, dMaxDate)
        
End Function

Private Function GetSumSIByDate(ByVal lCustID As Long, ByVal dMinDate As Date, dMaxDate As Date) As Double

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSumSIByDate = -1
    
    sSQL = "SELECT Sum(tblSI.TotalAmt) AS SumOfTotalAmt" & _
            " From tblSI" & _
            " WHERE tblSI.FK_CustID=" & lCustID & " AND DateValue(tblSI.SIDate)>=DateValue(#" & DateValue(dMinDate) & "#) AND DateValue(tblSI.SIDate)<=DateValue(#" & DateValue(dMaxDate) & "#)"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAR", "GetSumSIByDate", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetSumSIByDate = 0
        GoTo RAE
    End If
    
    vRS.MoveFirst
        
    GetSumSIByDate = ReadField(vRS.Fields("SumOfTotalAmt"))
    
RAE:
    Set vRS = Nothing
    Err.Clear

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
















Public Function GetAllAR(ByVal dMinDate As Date, dMaxDate As Date) As Double

    GetAllAR = (modRSCust.GetAllBegAR + GetAllSumSIByDate(dMinDate, dMaxDate)) - GetAllSumCustPayByDate(dMinDate, dMaxDate)
        
End Function

Private Function GetAllSumSIByDate(ByVal dMinDate As Date, dMaxDate As Date) As Double

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetAllSumSIByDate = -1
    
    sSQL = "SELECT Sum(tblSI.TotalAmt) AS SumOfTotalAmt" & _
            " From tblSI" & _
            " WHERE DateValue(tblSI.SIDate)>=DateValue(#" & DateValue(dMinDate) & "#) AND DateValue(tblSI.SIDate)<=DateValue(#" & DateValue(dMaxDate) & "#)"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAR", "GetAllSumSIByDate", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetAllSumSIByDate = 0
        GoTo RAE
    End If
    
    vRS.MoveFirst
        
    GetAllSumSIByDate = ReadField(vRS.Fields("SumOfTotalAmt"))
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Function


Private Function GetAllSumCustPayByDate(ByVal dMinDate As Date, dMaxDate As Date) As Double

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetAllSumCustPayByDate = -1

    sSQL = "SELECT Sum(tblCustPay.Amount) AS SumOfAmount" & _
            " From tblCustPay" & _
            " WHERE DateValue(tblCustPay.CustPayDate)>=DateValue(#" & DateValue(dMinDate) & "#) AND DateValue(tblCustPay.CustPayDate)<=DateValue(#" & DateValue(dMaxDate) & "#) AND tblCustPay.Cleared=True"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAR", "GetAllSumCustPayByDate", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetAllSumCustPayByDate = 0
        GoTo RAE
    End If
    
    vRS.MoveFirst
        
    GetAllSumCustPayByDate = ReadField(vRS.Fields("SumOfAmount"))
    
RAE:
    Set vRS = Nothing
    Err.Clear

End Function



