Attribute VB_Name = "modDBMain"


Public Function ConnectDB(ByRef vDB As ADODB.Connection, PathFileName As String) As Boolean

On Error GoTo errh
 
    If vDB.State = adStateOpen Then vDB.Close
        
    vDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathFileName & ";Persist Security Info=False;Jet OLEDB:Database Password="
    
    ConnectDB = True
    
    Exit Function
    
errh:

    WriteErrorLog "modDBMain", "ConnectDB", err.Description
    ConnectDB = False
    
End Function

Public Function CloseDB(ByRef vDB As ADODB.Connection)
    On Error GoTo errh
    vDB.Close
errh:
End Function


Public Function ConnectRS(ByRef vDB As ADODB.Connection, ByRef vRS As ADODB.Recordset, sSQL As String, Optional sHowMSG As Boolean = True, Optional ByRef iErrNumber As Variant, Optional ByRef sErrDescription As Variant) As Boolean
    
On Error GoTo errh

    
    Set vRS = Nothing
    Set vRS = New ADODB.Recordset
  
  
    vRS.Open sSQL, vDB, adOpenStatic, adLockOptimistic
    ConnectRS = True

    
    Exit Function
    
'-------------------------------------------
errh:
    If sHowMSG = True Then
        WriteErrorLog "modDBMain", "ConnectRS", "Unable to connect Recordset / Err: " & err.Description
    End If
    If Not IsMissing(iErrNumber) Then
        iErrNumber = err.Number
    End If
    If Not IsMissing(sErrDescription) Then
        sErrDescription = err.Description
    End If
    ConnectRS = False
End Function


Public Function RecordNoMatch(ByRef vRS As ADODB.Recordset) As Boolean
On Error GoTo errh:

    RecordNoMatch = (vRS.BOF = True Or vRS.EOF = True)

    Exit Function
    
errh:
    RecordNoMatch = False
    
End Function


Public Function AnyRecordExisted(ByRef vRS As ADODB.Recordset) As Boolean
    If vRS.State = adStateClosed Then
        AnyRecordExisted = False
        Exit Function
    End If
    
    
    vRS.Requery
    
    If (vRS.BOF = True) And (vRS.EOF = True) Then
        AnyRecordExisted = False
    Else
        On Error GoTo errh
        vRS.MoveFirst
        AnyRecordExisted = True
    End If

    Exit Function
    '--------------------------
    
errh:
    AnyRecordExisted = False
End Function


Public Function ReadField(ByRef vField As Field) As Variant
    
    On Error GoTo errh

    If Not IsNull(vField.Value) Then
        ReadField = vField.Value
    Else
        Select Case vField.Type
            Case adBigInt
                ReadField = 0
            Case adBinary
                ReadField = 0
            Case adBoolean
                ReadField = False
            Case adByRef 'temp
                ReadField = 0
            Case adBSTR
                ReadField = ""
            Case adChar
                ReadField = ""
            Case adCurrency
                ReadField = 0
            Case adDate
                ReadField = CDate(0)
            Case adDBDate
                ReadField = CDate(0)
            Case adDBTime
                ReadField = FormatDateTime(CDate(0), vbLongTime)
            Case adDBTimeStamp
                ReadField = CDate(0)
            Case adDecimal
                ReadField = 0
            Case adDouble
                ReadField = 0
            Case adEmpty 'temp
                ReadField = ""
            Case adError
                ReadField = 0
            
                
                
                
            Case adNumeric
                ReadField = 0
            Case adDouble
                ReadField = 0
            Case Else
                ReadField = ""
            End Select
    End If
    
    Exit Function
    
errh:
    ReadField = ""
End Function

Public Function getRecordCount(ByRef vRS As ADODB.Recordset) As Long
    If AnyRecordExisted(vRS) Then
        vRS.Requery
        vRS.MoveLast
        getRecordCount = vRS.RecordCount
    Else
        getRecordCount = 0
    End If
End Function

Public Function RSMoveFirst(ByRef vRS As ADODB.Recordset) As Boolean
    If AnyRecordExisted(vRS) Then
        vRS.MoveFirst
        RSMoveFirst = True
    Else
        RSMoveFirst = False
    End If
End Function




