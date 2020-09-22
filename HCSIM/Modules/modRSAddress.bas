Attribute VB_Name = "modRSAddress"
Option Explicit

Public Function AddBrgy(ByVal sBrgyTitle As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddBrgy = False
    
    If IsEmpty(sBrgyTitle) Then
        GoTo RAE
    End If
    
    sSQL = "SELECT * FROM tblBrgy WHERE BrgyTitle='" & sBrgyTitle & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAdress", "AddBrgy", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddBrgy = True
        GoTo RAE
    End If

    'add new record
    vRS.AddNew
    
    vRS.Fields("BrgyTitle").Value = sBrgyTitle
    
    vRS.Update
   
    
    AddBrgy = True
    
RAE:
    Set vRS = Nothing
End Function



Public Function AddCity(ByVal sCityTitle As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddCity = False
    
    If IsEmpty(sCityTitle) Then
        GoTo RAE
    End If
    
    sSQL = "SELECT * FROM tblCity WHERE CityTitle='" & sCityTitle & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAdress", "AddCity", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddCity = True
        GoTo RAE
    End If

    'add new record
    vRS.AddNew
    
    vRS.Fields("CityTitle").Value = sCityTitle
    
    vRS.Update
   
    
    AddCity = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function AddProvince(ByVal sProvinceTitle As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddProvince = False
    
    If IsEmpty(sProvinceTitle) Then
        GoTo RAE
    End If
    
    sSQL = "SELECT * FROM tblProvince WHERE ProvinceTitle='" & sProvinceTitle & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAdress", "AddProvince", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddProvince = True
        GoTo RAE
    End If

    'add new record
    vRS.AddNew
    
    vRS.Fields("ProvinceTitle").Value = sProvinceTitle
    
    vRS.Update
   
    
    AddProvince = True
    
RAE:
    Set vRS = Nothing
End Function














Public Function DeleteBrgy(ByVal sBrgyTitle As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    DeleteBrgy = False
    
    sSQL = "DELETE * FROM tblBrgy WHERE BrgyTitle='" & sBrgyTitle & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAdress", "DeleteBrgy", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    DeleteBrgy = True
    
RAE:
    Set vRS = Nothing
End Function



Public Function DeleteCity(ByVal sCityTitle As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    DeleteCity = False
    
    sSQL = "DELETE * FROM tblCity WHERE CityTitle='" & sCityTitle & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAdress", "DeleteCity", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    DeleteCity = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteProvince(ByVal sProvinceTitle As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    DeleteProvince = False
    
    sSQL = "DELETE * FROM tblProvince WHERE ProvinceTitle='" & sProvinceTitle & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAdress", "DeleteProvince", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
        
    DeleteProvince = True
    
RAE:
    Set vRS = Nothing
End Function





Public Sub FillBrgyToCMB(ByRef cmb As ComboBox)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    

    sSQL = "SELECT tblBrgy.BrgyTitle" & _
            " From tblBrgy" & _
            " ORDER BY tblBrgy.BrgyTitle"

    cmb.Clear
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAddress", "FillBrgyToCMB", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    vRS.MoveFirst
    While vRS.EOF = False
        cmb.AddItem ReadField(vRS.Fields("BrgyTitle"))
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    
End Sub



Public Sub FillCityToCMB(ByRef cmb As ComboBox)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    

    sSQL = "SELECT tblCity.CityTitle" & _
            " From tblCity" & _
            " ORDER BY tblCity.CityTitle"

    cmb.Clear
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAddress", "FillCityToCMB", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    vRS.MoveFirst
    While vRS.EOF = False
        cmb.AddItem ReadField(vRS.Fields("CityTitle"))
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    
End Sub



Public Sub FillProvinceToCMB(ByRef cmb As ComboBox)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    

    sSQL = "SELECT tblProvince.ProvinceTitle" & _
            " From tblProvince" & _
            " ORDER BY tblProvince.ProvinceTitle"

    cmb.Clear
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog "modRSAddress", "FillProvinceToCMB", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    vRS.MoveFirst
    While vRS.EOF = False
        cmb.AddItem ReadField(vRS.Fields("ProvinceTitle"))
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    
End Sub

