Attribute VB_Name = "modMain"
Option Explicit

'api declarations
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long


'public vars
Public CurrentUser As tUser

Public Const DBFileName = "PrimeData.mdb"
Public PrimeDB As New ADODB.Connection

Public DBPathFileName As String





Public Sub Main()

    'use system appearance style
    InitCommonControls
    
    'show author message
    frmMSG.ShowForm
    

    'check license
    'If modLicense.CheckLicense = False Then
    '    End
    'End If


    'init global variables
    Call modGV.InitGV
    
    'set Database Path
    If InitDB = False Then
        Exit Sub
    End If
    
    
    'Show Splash
    frmSplash.ShowSplash
       

End Sub

Public Sub Main_AfterSD()
    

    'Open Database File
    If OpenDB = False Then
        Exit Sub
    End If
     
    
    'TestUnit
    mdiMain.ShowForm
End Sub

Private Sub TestUnit()

    Exit Sub
    
    frmSplash.UnloadSplash

    Dim a As tProd
    Dim ee As Long
    Dim l As Long
    
    ee = GetNewProdID
    
    For l = ee To 200 + ee
        
        a.ProdID = 100 + l
        a.ProdDescription = "Des " & l
        a.FK_CatID = 1
        a.FK_PackID = 1
        a.SRPrice = 100
        a.SupPrice = 95
        a.Active = True
        
        If AddProd(a) = False Then
            MsgBox "Failed"
        End If
    Next
    
    
     
End Sub


Public Function InitDB() As Boolean
    
    Dim FSO As New FileSystemObject
    'default
    InitDB = False
    
    'check database file path
    If FSO.FileExists(App.Path & "\" & DBFileName) = False Then
        'unload splash
        frmSplash.UnloadSplash

        DBPathFileName = App.Path & "\" & DBFileName
        WriteErrorLog "modMain", "InitDB", "Database File Not Found."
        GoTo RAE
    End If
    
    'set path file name
    DBPathFileName = App.Path & "\" & DBFileName

    'return true
    InitDB = True
    
RAE:
    Set FSO = Nothing
End Function

Public Function OpenDB() As Boolean

    OpenDB = False
    
    'open databse
    If ConnectDB(PrimeDB, DBPathFileName) = False Then
        WriteErrorLog "modMain", "InitDB", "Unable to connect databse."
        GoTo RAE
    End If
    
    OpenDB = True
    
RAE:
End Function
