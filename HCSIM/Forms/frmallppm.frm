VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllPPM 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "PPM"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmallppm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox bgCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   540
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   7
      Top             =   1050
      Width           =   5760
      Begin b8Controls4.LynxGrid3 listEntries 
         Height          =   2355
         Left            =   870
         TabIndex        =   8
         Top             =   450
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4154
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorBkg    =   16056319
         BackColorSel    =   8438015
         ForeColorSel    =   0
         GridColor       =   11136767
         FocusRectColor  =   33023
         AllowUserResizing=   4
         Striped         =   -1  'True
         SBackColor1     =   16056319
         SBackColor2     =   14940667
      End
      Begin MSComctlLib.ImageList ilList 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmallppm.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin b8Controls4.b83DRect shpLBorder 
         Height          =   3015
         Left            =   360
         Top             =   60
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   5318
         Color1          =   16777215
         Color2          =   16777215
         Color3          =   14737632
         Color4          =   14737632
         BackColor       =   16119285
      End
   End
   Begin VB.PictureBox bgFooter 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   390
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   414
      TabIndex        =   0
      Top             =   4260
      Width           =   6210
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ready"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   150
         Width           =   525
      End
      Begin VB.Label lblRecSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   585
      End
      Begin b8Controls4.b83DRect shpFooter 
         Height          =   525
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   926
         Color1          =   16777215
         Color2          =   16777215
         Color3          =   14737632
         Color4          =   14737632
         BackColor       =   16119285
      End
      Begin b8Controls4.b83DRect shpFooter 
         Height          =   315
         Index           =   2
         Left            =   0
         Top             =   540
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         Color1          =   16777215
         Color2          =   16777215
         Color3          =   14737632
         Color4          =   14737632
         BackColor       =   16119285
      End
   End
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   530
      TabIndex        =   2
      Top             =   390
      Width           =   7950
      Begin b8Controls4.b8DataPicker b8DPSup 
         Height          =   360
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropWinWidth    =   6210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier: "
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   30
         Width           =   675
      End
      Begin b8Controls4.b83DRect shpFooter 
         Height          =   675
         Index           =   0
         Left            =   30
         Top             =   0
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1191
         Color1          =   16777215
         Color2          =   16777215
         Color3          =   14737632
         Color4          =   14737632
         BackColor       =   16119285
      End
   End
   Begin b8Controls4.b8TitleBar b8TB 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   661
      Caption         =   "Purchases / Payments Monitoring"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowColor     =   0
      BorderColor     =   4210752
      BackColor       =   8421504
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New P.O. Entry"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuAddPayment 
         Caption         =   "Add New &Payment Entry"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Modify"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuS01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmAllPPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const POBgColor1 = &HF4FFFF
Private Const POBgColor2 = &HDDFFEF '&HF0D8C9

Dim bReadyToDisplay As Boolean
Dim bFormStarted As Boolean

Public Function ShowForm()


    If bFormStarted = True Then
        modFuncChild.ActivateMDIChildForm Me.Name
        Exit Function
    End If
    bFormStarted = True

    'add form
    mdiMain.AddChild Me
    
    mdiMain.b8DateP.ViewIndex = 1
    
    
    'set display flag
    bReadyToDisplay = True

    'load entries
    Form_Refresh
    
End Function






'----------------------------------------------------------
' Controls Procedures
'----------------------------------------------------------
Private Sub b8DPSup_Change()
    Call Form_Refresh
End Sub


Private Sub listEntries_DblClick()
    Form_Edit
End Sub

Private Sub listEntries_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        
        Me.PopupMenu Me.mnuAction
    End If
End Sub

Private Sub listEntries_RowColChanged()

    RefreshRecInfo

End Sub

'----------------------------------------------------------
' >>> END Controls Procedures
'----------------------------------------------------------







'----------------------------------------------------------
' Form Procedures
'----------------------------------------------------------

Private Sub Form_Activate()
    mdiMain.ActivateChild Me
End Sub

Private Sub Form_Load()
    
    'set list columns
    With listEntries
    
        .Redraw = False
        'PO
        .AddColumn "Date", 110       '0
        .AddColumn "DateID", 0, , lgDate    '1
        .AddColumn "POID", 0        '2
        .AddColumn "Ref. #", 80     '3
        .AddColumn "Supplier", 120  '4
        .AddColumn "CA", 100         '5
        .AddColumn "FP", 60         '6
        .AddColumn "Total Amount", 90, lgAlignRightCenter '7
        .AddColumn "Remarks", 0     '8
        
        'spacer
        .AddColumn "", 10 '9
        
        'PTS
        .AddColumn "Payment ID", 80     '10
        .AddColumn "Date", 0     '11
        .AddColumn "Suplier ID", 0     '12
        .AddColumn "Supplier", 120     '13
        .AddColumn "FP", 60     '14
        .AddColumn "Check No.", 0     '15
        .AddColumn "Account Name", 0     '16
        .AddColumn "Date Issued", 0     '17
        .AddColumn "Date Due", 0     '18
        .AddColumn "Account No.", 0     '19
        .AddColumn "Bank", 0     '20
        .AddColumn "Amount", 90, lgAlignRightCenter     '21
        .AddColumn "Cleared", 60, lgAlignCenterCenter     '22
        .AddColumn "Remarks", 0     '23
        
        .AddColumn "RC", 0     '24
        .AddColumn "RM", 0     '25
        .AddColumn "RCU", 0     '26
        .AddColumn "RMU", 0     '27
        
        
        
        .RowHeightMin = 21
        .ImageList = ilList
        .Redraw = True
        .Refresh
    End With
    
    'set supplier list
    With b8DPSup
        Set .DropDBCon = PrimeDB
        .SQLFields = "String(10-Len(Trim([SupID])),'0') & [SupID] AS CSupID, tblSup.SupName"
        .SQLTable = "tblSup"
        .SQLWhereFields = "tblSup.SupID, tblSup.SupName"
        .SQLOrderBy = "tblSup.SupName"
        
        .BoundFieldIndex = 0 'Bound Index
        .DisplayFieldIndex = 1
        .AddColumn "Supplier ID", 100
        .AddColumn "Supplier", 240
    End With
End Sub

Private Sub Form_Resize()
    
    Dim i As Integer


    On Error Resume Next
    
    'resize title bar
    b8TB.Move 2, 1, Me.ScaleWidth - 4
    'resize header
    bgHeader.Move 0, b8TB.Top + b8TB.Height, Me.ScaleWidth
    'resize footer
    bgFooter.Move 0, Me.ScaleHeight - bgFooter.Height, Me.ScaleWidth
    
    For i = 0 To shpFooter.UBound
    shpFooter(i).Move 2, shpFooter(i).Top, bgFooter.Width - 4, shpFooter(i).Height
    Next
    
    
    'resize center
    bgCenter.Move 0, bgHeader.Top + bgHeader.Height, Me.ScaleWidth, bgFooter.Top - (bgHeader.Top + bgHeader.Height)

    'resize list
    shpLBorder.Move 2, 0, bgCenter.Width - 4, bgCenter.Height - 0
    listEntries.Redraw = False
    listEntries.Move shpLBorder.Left + 3, shpLBorder.Top + 3, shpLBorder.Width - 6, shpLBorder.Height - 6
    'listEntries.Move shpLBorder.Left + 1, shpLBorder.Top + 1, shpLBorder.Width - 2, shpLBorder.Height - 2
    listEntries.Redraw = True
    listEntries.Refresh
    
    
    Err.Clear
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'remove form
    mdiMain.RemoveChild Me.Name
    'clear flag
    bFormStarted = False
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF3 Or (KeyCode = 107 And Shift = 2) Then
        'F3 | Ctrl + '+' - add
        If Form_CanAdd Then
            Form_Add
        End If
    
    ElseIf KeyCode = vbKeyF2 Or (KeyCode = 13 And Shift = 2) Then
        'F2 | Ctrl + Enter - edit
        If Form_CanEdit Then
            Form_Edit
        End If
        
    ElseIf KeyCode = vbKeyDelete Or (KeyCode = 109 And Shift = 2) Then
        'Del : Ctrl + '-' - delete
        If Form_CanDelete Then
            Form_Delete
        End If
    
    ElseIf KeyCode = vbKeyF5 Then
        'F5 - refresh
        If Form_CanRefresh Then
            Form_Refresh
        End If
        
    ElseIf KeyCode = 83 And Shift = 2 Then
        'Ctrl + S - Search
        If Form_CanSearch Then
            mdiMain.Form_ShowSearch
        End If
    ElseIf KeyCode = 68 And Shift = 2 Then
        'Ctrl + D - Date Filter
            mdiMain.Form_ShowDateFilter
            
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'pass keyinfo to mdiMain
    mdiMain.AFForm_KeyDown KeyCode, Shift
End Sub

'----------------------------------------------------------
' >>> END Form Procedures
'----------------------------------------------------------






'----------------------------------------------------------
' Record Procedures
'----------------------------------------------------------

Private Sub LoadPO(ByRef vRS As ADODB.Recordset)
    
    Dim sSQL As String
   
    
    If IsEmpty(b8DPSup.BoundData) Then
        'without supplier
        sSQL = "SELECT tblPO.PODate, tblPO.POID, tblPO.RefNum, tblSup.SupName, tblPO.CA, tblPO.FP, tblPO.TotalAmt, tblPO.Remarks" & _
                " FROM tblSup INNER JOIN tblPO ON tblSup.SupID = tblPO.FK_SupID" & _
                " WHERE DateValue(tblPO.PODate)>=DateValue(#" & DateValue(mdiMain.b8DateP.MinDate) & "#) AND DateValue(tblPO.PODate)<=DateValue(#" & DateValue(mdiMain.b8DateP.MaxDate) & "#)" & _
                " ORDER BY tblPO.PODate"
    Else
        'with supplier
        sSQL = "SELECT tblPO.PODate, tblPO.POID, tblPO.RefNum, tblSup.SupName, tblPO.CA, tblPO.FP, tblPO.TotalAmt, tblPO.Remarks" & _
                " FROM tblSup INNER JOIN tblPO ON tblSup.SupID = tblPO.FK_SupID" & _
                " WHERE tblPO.FK_SupID=" & CLng(b8DPSup.BoundData) & " AND DateValue(tblPO.PODate)>=DateValue(#" & DateValue(mdiMain.b8DateP.MinDate) & "#) AND DateValue(tblPO.PODate)<=DateValue(#" & DateValue(mdiMain.b8DateP.MaxDate) & "#)" & _
                " ORDER BY tblPO.PODate"

    End If
             
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "LoadPO", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If

RAE:

End Sub

Private Sub LoadPTS(ByRef vRS As ADODB.Recordset)

    Dim sSQL As String
    
    If IsEmpty(b8DPSup.BoundData) Then
        'without supplier
        sSQL = "SELECT tblPTS.PTSID, tblPTS.FK_SupID, tblSup.SupName, tblPTS.FP, tblPTS.PTSDate, tblPTS.CheckNo, tblPTS.AccountName, tblPTS.DateIssued, tblPTS.DateDue, tblPTS.AccountNo, tblPTS.BankName, tblPTS.Amount, tblPTS.Remarks, tblPTS.Cleared, tblPTS.RC, tblPTS.RM, tblPTS.RCU, tblPTS.RMU" & _
                " FROM tblSup INNER JOIN tblPTS ON tblSup.SupID = tblPTS.FK_SupID" & _
                " WHERE DateValue(tblPTS.PTSDate)>=DateValue(#" & DateValue(mdiMain.b8DateP.MinDate) & "#) AND DateValue(tblPTS.PTSDate)<=DateValue(#" & DateValue(mdiMain.b8DateP.MaxDate) & "#)" & _
                " ORDER BY tblPTS.PTSDate"
    Else
    
        'with supplier
        sSQL = "SELECT tblPTS.PTSID, tblPTS.FK_SupID, tblSup.SupName, tblPTS.FP, tblPTS.PTSDate, tblPTS.CheckNo, tblPTS.AccountName, tblPTS.DateIssued, tblPTS.DateDue, tblPTS.AccountNo, tblPTS.BankName, tblPTS.Amount, tblPTS.Remarks, tblPTS.Cleared, tblPTS.RC, tblPTS.RM, tblPTS.RCU, tblPTS.RMU" & _
                " FROM tblSup INNER JOIN tblPTS ON tblSup.SupID = tblPTS.FK_SupID" & _
                " WHERE tblPTS.FK_SupID=" & CLng(b8DPSup.BoundData) & " AND DateValue(tblPTS.PTSDate)>=DateValue(#" & DateValue(mdiMain.b8DateP.MinDate) & "#) AND DateValue(tblPTS.PTSDate)<=DateValue(#" & DateValue(mdiMain.b8DateP.MaxDate) & "#)" & _
                " ORDER BY tblPTS.PTSDate"
                
    End If
             
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "LoadPTS", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
                
RAE:
End Sub

Private Sub RefreshRecInfo()

    lblRecSum.Caption = "No selected"
    Me.mnuDelete.Caption = "Delete"
    Me.mnuEdit.Caption = "Edit"
    Me.mnuDelete.Enabled = False
    Me.mnuEdit.Enabled = False

    
    If listEntries.RowCount > 0 Then
        If IsEmpty(listEntries.CellText(listEntries.Row, 0)) Then
            If listEntries.Col >= 0 And listEntries.Col <= 8 Then
                'PO
                If Not IsEmpty(listEntries.CellText(listEntries.Row, 2)) Then
                    lblRecSum.Caption = "Selected P.O. entry Ref. # " & listEntries.CellText(listEntries.Row, 2)
                    Me.mnuDelete.Caption = "Delete P.O. entry"
                    Me.mnuEdit.Caption = "Edit P.O. entry"
                    Me.mnuDelete.Enabled = True
                    Me.mnuEdit.Enabled = True
                    
                End If
            ElseIf listEntries.Col >= 10 And listEntries.Col <= 27 Then
                'PTS
                If Not IsEmpty(listEntries.CellText(listEntries.Row, 10)) Then
                    lblRecSum.Caption = "Selected Payment entry ID # " & listEntries.CellText(listEntries.Row, 10)
                    Me.mnuDelete.Caption = "Delete Payment entry"
                    Me.mnuEdit.Caption = "Edit Payment entry"
                    Me.mnuDelete.Enabled = True
                    Me.mnuEdit.Enabled = True
                End If
            End If
            
        Else
            'date
            lblRecSum.Caption = listEntries.CellText(listEntries.Row, 0) & " Transactions"
        End If
    Else
        'no record
        Me.mnuDelete.Enabled = False
        Me.mnuEdit.Enabled = False
        lblRecSum.Caption = "No Record"
    End If
End Sub
'----------------------------------------------------------
' >>> END Record Procedures
'----------------------------------------------------------









'----------------------------------------------------------
' Parent Form Calling Functions
'----------------------------------------------------------

Public Function Form_CanAdd() As Boolean

    Form_CanAdd = True

End Function
Public Function Form_Add()
    
    If frmPOEntry.ShowAdd(mdiMain.b8DateP.MaxDate, CLng(GetTxtVal(b8DPSup.BoundData))) = True Then
        Form_Refresh
    End If
    
End Function

Public Function Form_CanEdit() As Boolean
    If listEntries.RowCount > 0 Then
        Form_CanEdit = True
    Else
        Form_CanEdit = False
    End If
End Function

Public Function Form_Edit()

    Dim vID As Variant
    
    If listEntries.RowCount < 1 Then
        Exit Function
    End If


    If IsEmpty(listEntries.CellText(listEntries.Row, 0)) Then
        If listEntries.Col >= 0 And listEntries.Col <= 8 Then
            'PO
            If Not IsEmpty(listEntries.CellText(listEntries.Row, 2)) Then
                
                vID = GetTxtVal(listEntries.CellText(listEntries.Row, 2))
                If frmPOEntry.ShowEdit(CLng(vID)) = True Then
                    Form_Refresh
                End If
                
            End If
        ElseIf listEntries.Col >= 10 And listEntries.Col <= 27 Then
            'PTS
            If Not IsEmpty(listEntries.CellText(listEntries.Row, 10)) Then
                
                vID = GetTxtVal(listEntries.CellText(listEntries.Row, 10))
                If frmPTSEntry.ShowEdit(CLng(vID)) = True Then
                    Form_Refresh
                End If
                
            End If
        End If
    End If

    
End Function

Public Function Form_CanDelete() As Boolean
    If listEntries.RowCount > 0 Then
        Form_CanDelete = True
    Else
        Form_CanDelete = False
    End If
End Function

Public Function Form_Delete()
    
    Dim vID As Long
    Dim vPO As tPO
    Dim tmpPTS As tPTS
    
    If listEntries.RowCount < 1 Then
        Exit Function
    End If
    
    If IsEmpty(listEntries.CellText(listEntries.Row, 0)) Then
        If listEntries.Col >= 0 And listEntries.Col <= 8 Then
            'PO
            If Not IsEmpty(listEntries.CellText(listEntries.Row, 2)) Then
                
                If MsgBox("Are you sure you want to delete this P.O. entry with Ref.# " & listEntries.CellText(listEntries.Row, 3) & " ?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                    Exit Function
                End If
                'get ID
                vID = listEntries.CellText(listEntries.Row, 2)
            
                If GetPOByID(CLng(vID), vPO) = False Then
                    MsgBox "PO entry notfound.", vbExclamation
                    Exit Function
                End If
                
                If DeletePO(vPO.POID) = True Then
                    If GetPTSByID(vPO.OptFK_PTSID, tmpPTS) = True Then
                        If MsgBox("Do you want to delete Payment entry that associated with this P.O. entry?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                            'delete associated PTS
                            If modRSPTS.DeletePTS(vPO.OptFK_PTSID) = False Then
                                WriteErrorLog Me.Name, "Form_Delete", "Failed on: 'modRSPTS.DeletePTS(vPO.OptFK_PTSID) = False'"
                            End If
                        End If
                    End If
                    
                    Form_Refresh
                Else
                    WriteErrorLog Me.Name, "Form_Delete", "Failed on: 'DeletePO(CLng(vID)) = True'"
                End If
    
            End If
        ElseIf listEntries.Col >= 10 And listEntries.Col <= 27 Then
            'PTS
            If Not IsEmpty(listEntries.CellText(listEntries.Row, 10)) Then
                If MsgBox("Are you sure you want to delete this Payment entry with ID " & listEntries.CellText(listEntries.Row, 10) & " ?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                    Exit Function
                End If
                'get ID
                vID = listEntries.CellText(listEntries.Row, 10)
            
                If DeletePTS(CLng(vID)) = True Then
                    Form_Refresh
                Else
                    WriteErrorLog Me.Name, "Form_Delete", "Failed on: 'DeletePTS(CLng(vID)) = True'"
                End If
            End If
        End If
        
    Else
        MsgBox "Please select 'P.O.' entry or 'Payment' entry.", vbExclamation
    End If

End Function

Public Function Form_CanRefresh() As Boolean
    Form_CanRefresh = True
End Function

Public Function Form_Refresh()

    Dim vPORS As New ADODB.Recordset
    Dim vPTSRS As New ADODB.Recordset
    Dim li As Long
    Dim lhi As Long
    Dim dD As Date
    Dim bAdded As Boolean
    
    Dim dPODate As Date
    Dim dPTSDate As Date
    Dim dOldDate As Date
    
    Dim sFP As String
    
        
    
    If bReadyToDisplay = False Then
        GoTo RAE
    End If
    
    
    'set app mouse icon
    mdiMain.Form_StartBussy
    
    'clear info
    lblInfo.Caption = ""
    'clear list
    listEntries.Redraw = False
    listEntries.Clear
    
    lblRecSum.Caption = "Loading Data. Please wait..."
    DoEvents
    
    'load data
    Call LoadPO(vPORS)
    Call LoadPTS(vPTSRS)
    
    
    
    If AnyRecordExisted(vPORS) = False And AnyRecordExisted(vPTSRS) = False Then
        GoTo RAE
    End If
    
    If AnyRecordExisted(vPORS) = True Then
        vPORS.MoveFirst
    End If
    If AnyRecordExisted(vPTSRS) = True Then
        vPTSRS.MoveFirst
    End If
    
    dD = mdiMain.b8DateP.MinDate
    
    While vPORS.EOF = False Or vPTSRS.EOF = False

        bAdded = False

        'add PO
        If vPORS.EOF = False Then
            dPODate = ReadField(vPORS.Fields("PODate"))
            If DateValue(dD) = DateValue(dPODate) Then
            
                'add Header Date
                If DateValue(dOldDate) <> DateValue(dD) Then
                    lhi = listEntries.AddItem(Format(dD, modGV.GV_DateFormat))
                    listEntries.FillBackColor lhi, lhi, 0, 27, &HF5F5F5
                End If
        
                'same, then add
                With listEntries
                li = .AddItem("")
                .FillBackColor li, li, 0, 8, POBgColor2
                .ItemImage(li) = 1
                .CellText(li, 1) = ReadField(vPORS.Fields("PODate"))
                .CellText(li, 2) = ReadField(vPORS.Fields("POID"))
                .CellText(li, 3) = modFunction.ComNumZ(ReadField(vPORS.Fields("RefNum")), 10)
                .CellText(li, 4) = ReadField(vPORS.Fields("SupName"))
                .CellText(li, 5) = LCase(ReadField(vPORS.Fields("CA")))
                .CellText(li, 6) = sFP = LCase(ReadField(vPORS.Fields("FP")))
                .CellFontBold(li, 7) = True
                .CellText(li, 7) = FormatNumber(ReadField(vPORS.Fields("TotalAmt")), 2)
                .CellText(li, 8) = ReadField(vPORS.Fields("Remarks"))
                End With
                
                vPORS.MoveNext
                
                'set flag
                bAdded = True
            End If
        End If
        
        'add PTS
        If vPTSRS.EOF = False Then
            dPTSDate = ReadField(vPTSRS.Fields("PTSDate"))
            
            If DateValue(dD) = DateValue(dPTSDate) Then
            
                'same, then add
                With listEntries
                If bAdded = False Then
                    'add Header Date
                    If DateValue(dOldDate) <> DateValue(dD) Then
                        lhi = listEntries.AddItem(Format(dD, modGV.GV_DateFormat))
                        listEntries.FillBackColor lhi, lhi, 0, 27, &HA9EEFF
                    End If
                    
                    li = .AddItem(" ")
                End If
                
                .FillBackColor li, li, 10, 27, POBgColor1
                .ItemImage(li) = 1
                .CellText(li, 10) = modFunction.ComNumZ(ReadField(vPTSRS.Fields("PTSID")), 10)
                .CellText(li, 11) = dPTSDate
                .CellText(li, 12) = ReadField(vPTSRS.Fields("FK_SupID"))
                .CellText(li, 13) = ReadField(vPTSRS.Fields("SupName"))
                sFP = LCase(ReadField(vPTSRS.Fields("FP")))
                .CellText(li, 14) = sFP
                .CellText(li, 15) = ReadField(vPTSRS.Fields("CheckNo"))
                .CellText(li, 16) = ReadField(vPTSRS.Fields("AccountName"))
                .CellText(li, 17) = ReadField(vPTSRS.Fields("DateIssued"))
                .CellText(li, 18) = ReadField(vPTSRS.Fields("DateDue"))
                .CellText(li, 19) = ReadField(vPTSRS.Fields("AccountNo"))
                .CellText(li, 20) = ReadField(vPTSRS.Fields("BankName"))
                .CellFontBold(li, 21) = True
                .CellText(li, 21) = FormatNumber(ReadField(vPTSRS.Fields("Amount")), 2)
                .CellText(li, 22) = IIf(sFP = "check", IIf(ReadField(vPTSRS.Fields("Cleared")), "Yes", "No"), "N/A")
                .CellText(li, 23) = ReadField(vPTSRS.Fields("Remarks"))
                .CellText(li, 24) = ReadField(vPTSRS.Fields("RC"))
                .CellText(li, 25) = ReadField(vPTSRS.Fields("RM"))
                .CellText(li, 26) = ReadField(vPTSRS.Fields("RCU"))
                .CellText(li, 27) = ReadField(vPTSRS.Fields("RMU"))

                End With
                
                vPTSRS.MoveNext
                
                'set flag
                bAdded = True
            End If
        End If
        
        'get old date
        dOldDate = dD

        If bAdded = False Then
            dD = dD + 1
        End If

    Wend

    

RAE:
    
    lblRecSum.Caption = "Calculating A/P..."
    DoEvents
    'set fotter info
    If Not IsEmpty(b8DPSup.BoundData) Then
        lblInfo.Caption = "A/P (Beg): " & _
                            FormatNumber(modRSAP.GetAPBySup(CLng(b8DPSup.BoundData), CDate(0), mdiMain.b8DateP.MinDate - 1), 2) & _
                            "   |   A/P (End): " & _
                            FormatNumber(modRSAP.GetAPBySup(CLng(b8DPSup.BoundData), CDate(0), mdiMain.b8DateP.MaxDate), 2)
    Else
        
        lblInfo.Caption = "For all Supplier, A/P (Beg): " & _
                            FormatNumber(modRSAP.GetAllAP(CDate(0), mdiMain.b8DateP.MinDate - 1), 2) & _
                            "   |   A/P (End): " & _
                            FormatNumber(modRSAP.GetAllAP(CDate(0), mdiMain.b8DateP.MaxDate), 2)
        
    End If
    lblRecSum.Caption = "Ready"


    listEntries.Redraw = True
    listEntries.Refresh
    Set vPORS = Nothing
    Set vPTSRS = Nothing
    'refresh rec sum
    RefreshRecInfo
    
    'refresh recopt buttons
    mdiMain.ActivateChild Me
    'restore mouse pointer
    mdiMain.Form_EndBussy
    
    End Function

Public Function Form_CanPrint() As Boolean
    Form_CanPrint = False
End Function

Public Function Form_Print()

End Function


Public Function Form_CanSearch() As Boolean
    If listEntries.RowCount > 0 Then
        Form_CanSearch = True
    End If
End Function

Public Function Form_SetSearch(ByRef sFields() As String)
    If listEntries.RowCount > 0 Then
    
        ReDim sFields(26)
          
        sFields(0) = "Date"
        sFields(1) = "DateID"
        sFields(2) = "POID"
        sFields(3) = "Ref. #"
        sFields(4) = "Supplier"
        sFields(5) = "CA"
        sFields(6) = "FP"
        sFields(7) = "Total Amount"
        sFields(8) = "Remarks"
      
        'PTS
        sFields(9) = "Payment ID"
        sFields(10) = "Date"
        'sFields(10) = "FK_SupID"
        sFields(11) = "Supplier"
        sFields(12) = "FP"
        sFields(13) = "Check No."
        sFields(14) = "Account Name"
        sFields(15) = "Date Issued"
        sFields(16) = "Date Due"
        sFields(17) = "Account No."
        sFields(18) = "Bank"
        sFields(19) = "Amount"
        sFields(20) = "Remarks"
        sFields(21) = "Cleared"
        sFields(22) = "RC"
        sFields(23) = "RM"
        sFields(24) = "RCU"
        sFields(26) = "RMU"

        
        Form_SetSearch = True

    End If
End Function
Public Function Form_Search(ByVal sSearchWhat As String, ByVal sField As String) As Boolean
    
    Dim li As Long
    Dim lx As Long
    Dim NewSelIndex As Long
    
    
    'default
    NewSelIndex = -1
    Form_Search = False
    
    listEntries.Redraw = False
    
    If listEntries.RowCount < 1 Then
        GoTo RAE
    End If
    
    If LCase(sField) = "all fields" Then
        
        'all fields
        For lx = 0 To listEntries.Cols - 1
            NewSelIndex = listEntries.FindItem(sSearchWhat, lx, lgWith, False)
            If NewSelIndex >= 0 Then
                listEntries.ItemSelected(NewSelIndex) = True
                listEntries.EnsureVisible NewSelIndex
                Exit For
            End If
        Next
        
    Else
    
        'by column
        For lx = 0 To listEntries.Cols - 1
            If LCase(sField) = LCase(listEntries.ColHeading(lx)) Then
                NewSelIndex = listEntries.FindItem(sSearchWhat, lx, lgWith, False)
                If NewSelIndex >= 0 Then
                    listEntries.ItemSelected(NewSelIndex) = True
                    listEntries.EnsureVisible NewSelIndex
                    
                    Exit For
                End If
                
                Exit For
            End If
        Next
    End If
    
RAE:
    If listEntries.SelectedCount > 1 Then
        For li = 0 To listEntries.RowCount - 1
            If NewSelIndex <> li Then
            listEntries.ItemSelected(li) = False
            End If
        Next
    End If
    
    listEntries.Redraw = True
    listEntries.Refresh
    
    'return
    If NewSelIndex >= 0 Then
        Form_Search = True
    End If
    
End Function


Public Sub Form_DateChange()
    Call Form_Refresh
End Sub




'----------------------------------------------------------
' Menu Procedures
'----------------------------------------------------------


Private Sub mnuAction_Click()
    mnuAdd.Enabled = Form_CanAdd
    mnuEdit.Enabled = Form_CanEdit
    mnuDelete.Enabled = Form_CanDelete
    mnuRefresh.Enabled = Form_CanRefresh
    mnuPrint.Enabled = Form_CanPrint
    mnuSearch.Enabled = Form_CanSearch
End Sub

Private Sub mnuAdd_Click()
    Form_Add
End Sub

Private Sub mnuAddPayment_Click()
    If frmPTSEntry.ShowAdd(mdiMain.b8DateP.MaxDate, CLng(GetTxtVal(b8DPSup.BoundData))) = True Then
        Form_Refresh
    End If
End Sub

Private Sub mnuDelete_Click()
    Form_Delete
End Sub

Private Sub mnuEdit_Click()
    Form_Edit
End Sub

Private Sub mnuPrint_Click()
    Form_Print
End Sub

Private Sub mnuRefresh_Click()
    Form_Refresh
End Sub

Private Sub mnuSearch_Click()
    
    mdiMain.Form_ShowSearch
End Sub

'----------------------------------------------------------
' >>> END Menu Procedures
'----------------------------------------------------------


