VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllSICPM 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Sales/Cust. Payments"
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
   Icon            =   "frmAllSICPM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox bgFooter 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   300
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   414
      TabIndex        =   5
      Top             =   4410
      Width           =   6210
      Begin VB.Label lblRecSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   585
      End
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
         TabIndex        =   6
         Top             =   150
         Width           =   525
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
   Begin VB.PictureBox bgCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   540
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   4
      Top             =   1050
      Width           =   5760
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
               Picture         =   "frmAllSICPM.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin b8Controls4.LynxGrid3 listEntries 
         Height          =   2355
         Left            =   510
         TabIndex        =   8
         Top             =   390
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
      Begin b8Controls4.b83DRect shpLBorder 
         Height          =   3015
         Left            =   0
         Top             =   0
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
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   530
      TabIndex        =   0
      Top             =   420
      Width           =   7950
      Begin b8Controls4.b8DataPicker b8DPCust 
         Height          =   360
         Left            =   90
         TabIndex        =   1
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
         Caption         =   "Customer:"
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   30
         Width           =   750
      End
      Begin b8Controls4.b83DRect shpFooter 
         Height          =   645
         Index           =   0
         Left            =   60
         Top             =   0
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1138
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
      TabIndex        =   3
      Top             =   0
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   661
      Caption         =   "Sales / Customer Payments Monitoring"
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
         Caption         =   "&New Sales Entry"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuAddPayment 
         Caption         =   "New Customer &Payment Entry"
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
Attribute VB_Name = "frmAllSICPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SIBgColor1 = &HF4FFFF
Private Const SIBgColor2 = &HDDFFEF   '&HF0D8C9

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
Private Sub b8DPCust_Change()
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
        'SI
        .AddColumn "Date", 110       '0
        .AddColumn "DateID", 0, , lgDate    '1
        .AddColumn "Sales Invoice ID", 0        '2
        .AddColumn "Ref. #", 80     '3
        .AddColumn "Customer", 120  '4
        .AddColumn "Total Amount", 90, lgAlignRightCenter '5
        .AddColumn "Remarks", 0     '6
        
        
        'spacer
        .AddColumn "", 10 '7
        
        'CustPay
        .AddColumn "CustPayID", 80     '8
        .AddColumn "Payment Date", 0     '9
        .AddColumn "Customer ID", 0     '10
        .AddColumn "Customer", 120     '11
        .AddColumn "FP", 60     '12
        .AddColumn "Check No.", 0     '13
        .AddColumn "Account Name", 0     '14
        .AddColumn "Date Issued", 0     '15
        .AddColumn "Date Due", 0     '16
        .AddColumn "Account No.", 0     '17
        .AddColumn "Bank", 0     '18
        .AddColumn "Amount", 90, lgAlignRightCenter     '19
        .AddColumn "Cleared", 60     '20
        .AddColumn "Remarks", 0     '21
        .AddColumn "RC", 0     '22
        .AddColumn "RM", 0     '23
        .AddColumn "RCU", 0     '24
        .AddColumn "RMU", 0     '25
   
        
        
        
        .RowHeightMin = 21
        .ImageList = ilList
        .Redraw = True
        .Refresh
    End With
    
    'set Customer list
    With b8DPCust
        Set .DropDBCon = PrimeDB
        .SQLFields = "String(10-Len(Trim([CustID])),'0') & [CustID] AS CCustID, tblCust.CustName"
        .SQLTable = "tblCust"
        .SQLWhereFields = "tblCust.CustID, tblCust.CustName"
        .SQLWhere = " tblCust.Active = True "
        .SQLOrderBy = "tblCust.CustName"
        
        .BoundFieldIndex = 0 'Bound Index
        .DisplayFieldIndex = 1
        .AddColumn "Customer ID", 100
        .AddColumn "Customer", 240
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

Private Sub LoadSI(ByRef vRS As ADODB.Recordset)
    
    Dim sSQL As String
   
    
    If IsEmpty(b8DPCust.BoundData) Then
        'without supplier
                
        sSQL = "SELECT tblSI.SIDate, tblSI.SIID, tblSI.RefNum, tblCust.CustName, tblSI.TotalAmt, tblSI.Remarks" & _
                " FROM tblCust INNER JOIN tblSI ON tblCust.CustID = tblSI.FK_CustID" & _
                " WHERE DateValue(tblSI.SIDate)>=DateValue(#" & DateValue(mdiMain.b8DateP.MinDate) & "#) AND DateValue(tblSI.SIDate)<=DateValue(#" & DateValue(mdiMain.b8DateP.MaxDate) & "#)" & _
                " ORDER BY tblSI.SIDate"
                
        
    Else
        'with supplier
        sSQL = "SELECT tblSI.SIDate, tblSI.SIID, tblSI.RefNum, tblCust.CustName, tblSI.TotalAmt, tblSI.Remarks" & _
                " FROM tblCust INNER JOIN tblSI ON tblCust.CustID = tblSI.FK_CustID" & _
                " WHERE tblSI.FK_CustID=" & CLng(b8DPCust.BoundData) & " AND DateValue(tblSI.SIDate)>=DateValue(#" & DateValue(mdiMain.b8DateP.MinDate) & "#) AND DateValue(tblSI.SIDate)<=DateValue(#" & DateValue(mdiMain.b8DateP.MaxDate) & "#)" & _
                " ORDER BY tblSI.SIDate"

    End If
             
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "LoadSI", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If

RAE:

End Sub

Private Sub LoadCustPay(ByRef vRS As ADODB.Recordset)

    Dim sSQL As String
    
    If IsEmpty(b8DPCust.BoundData) Then
        'without supplier
        
        
        sSQL = "SELECT tblCustPay.CustPayID, tblCustPay.FK_CustID, tblCust.CustName, tblCustPay.FP, tblCustPay.CustPayDate, tblCustPay.CheckNo, tblCustPay.AccountName, tblCustPay.DateIssued, tblCustPay.DateDue, tblCustPay.AccountNo, tblCustPay.BankName, tblCustPay.Amount, tblCustPay.Remarks, tblCustPay.Cleared, tblCustPay.RC, tblCustPay.RM, tblCustPay.RCU, tblCustPay.RMU" & _
                " FROM tblCust INNER JOIN tblCustPay ON tblCust.CustID = tblCustPay.FK_CustID" & _
                " WHERE DateValue(tblCustPay.CustPayDate)>=DateValue(#" & DateValue(mdiMain.b8DateP.MinDate) & "#) AND DateValue(tblCustPay.CustPayDate)<=DateValue(#" & DateValue(mdiMain.b8DateP.MaxDate) & "#)" & _
                " ORDER BY tblCustPay.CustPayDate"
    Else
    
        'with supplier
        sSQL = "SELECT tblCustPay.CustPayID, tblCustPay.FK_CustID, tblCust.CustName, tblCustPay.FP, tblCustPay.CustPayDate, tblCustPay.CheckNo, tblCustPay.AccountName, tblCustPay.DateIssued, tblCustPay.DateDue, tblCustPay.AccountNo, tblCustPay.BankName, tblCustPay.Amount, tblCustPay.Remarks, tblCustPay.Cleared, tblCustPay.RC, tblCustPay.RM, tblCustPay.RCU, tblCustPay.RMU" & _
                " FROM tblCust INNER JOIN tblCustPay ON tblCust.CustID = tblCustPay.FK_CustID" & _
                " WHERE tblCustPay.FK_CustID=" & CLng(b8DPCust.BoundData) & " AND DateValue(tblCustPay.CustPayDate)>=DateValue(#" & DateValue(mdiMain.b8DateP.MinDate) & "#) AND DateValue(tblCustPay.CustPayDate)<=DateValue(#" & DateValue(mdiMain.b8DateP.MaxDate) & "#)" & _
                " ORDER BY tblCustPay.CustPayDate"
                
                
    End If
             
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "LoadCustPay", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
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
                'SI
                If Not IsEmpty(listEntries.CellText(listEntries.Row, 2)) Then
                    lblRecSum.Caption = "Selected Sales Invoice entry Ref. # " & listEntries.CellText(listEntries.Row, 3)
                    Me.mnuDelete.Caption = "Delete Sales Invoice entry"
                    Me.mnuEdit.Caption = "Edit Sales Invoice entry"
                    Me.mnuDelete.Enabled = True
                    Me.mnuEdit.Enabled = True
                End If
            ElseIf listEntries.Col >= 10 And listEntries.Col <= 25 Then
                'CustPay
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
    
    If frmSIEntry.ShowAdd(mdiMain.b8DateP.MaxDate, CLng(GetTxtVal(b8DPCust.BoundData))) = True Then
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
        If listEntries.Col >= 0 And listEntries.Col <= 6 Then
            'SI
            If Not IsEmpty(listEntries.CellText(listEntries.Row, 2)) Then
                
                vID = GetTxtVal(listEntries.CellText(listEntries.Row, 2))
                If frmSIEntry.ShowEdit(CLng(vID)) = True Then
                    Form_Refresh
                End If
                
            End If
        ElseIf listEntries.Col >= 8 And listEntries.Col <= 25 Then
            'CustPay
            If Not IsEmpty(listEntries.CellText(listEntries.Row, 8)) Then
                vID = GetTxtVal(listEntries.CellText(listEntries.Row, 8))
                If frmCustPayEntry.ShowEdit(CLng(vID)) = True Then
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
    Dim vSI As tSI
    Dim tmpCustPay As tCustPay
    
    If listEntries.RowCount < 1 Then
        Exit Function
    End If
    
    If IsEmpty(listEntries.CellText(listEntries.Row, 0)) Then
        If listEntries.Col >= 0 And listEntries.Col <= 8 Then
            'SI
            If Not IsEmpty(listEntries.CellText(listEntries.Row, 2)) Then
                
                If MsgBox("Are you sure you want to delete this Sales Invoice entry with Ref.# " & listEntries.CellText(listEntries.Row, 3) & " ?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                    Exit Function
                End If
                'get ID
                vID = listEntries.CellText(listEntries.Row, 2)
            
                If GetSIByID(CLng(vID), vSI) = False Then
                    MsgBox "SI entry notfound.", vbExclamation
                    Exit Function
                End If
                
                If DeleteSI(vSI.SIID) = True Then
                    If GetCustPayByID(vSI.OptFK_CustPayID, tmpCustPay) = True Then
                        If MsgBox("Do you want to delete Payment entry that associated with this Sales Invoice entry?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                            'delete associated CustPay
                            If modRSCustPay.DeleteCustPay(vSI.OptFK_CustPayID) = False Then
                                WriteErrorLog Me.Name, "Form_Delete", "Failed on: 'modRSCustPay.DeleteCustPay(vSI.OptFK_CustPayID) = False'"
                            End If
                        End If
                    End If
                    
                    Form_Refresh
                Else
                    WriteErrorLog Me.Name, "Form_Delete", "Failed on: 'DeleteSI(CLng(vID)) = True'"
                End If
    
            End If
        ElseIf listEntries.Col >= 10 And listEntries.Col <= 25 Then
            'CustPay
            If Not IsEmpty(listEntries.CellText(listEntries.Row, 8)) Then
                If MsgBox("Are you sure you want to delete this Payment entry with ID " & listEntries.CellText(listEntries.Row, 10) & " ?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                    Exit Function
                End If
                'get ID
                vID = listEntries.CellText(listEntries.Row, 8)
            
                If DeleteCustPay(CLng(vID)) = True Then
                    Form_Refresh
                Else
                    WriteErrorLog Me.Name, "Form_Delete", "Failed on: 'DeleteCustPay(CLng(vID)) = True'"
                End If
            End If
        End If
        
    Else
        MsgBox "Please select 'Sales Invoice' entry or 'Payment' entry.", vbExclamation
    End If

End Function

Public Function Form_CanRefresh() As Boolean
    Form_CanRefresh = True
End Function

Public Function Form_Refresh()

    Dim vSIRS As New ADODB.Recordset
    Dim vCustPayRS As New ADODB.Recordset
    
    Dim li As Long
    Dim lhi As Long
    Dim dD As Date
    Dim bAdded As Boolean
    
    Dim dSIDate As Date
    Dim dCustPayDate As Date
    Dim dOldDate As Date
    
    Dim FP As String
        
    
    
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
    Call LoadSI(vSIRS)
    Call LoadCustPay(vCustPayRS)
    
    
    
    
    
    If AnyRecordExisted(vSIRS) = False And AnyRecordExisted(vCustPayRS) = False Then
        GoTo RAE
    End If
    
    If AnyRecordExisted(vSIRS) = True Then
        vSIRS.MoveFirst
    End If
    If AnyRecordExisted(vCustPayRS) = True Then
        vCustPayRS.MoveFirst
    End If
    
    dD = mdiMain.b8DateP.MinDate
 
    
    While vSIRS.EOF = False Or vCustPayRS.EOF = False

        bAdded = False

        'add SI
        If vSIRS.EOF = False Then
            dSIDate = ReadField(vSIRS.Fields("SIDate"))
            If DateValue(dD) = DateValue(dSIDate) Then
            
                'add Header Date
                If DateValue(dOldDate) <> DateValue(dD) Then
                    lhi = listEntries.AddItem(Format(dD, modGV.GV_DateFormat))
                    listEntries.FillBackColor lhi, lhi, 0, 25, &HF5F5F5
                End If
        
                'same, then add
                With listEntries
                li = .AddItem("")
                .FillBackColor li, li, 0, 6, SIBgColor2
                .ItemImage(li) = 1
                .CellText(li, 1) = ReadField(vSIRS.Fields("SIDate"))
                .CellText(li, 2) = ReadField(vSIRS.Fields("SIID"))
                .CellText(li, 3) = modFunction.ComNumZ(ReadField(vSIRS.Fields("RefNum")), 10)
                .CellText(li, 4) = ReadField(vSIRS.Fields("CustName"))
                .CellFontBold(li, 5) = True
                .CellText(li, 5) = FormatNumber(ReadField(vSIRS.Fields("TotalAmt")), 2)
                .CellText(li, 6) = ReadField(vSIRS.Fields("Remarks"))
                End With
                
                vSIRS.MoveNext
                
                'set flag
                bAdded = True
            End If
        End If
        
        'add CustPay
        If vCustPayRS.EOF = False Then
            dCustPayDate = ReadField(vCustPayRS.Fields("CustPayDate"))
            
            If DateValue(dD) = DateValue(dCustPayDate) Then
            
                'same, then add
                With listEntries
                If bAdded = False Then
                    'add Header Date
                    If DateValue(dOldDate) <> DateValue(dD) Then
                        lhi = listEntries.AddItem(Format(dD, modGV.GV_DateFormat))
                        listEntries.FillBackColor lhi, lhi, 0, 25, &HA9EEFF
                    End If
                    
                    li = .AddItem(" ")
                End If
                
                .FillBackColor li, li, 8, 25, SIBgColor1
                .ItemImage(li) = 1
                .CellText(li, 8) = modFunction.ComNumZ(ReadField(vCustPayRS.Fields("CustPayID")), 10)
                .CellText(li, 9) = dCustPayDate
                .CellText(li, 10) = ReadField(vCustPayRS.Fields("FK_CustID"))
                .CellText(li, 11) = ReadField(vCustPayRS.Fields("CustName"))
                 FP = LCase(ReadField(vCustPayRS.Fields("FP")))
                .CellText(li, 12) = FP
                .CellText(li, 13) = ReadField(vCustPayRS.Fields("CheckNo"))
                .CellText(li, 14) = ReadField(vCustPayRS.Fields("AccountName"))
                .CellText(li, 15) = ReadField(vCustPayRS.Fields("DateIssued"))
                .CellText(li, 16) = ReadField(vCustPayRS.Fields("DateDue"))
                .CellText(li, 17) = ReadField(vCustPayRS.Fields("AccountNo"))
                .CellText(li, 18) = ReadField(vCustPayRS.Fields("BankName"))
                .CellFontBold(li, 19) = True
                .CellText(li, 19) = FormatNumber(ReadField(vCustPayRS.Fields("Amount")), 2)
                .CellText(li, 20) = IIf(FP = "check", IIf(ReadField(vCustPayRS.Fields("Cleared")), "Yes", "No"), "N/A")
                .CellText(li, 21) = ReadField(vCustPayRS.Fields("Remarks"))
                .CellText(li, 22) = ReadField(vCustPayRS.Fields("RC"))
                .CellText(li, 23) = ReadField(vCustPayRS.Fields("RM"))
                .CellText(li, 24) = ReadField(vCustPayRS.Fields("RCU"))
                .CellText(li, 25) = ReadField(vCustPayRS.Fields("RMU"))

                End With
                
                vCustPayRS.MoveNext
                
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
    
    lblRecSum.Caption = "Calculating A/R..."
    DoEvents
    'set fotter info
    If Not IsEmpty(b8DPCust.BoundData) Then
            lblInfo.Caption = "A/R (Beg): " & _
                            FormatNumber(modRSAR.GetARByCust(CLng(b8DPCust.BoundData), CDate(0), mdiMain.b8DateP.MinDate - 1), 2) & _
                            "   |   A/R (End): " & _
                            FormatNumber(modRSAR.GetARByCust(CLng(b8DPCust.BoundData), CDate(0), mdiMain.b8DateP.MaxDate), 2)
    Else
        
        lblInfo.Caption = "For all Customer,  A/R (Beg): " & _
                            FormatNumber(modRSAR.GetAllAR(CDate(0), mdiMain.b8DateP.MinDate - 1), 2) & _
                            "   |   A/R (End): " & _
                            FormatNumber(modRSAR.GetAllAR(CDate(0), mdiMain.b8DateP.MaxDate), 2)
    End If
    lblRecSum.Caption = "Ready"

    listEntries.Redraw = True
    listEntries.Refresh
    Set vSIRS = Nothing
    Set vCustPayRS = Nothing
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
        sFields(2) = "Sales Invoice ID"
        sFields(3) = "Ref. #"
        sFields(4) = "Customer"
        sFields(5) = "CA"
        sFields(6) = "FP"
        sFields(7) = "Total Amount"
        sFields(8) = "Remarks"

      
        'CustPay
        sFields(9) = "CustPayID"
        sFields(10) = "Payment Date"
        'sFields(10) = "FK_CustID"
        sFields(11) = "Customer"
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
    If frmCustPayEntry.ShowAdd(mdiMain.b8DateP.MaxDate, CLng(GetTxtVal(b8DPCust.BoundData))) = True Then
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


