VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllCustPayDueCheck 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "Due Checks (Cust.)"
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
   Icon            =   "frmAllCustPayDueCheck.frx":0000
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
      Height          =   345
      Left            =   300
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   414
      TabIndex        =   3
      Top             =   4320
      Width           =   6210
      Begin VB.Label lblRecSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   60
         Width           =   585
      End
      Begin b8Controls4.b83DRect shpFooter 
         Height          =   285
         Index           =   1
         Left            =   0
         Top             =   30
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   503
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
      Left            =   450
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   2
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
               Picture         =   "frmAllCustPayDueCheck.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin b8Controls4.LynxGrid3 listEntries 
         Height          =   2355
         Left            =   510
         TabIndex        =   5
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
      Height          =   0
      Left            =   60
      ScaleHeight     =   0
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   530
      TabIndex        =   0
      Top             =   420
      Width           =   7950
   End
   Begin b8Controls4.b8TitleBar b8TB 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   661
      Caption         =   "Manage Due Checks (Customer Payments)"
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
Attribute VB_Name = "frmAllCustPayDueCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public dCurMinDate As Date
Public dCurMaxDate As Date

Private Const SIBgColor1 = &HF4FFFF
Private Const SIBgColor2 = &HE3F9FB

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
Private Sub b8DPProd_Change()
    Call Form_Refresh
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'pass keyinfo to mdiMain
    mdiMain.AFForm_KeyDown KeyCode, Shift
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
        
        .AddColumn "Pay. Date", 100
        .AddColumn "ID", 90
        .AddColumn "FK_SupID", 0
        .AddColumn "Supplier", 120
        .AddColumn "Check No.", 100
        .AddColumn "Account Name", 120
        .AddColumn "Date Issued", 100, lgAlignCenterCenter
        .AddColumn "Date Due", 100, lgAlignCenterCenter
        .AddColumn "Account No.", 100
        .AddColumn "Bank", 120
        .AddColumn "Amount", 100, lgAlignRightCenter
        .AddColumn "Remarks", 120
        .AddColumn "RC", 0
        .AddColumn "RM", 0
        .AddColumn "RCU", 0
        .AddColumn "RMU", 0
        .AddColumn "", 10 'Spacer
        .AddColumn "Number of days before Due", 120
                
        .RowHeightMin = 21
        .ImageList = ilList
        .Redraw = True
        .Refresh
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

'----------------------------------------------------------
' >>> END Form Procedures
'----------------------------------------------------------






'----------------------------------------------------------
' Record Procedures
'----------------------------------------------------------

Private Function FillDueCheck() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long
    Dim iDE As Integer
   
    'default
    FillDueCheck = False
   
     
    'get min and max date
    dCurMinDate = mdiMain.b8DateP.MinDate
    dCurMaxDate = mdiMain.b8DateP.MaxDate
    
    'clear list
    listEntries.Redraw = False
    listEntries.Clear
    
           
    sSQL = "SELECT tblCustPay.CustPayDate, tblCustPay.CustPayID, tblCustPay.FK_CustID, tblCust.CustName, tblCustPay.CheckNo, tblCustPay.AccountName, tblCustPay.DateIssued, tblCustPay.DateDue, tblCustPay.AccountNo, tblCustPay.BankName, tblCustPay.Amount, tblCustPay.Remarks, tblCustPay.RC, tblCustPay.RM, tblCustPay.RCU, tblCustPay.RMU" & _
            " FROM tblCust INNER JOIN tblCustPay ON tblCust.CustID = tblCustPay.FK_CustID" & _
            " Where LCase$([tblCustPay].[FP]) = 'check' And tblCustPay.Cleared = False AND DateValue(tblCustPay.DateDue)>=DateValue(#" & dCurMinDate & "#) AND DateValue(tblCustPay.DateDue)<=DateValue(#" & dCurMaxDate & "#)" & _
            " ORDER BY tblCustPay.CustPayDate"


             
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "FillDueCheck", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    vRS.MoveFirst
    While vRS.EOF = False
        
        With listEntries
            li = .AddItem(Format$(ReadField(vRS.Fields("CustPayDate")), modGV.GV_DateFormat))
            .ItemImage(li) = 1
            
            .CellText(li, 1) = modFunction.ComNumZ(ReadField(vRS.Fields("CustPayID")), 10)
            .CellText(li, 2) = modFunction.ComNumZ(ReadField(vRS.Fields("FK_CustID")), 10)
            .CellText(li, 3) = ReadField(vRS.Fields("CustName"))
            .CellText(li, 4) = ReadField(vRS.Fields("CheckNo"))
            .CellText(li, 5) = ReadField(vRS.Fields("AccountName"))
            .CellText(li, 6) = Format$(ReadField(vRS.Fields("DateIssued")), modGV.GV_DateFormat)
            .CellText(li, 7) = Format$(ReadField(vRS.Fields("DateDue")), modGV.GV_DateFormat)
            .CellText(li, 8) = ReadField(vRS.Fields("AccountNo"))
            .CellText(li, 9) = ReadField(vRS.Fields("BankName"))
            .CellText(li, 10) = FormatNumber(ReadField(vRS.Fields("Amount")), 2)
            .CellText(li, 11) = ReadField(vRS.Fields("Remarks"))
            .CellText(li, 12) = ReadField(vRS.Fields("RC"))
            .CellText(li, 13) = ReadField(vRS.Fields("RM"))
            .CellText(li, 14) = ReadField(vRS.Fields("RCU"))
            .CellText(li, 15) = ReadField(vRS.Fields("RMU"))
            
            iDE = DateValue(.CellText(li, 7)) - DateValue(Now)
            If iDE = 0 Then
                .CellText(li, 17) = "Today."
            ElseIf iDE > 0 Then
                .CellText(li, 17) = iDE & IIf(iDE > 1, " days", " day") & " more."
            Else
                .CellText(li, 17) = "Due."
            End If
            
        End With
        
        vRS.MoveNext
    Wend
    
    FillDueCheck = True

RAE:
    listEntries.Redraw = True
    listEntries.Refresh
    Set vRS = Nothing
End Function


Private Sub RefreshRecInfo()

    lblRecSum.Caption = "No selected"
    Me.mnuDelete.Caption = "Delete"
    Me.mnuEdit.Caption = "Edit"
    Me.mnuDelete.Enabled = False
    Me.mnuEdit.Enabled = False

    
    If listEntries.RowCount > 0 Then
        lblRecSum.Caption = "Selected Check entry Check No. " & listEntries.CellText(listEntries.Row, 4)
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

    If frmCustPayEntry.ShowAdd(mdiMain.b8DateP.MaxDate) = True Then
        Form_Refresh
    End If

End Function

Public Function Form_CanEdit() As Boolean
    If listEntries.RowCount > 0 Then
        Form_CanEdit = True
    End If
End Function

Public Function Form_Edit()
    
    Dim vID As Variant
    
    If listEntries.RowCount < 1 Then
        Exit Function
    End If

    'get ID
    vID = GetTxtVal(listEntries.CellText(listEntries.Row, 1))

    If frmCustPayEntry.ShowEdit(CLng(vID)) = True Then
        Form_Refresh
    End If
    
End Function

Public Function Form_CanDelete() As Boolean
    If listEntries.RowCount > 0 Then
        Form_CanDelete = True
    End If
End Function

Public Function Form_Delete()
    
    Dim vID As Variant
    Dim tmpCustPay As tCustPay
    
    If listEntries.RowCount < 1 Then
        Exit Function
    End If

    If MsgBox("Are you sure you want to delete this Payment entry with ID " & listEntries.CellText(listEntries.Row, 1) & "   ?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
        Exit Function
    End If
    
    'get ID
    vID = GetTxtVal(listEntries.CellText(listEntries.Row, 1))

    If GetCustPayByID(CLng(vID), tmpCustPay) = False Then
        Exit Function
    End If
    
    If DeleteCustPay(CLng(vID)) = True Then
        Form_Refresh
    Else
        WriteErrorLog Me.Name, "Form_Delete", "Failed on: 'DeleteCustPay(CLng(vID)) = True'"
    End If

End Function

Public Function Form_CanRefresh() As Boolean
    Form_CanRefresh = True
End Function

Public Function Form_Refresh()

    If bReadyToDisplay = False Then
        GoTo RAE
    End If
        
    'set app mouse icon
    mdiMain.Form_StartBussy

    lblRecSum.Caption = "Loading Data. Please wait..."
    DoEvents

    Call FillDueCheck

RAE:
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
    
        ReDim sFields(15)
        sFields(0) = "Pay. Date"
        sFields(1) = "ID"
        sFields(2) = "FK_SupID"
        sFields(3) = "SupName"
        sFields(4) = "Check No."
        sFields(5) = "Account Name"
        sFields(6) = "Date Issued"
        sFields(7) = "Date Due"
        sFields(8) = "Account No."
        sFields(9) = "Bank"
        sFields(10) = "Amount"
        sFields(11) = "Remarks"
        sFields(12) = "RC"
        sFields(13) = "RM"
        sFields(14) = "RCU"
        sFields(15) = "RMU"

        
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
    mnuEdit.Enabled = Form_CanEdit
    mnuDelete.Enabled = Form_CanDelete
    mnuRefresh.Enabled = Form_CanRefresh
    mnuPrint.Enabled = Form_CanPrint
    mnuSearch.Enabled = Form_CanSearch
End Sub


Private Sub mnuAddPayment_Click()
    Call Form_Add
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


