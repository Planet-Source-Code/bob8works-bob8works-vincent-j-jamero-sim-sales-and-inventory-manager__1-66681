VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllStockInv 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Stock Mon."
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
   Icon            =   "frmAllStockInv.frx":0000
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
      Height          =   375
      Left            =   570
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   414
      TabIndex        =   5
      Top             =   4500
      Width           =   6210
      Begin VB.Label lblRecSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   90
         Width           =   585
      End
      Begin b8Controls4.b83DRect shpFooter 
         Height          =   315
         Index           =   0
         Left            =   0
         Top             =   30
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
      Height          =   3765
      Left            =   540
      ScaleHeight     =   251
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
               Picture         =   "frmAllStockInv.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin b8Controls4.LynxGrid3 listEntries 
         Height          =   2355
         Left            =   510
         TabIndex        =   7
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
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   530
      TabIndex        =   0
      Top             =   390
      Width           =   7950
      Begin b8Controls4.b8DataPicker b8DPProd 
         Height          =   360
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Width           =   6525
         _ExtentX        =   11509
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
         DropWinWidth    =   9735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product:"
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   60
         Width           =   615
      End
      Begin b8Controls4.b83DRect shpFooter 
         Height          =   645
         Index           =   2
         Left            =   0
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
      Caption         =   "Stock Monitoring"
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
Attribute VB_Name = "frmAllStockInv"
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


Private Sub listEntries_DblClick()
    Form_Edit
End Sub

Private Sub listEntries_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
        .AddColumn "Category", 100          '0
        .AddColumn "CatID", 0          '1
        .AddColumn "ProdID", 0          '2
        .AddColumn "Code", 90, lgAlignCenterCenter          '3
        .AddColumn "Description", 140          '4
        .AddColumn "PackID", 0          '5
        .AddColumn "PackTitle", 100          '6
        .AddColumn "BegInvStock", 0          '7
        
        .AddColumn " ", 10 'Spacer          '8
        
        .AddColumn "Beg. Inv.", 100, lgAlignRightCenter           '9
        
        .AddColumn " ", 10 'Spacer          '10
        
        'In
        .AddColumn "Purchsed", 90, lgAlignRightCenter          '11
        .AddColumn "Other", 0, lgAlignRightCenter     'temp 101          '12
        
        .AddColumn " ", 2 'Spacer          '13
        'out
        .AddColumn "Sold", 90, lgAlignRightCenter              '14
        .AddColumn "Void", 90, lgAlignRightCenter           '15
        .AddColumn "Other", 0 'temp 101          '16
        
        .AddColumn " ", 10 'Spacer          '17
        
        .AddColumn "End. Inv.", 100, lgAlignRightCenter '11          '18
        
        .AddColumn " ", 10 'Spacer          '19
        
        .AddColumn "SupPrice", 90, lgAlignRightCenter          '20
        .AddColumn "SRPrice", 90, lgAlignRightCenter          '21
        
        .AddColumn " ", 10 'Spacer          '22
        .AddColumn "Est. Sup. Amt.", 90, lgAlignRightCenter          '23
        .AddColumn "Est. SRP Amt.", 90, lgAlignRightCenter          '24
        
        'hiden
        .AddColumn "Info"
        .AddColumn "RecInfo"
        
        
        
        .RowHeightMin = 21
        .ImageList = ilList
        .Redraw = True
        .Refresh
    End With
    
   With b8DPProd
        Set .DropDBCon = PrimeDB
        .SQLFields = "String(10-Len(trim(tblProd.ProdID)),'0') & tblProd.ProdID as CProdID, tblProd.ProdCode, tblProd.ProdDescription, tblPack.PackTitle, tblCat.CatTitle, Format$([SupPrice],'Fixed') as SP, Format$([SRPrice],'Fixed') as SRP" ' tblProd.SupPrice, tblProd.SRPrice"
        .SQLTable = "tblPack INNER JOIN (tblCat INNER JOIN tblProd ON tblCat.CatID = tblProd.FK_CatID) ON tblPack.PackID = tblProd.FK_PackID"
        .SQLWhere = "tblProd.Active=True"
        .SQLWhereFields = "tblProd.ProdID, tblProd.ProdCode, tblProd.ProdDescription, tblPack.PackTitle, tblCat.CatTitle, tblProd.SupPrice, tblProd.SRPrice"
        .SQLOrderBy = "tblProd.ProdDescription"
        
        .BoundFieldIndex = 0 'Bound Index
        .DisplayFieldIndex = 2
        
        .AddColumn "ID", 100
        .AddColumn "Code", 100
        .AddColumn "Description", 180
        .AddColumn "Unit", 70
        .AddColumn "Category", 80
        .AddColumn "Sup. Price", 60, lgAlignRightCenter
        .AddColumn "SRP", 60, lgAlignRightCenter
        
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

Private Function FillProducts(ByRef lRecCount As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long
   
   'default
   FillProducts = False
    
    If IsEmpty(b8DPProd.BoundData) Then
        'without Product
        sSQL = "SELECT tblCat.CatID, tblCat.CatTitle, tblProd.ProdID, tblProd.ProdCode, tblProd.ProdDescription, tblPack.PackID, tblPack.PackTitle, tblProd.BegInvStock, tblProd.SupPrice, tblProd.SRPrice" & _
                " FROM tblPack INNER JOIN (tblCat INNER JOIN tblProd ON tblCat.CatID = tblProd.FK_CatID) ON tblPack.PackID = tblProd.FK_PackID" & _
                " Where tblProd.Active = Yes" & _
                " ORDER BY tblCat.CatTitle, tblProd.ProdCode"

        
    Else
        'with Product
        sSQL = "SELECT tblCat.CatID, tblCat.CatTitle, tblProd.ProdID, tblProd.ProdCode, tblProd.ProdDescription, tblPack.PackID, tblPack.PackTitle, tblProd.BegInvStock, tblProd.SupPrice, tblProd.SRPrice" & _
                " FROM tblPack INNER JOIN (tblCat INNER JOIN tblProd ON tblCat.CatID = tblProd.FK_CatID) ON tblPack.PackID = tblProd.FK_PackID" & _
                " Where tblProd.ProdID=" & CLng(b8DPProd.BoundData) & " AND tblProd.Active = Yes" & _
                " ORDER BY tblCat.CatTitle, tblProd.ProdCode"
    End If
             
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "FillProducts", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    listEntries.Redraw = False
    listEntries.Clear
    
    vRS.MoveFirst
    While vRS.EOF = False
        
        With listEntries
            li = .AddItem(ReadField(vRS.Fields("CatTitle")))
            .ItemImage(li) = 1
            .CellText(li, 1) = ReadField(vRS.Fields("CatID")) '1
            .CellText(li, 2) = ReadField(vRS.Fields("ProdID")) ' 0
            .CellText(li, 3) = ReadField(vRS.Fields("ProdCode")) ' 90
            .CellText(li, 4) = ReadField(vRS.Fields("ProdDescription")) ' 120
            .CellText(li, 5) = ReadField(vRS.Fields("PackID")) ' 0
            .CellText(li, 6) = ReadField(vRS.Fields("PackTitle")) ' 0
            .CellText(li, 7) = ReadField(vRS.Fields("BegInvStock")) ' 0
            
            .CellText(li, 20) = FormatNumber(ReadField(vRS.Fields("SupPrice")), 2)
            .CellText(li, 21) = FormatNumber(ReadField(vRS.Fields("SRPrice")), 2)
        
        End With
        
        vRS.MoveNext
    Wend
    
    lRecCount = modDBMain.getRecordCount(vRS)
    
    FillProducts = True

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
        lblRecSum.Caption = "Selected Product: " & listEntries.CellText(listEntries.Row, 4) & _
                            "        Inventory End: " & listEntries.CellText(listEntries.Row, 18) & " " & listEntries.CellText(listEntries.Row, 6)
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


End Function
Public Function Form_Add()
    
End Function

Public Function Form_CanEdit() As Boolean

End Function

Public Function Form_Edit()
    
End Function

Public Function Form_CanDelete() As Boolean

End Function

Public Function Form_Delete()
    
End Function

Public Function Form_CanRefresh() As Boolean
    Form_CanRefresh = True
End Function

Public Function Form_Refresh()
   
    Dim li As Long
    Dim sCat As String
    Dim lBgColor As Long
    
    'beg inv
    Dim dTBegInv As Double
    'purchased
    Dim dTPurchased As Double
    Dim dTSold As Double
    'end inv
    Dim dTEndInv As Double
    Dim dTSupAmt As Double
    Dim dTSRP As Double
    
    'Product ID
    Dim lPID As Long
    'record Count
    Dim lTotalRecCount As Long
    
    
       
    If bReadyToDisplay = False Then
        GoTo RAE
    End If
        
    'set app mouse icon
    mdiMain.Form_StartBussy
    
    'get min and max date
    dCurMinDate = mdiMain.b8DateP.MinDate
    dCurMaxDate = mdiMain.b8DateP.MaxDate

    
    

    lblRecSum.Caption = "Loading Data. Please wait..."
    DoEvents
    
    '---------------------------------------------------
    'load data
    If FillProducts(lTotalRecCount) = False Then
        GoTo RAE
    End If
    
     
    
    listEntries.Redraw = False
    For li = 0 To listEntries.RowCount - 1
    
        lPID = CLng(listEntries.CellText(li, 2))
        'beg inv
        listEntries.CellText(li, 9) = FormatNumber(modRSStockInv.GetProdStock(lPID, dCurMinDate - 1), 0)
        'purchased
        listEntries.CellText(li, 11) = FormatNumber(modRSStockInv.GetSumPOInvQtyByDate(lPID, dCurMinDate, dCurMaxDate), 0)
        'sold
        listEntries.CellText(li, 14) = FormatNumber(modRSStockInv.GetSumSIInvQtyByDate(lPID, dCurMinDate, dCurMaxDate), 0)
        'void
        listEntries.CellText(li, 15) = FormatNumber(modRSStockInv.GetSumVoidInvQtyByDate(lPID, dCurMinDate, dCurMaxDate), 0)
        'end inv
        listEntries.CellText(li, 18) = FormatNumber((GetTxtVal(listEntries.CellText(li, 9)) + GetTxtVal(listEntries.CellText(li, 11)) + GetTxtVal(listEntries.CellText(li, 12))) _
                                        - (GetTxtVal(listEntries.CellText(li, 14)) + GetTxtVal(listEntries.CellText(li, 15)) + GetTxtVal(listEntries.CellText(li, 16))), 0)
    
        listEntries.CellText(li, 23) = FormatNumber(GetTxtVal(listEntries.CellText(li, 18)) * GetTxtVal(listEntries.CellText(li, 20)), 2)
        listEntries.CellText(li, 24) = FormatNumber(GetTxtVal(listEntries.CellText(li, 18)) * GetTxtVal(listEntries.CellText(li, 21)), 2)
    
        'add total
        dTBegInv = dTBegInv + GetTxtVal(listEntries.CellText(li, 9))
        dTPurchased = dTPurchased + GetTxtVal(listEntries.CellText(li, 11))
        dTSold = dTSold + GetTxtVal(listEntries.CellText(li, 14))
        dTEndInv = dTEndInv + GetTxtVal(listEntries.CellText(li, 18))
        dTSupAmt = dTSupAmt + GetTxtVal(listEntries.CellText(li, 23))
        dTSRP = dTSRP + GetTxtVal(listEntries.CellText(li, 24))
        
    
    
        If sCat <> listEntries.CellText(li, 0) Then
            'change bgcolor
            If lBgColor = SIBgColor1 Then
                lBgColor = SIBgColor2
            Else
                lBgColor = SIBgColor1
            End If
        End If
        listEntries.ItemBackColor(li) = lBgColor
        listEntries.CellFontBold(li, 4) = True
        listEntries.CellFontBold(li, 9) = True
        listEntries.CellFontBold(li, 18) = True
        sCat = listEntries.CellText(li, 0)
                
        
        'doevents
        lblRecSum.Caption = "Loading Products Info. [ " & li & " / " & lTotalRecCount & " ] "
        DoEvents
        
    Next
    
    
    'add total
    listEntries.AddItem ""
    li = listEntries.AddItem("Total")
    listEntries.CellText(li, 9) = FormatNumber(dTBegInv, 0)
    listEntries.CellText(li, 11) = FormatNumber(dTPurchased, 0)
    listEntries.CellText(li, 14) = FormatNumber(dTSold, 0)
    listEntries.CellText(li, 18) = FormatNumber(dTEndInv, 0)
    listEntries.CellText(li, 23) = FormatNumber(dTSupAmt, 2)
    listEntries.CellText(li, 24) = FormatNumber(dTSRP, 2)
    'apply formatting
    listEntries.ItemBackColor(li) = &H80&
    listEntries.ItemForeColor(li) = vbWhite
    listEntries.ItemFontBold(li) = True
    '---------------------------------------------------
    

RAE:
    
    listEntries.Redraw = True
    listEntries.Refresh

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
        sFields(2) = "SIID"
        sFields(3) = "Ref. #"
        sFields(4) = "Product"
        sFields(5) = "CA"
        sFields(6) = "FP"
        sFields(7) = "Total Amount"
        sFields(8) = "Remarks"
      
        'CustPay
        sFields(9) = "CustPayID"
        sFields(10) = "CustPayDate"
        'sFields(10) = "FK_CustID"
        sFields(11) = "Product"
        sFields(12) = "FP"
        sFields(13) = "CheckNo"
        sFields(14) = "AccountName"
        sFields(15) = "DateIssued"
        sFields(16) = "DateDue"
        sFields(17) = "AccountNo"
        sFields(18) = "BankName"
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
    'mnuAdd.Enabled = Form_CanAdd
    'mnuEdit.Enabled = Form_CanEdit
    'mnuDelete.Enabled = Form_CanDelete
    mnuRefresh.Enabled = Form_CanRefresh
    mnuPrint.Enabled = Form_CanPrint
    mnuSearch.Enabled = Form_CanSearch
End Sub

Private Sub mnuAdd_Click()
    Form_Add
End Sub

Private Sub mnuAddPayment_Click()
    If frmCustPayEntry.ShowAdd(mdiMain.b8DateP.MaxDate, CLng(GetTxtVal(b8DPProd.BoundData))) = True Then
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


