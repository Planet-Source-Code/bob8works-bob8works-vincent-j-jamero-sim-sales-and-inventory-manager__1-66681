VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVoidEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Void Entry"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVoidEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3000
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   3000
      Width           =   1395
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   2985
      Left            =   0
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   467
      TabIndex        =   5
      Top             =   600
      Width           =   7005
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1410
         MaxLength       =   100
         TabIndex        =   9
         ToolTipText     =   "Customer Name"
         Top             =   1590
         Width           =   2625
      End
      Begin VB.CommandButton Command1 
         Height          =   345
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   750
         Width           =   375
      End
      Begin VB.ComboBox cmbPackTitle 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1170
         Width           =   2625
      End
      Begin VB.TextBox txtVoidID 
         BackColor       =   &H00F5F5F5&
         Height          =   285
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   90
         Width           =   1635
      End
      Begin b8Controls4.b8DataPicker b8DPProd 
         Height          =   360
         Left            =   1410
         TabIndex        =   10
         Top             =   750
         Width           =   4845
         _ExtentX        =   8546
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
      Begin b8Controls4.b8Line b8Line1 
         Height          =   30
         Left            =   0
         TabIndex        =   11
         Top             =   420
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   30
         TabIndex        =   12
         Top             =   0
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line3 
         Height          =   30
         Left            =   0
         TabIndex        =   13
         Top             =   2280
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin MSComCtl2.DTPicker dtpVoidDate 
         Height          =   315
         Left            =   5070
         TabIndex        =   14
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM - dd - yyyy"
         Format          =   183042051
         CurrentDate     =   38961
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Quantity:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label lblRC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1320
         TabIndex        =   20
         Top             =   3870
         Width           =   45
      End
      Begin VB.Label lblRM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1320
         TabIndex        =   19
         Top             =   4050
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* &Product:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   810
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Unit:"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         Height          =   195
         Left            =   2250
         TabIndex        =   16
         Top             =   90
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Date:"
         Height          =   195
         Left            =   4590
         TabIndex        =   15
         Top             =   90
         Width           =   405
      End
   End
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   0
      Width           =   6975
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmVoidEntry.frx":000C
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VOID PRODUCT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   345
         Left            =   600
         TabIndex        =   4
         Top             =   60
         Width           =   2265
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fill all fields or fields with '*' then click 'Save' button to update."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   180
         Left            =   630
         TabIndex        =   3
         Top             =   420
         Width           =   3900
      End
   End
End
Attribute VB_Name = "frmVoidEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mFormState As String

Dim curVoid As tVoid
Dim newVoid As tVoid

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isOn As Boolean

Dim ProdPackList() As tProdPack

Public Function ShowAdd(Optional ByVal lProdID As Long = -1, Optional dVoidDate As Date = 0, Optional dQty As Double) As Boolean
    
    'set form state
    mFormState = "add"
    
    'set parameter
    newVoid.FK_ProdID = lProdID
    newVoid.Qty = dQty
    If dVoidDate = 0 Then
        dVoidDate = Now
    End If
    newVoid.VoidDate = dVoidDate
    
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function

Public Function ShowEdit(ByVal lVoidID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curVoid.VoidID = lVoidID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function




Private Sub b8DPProd_Change()
    If RefeshCurProd(CLng(GetTxtVal(b8DPProd.BoundData))) = False Then
        Exit Sub
    End If
End Sub

Private Sub cmdSave_Click()

    Dim tmpProd As tProd
    
    
    'Add/Edit validations
    If GetProdByID(CLng(GetTxtVal(b8DPProd.BoundData)), tmpProd) = False Then
        MsgBox "Select enter valid Product.", vbExclamation
        b8DPProd.FocusedDropButton
        Exit Sub
    End If
    
    If Not (GetTxtVal(txtQty.Text) > 0) Then
        MsgBox "Please enter valid Quantity.", vbExclamation
        HLTxt txtQty
        Exit Sub
    End If
    
    If cmbPackTitle.ListIndex < 0 Then
        MsgBox "Please select valid Package.", vbExclamation
        cmbPackTitle.SetFocus
        Exit Sub
    End If
    

    Select Case mFormState
        Case "add"
            SaveAdd
        Case "edit"
            SaveEdit
    End Select
End Sub

Private Sub cmdCancel_Click()
    
    Select Case mFormState
        Case "add"
            mShowAdd = False
        Case "edit"
            mShowEdit = False
    End Select
    
    Unload Me
End Sub





Private Sub Form_Activate()
    
    Dim tmpProd As tProd
    
    
    'make mouse pointer bussy
    Me.MousePointer = vbHourglass
 
    Select Case mFormState
        Case "add"
                        
            'set form caption
            Me.Caption = "Add New Void Entry"
            
            'generate void ID
            txtVoidID.Text = modFunction.ComNumZ(modRSVoid.GetNewVoidID, 10)
            
            'set date
            dtpVoidDate.Value = newVoid.VoidDate
            
            
            If newVoid.FK_ProdID > 0 Then
                If GetProdByID(newVoid.FK_ProdID, tmpProd) = True Then
                    b8DPProd.BoundData = newVoid.FK_ProdID
                    b8DPProd.DisplayData = tmpProd.ProdDescription
                    Call b8DPProd_Change
                End If
            End If
            
            txtQty.Text = newVoid.Qty

            
        Case "edit"
        
            'set form caption
            Me.Caption = "Edit Void Entry"
            
            'generate void ID
            txtVoidID.Text = modFunction.ComNumZ(curVoid.VoidID, 10)

            If GetVoidByID(curVoid.VoidID, curVoid) = True Then
                
                If GetProdByID(curVoid.FK_ProdID, tmpProd) = True Then
                    b8DPProd.BoundData = curVoid.FK_ProdID
                    b8DPProd.DisplayData = tmpProd.ProdDescription
                    Call b8DPProd_Change
                End If
            End If
           
            dtpVoidDate.Value = curVoid.VoidDate
            txtQty.Text = curVoid.Qty
            
            
            'disable some controls
            dtpVoidDate.Enabled = False
            b8DPProd.Enabled = False
            
    End Select
    
    
RAE:
    'restoremouse pointer tonormal
    Me.MousePointer = vbNormal
    Me.AutoRedraw = True
End Sub


Private Sub Form_Load()

    isOn = False
    PaintGrad bgMain, &HF8F8F8, &HFFFFFF, 90
    PaintGrad bgHeader, &HEDEBE9, &HFFFFFF, 0
  
    
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






Private Sub SaveAdd()
    
    'set new void info
    
    With newVoid
        .VoidID = CLng(GetTxtVal(txtVoidID.Text))
        .FK_ProdID = CLng(GetTxtVal(b8DPProd.BoundData))
        .FK_PackID = CLng(ProdPackList(cmbPackTitle.ListIndex).FK_PackID)
        .Qty = GetTxtVal(txtQty.Text)
        .InvQty = .Qty * ProdPackList(cmbPackTitle.ListIndex).Qty
        .VoidDate = modFunction.GetRSec(dtpVoidDate.Value)
        
    End With
    
    
    'write
    If modRSVoid.AddVoid(newVoid) = True Then
    
        'set flag
        mShowAdd = True
        'close this form
        Unload Me
        
    Else
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSVoid.AddVoid(newVoid) = True'"
    End If
    
End Sub


Private Sub SaveEdit()
    
    With curVoid
        .VoidID = CLng(GetTxtVal(txtVoidID.Text))
        .FK_ProdID = CLng(GetTxtVal(b8DPProd.BoundData))
        .FK_PackID = CLng(ProdPackList(cmbPackTitle.ListIndex).FK_PackID)
        .Qty = GetTxtVal(txtQty.Text)
        .InvQty = .Qty * ProdPackList(cmbPackTitle.ListIndex).Qty
        .VoidDate = modFunction.GetRSec(dtpVoidDate.Value)
        
    End With
    
    
    'write
    If modRSVoid.EditVoid(curVoid) = True Then
    
        'set flag
        mShowEdit = True
        'close this form
        Unload Me
        
    Else
        WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'modRSVoid.EditVoid(curVoid) = True'"
    End If
    
End Sub




Private Function RefeshCurProd(ByVal lProdID As Long, Optional lPackID As Long) As Boolean
    
    Dim i As Integer
    Dim vProd As tProd

    'default
    RefeshCurProd = False
    
    If GetProdByID(lProdID, vProd) = False Then
        Exit Function
    End If
    
    'fill packages
    If modRSProdPack.FillProdPackToTypeArray(lProdID, ProdPackList) = False Then
        WriteErrorLog Me.Name, "RefeshCurProd", "Failed on: 'modRSProdPack.FillProdPackToTypeArray(lProdID, prodpacklis) = False'"
        Exit Function
    End If
    
    If UBound(ProdPackList) >= 0 Then
        For i = 0 To UBound(ProdPackList)
            cmbPackTitle.AddItem ProdPackList(i).PackTitle
        Next
    Else
        Exit Function
    End If
    
    cmbPackTitle.Enabled = True
    'default package
    cmbPackTitle.ListIndex = 0
    'set current package base on parameter
    For i = 0 To UBound(ProdPackList)
        If ProdPackList(i).FK_PackID = lPackID Then
            cmbPackTitle.ListIndex = i
            Exit For
        End If
    Next
        
    'return sucess
    RefeshCurProd = True
    
End Function

