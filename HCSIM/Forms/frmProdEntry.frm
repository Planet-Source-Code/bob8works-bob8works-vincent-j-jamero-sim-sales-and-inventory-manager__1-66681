VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProdEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Entry"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProdEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   398
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   737
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
      Left            =   8010
      TabIndex        =   10
      Top             =   5490
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   9570
      TabIndex        =   11
      Top             =   5490
      Width           =   1395
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   5385
      Left            =   0
      ScaleHeight     =   359
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   737
      TabIndex        =   12
      Top             =   600
      Width           =   11055
      Begin b8Controls4.b8GradLine b8GradLine1 
         Height          =   60
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   106
         Color1          =   9594695
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList ilList 
         Left            =   6090
         Top             =   3150
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
               Picture         =   "frmProdEntry.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDeleteProdPack 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   9960
         TabIndex        =   24
         Top             =   1890
         Width           =   825
      End
      Begin VB.CommandButton cmdEditProdPack 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   9150
         TabIndex        =   23
         Top             =   1890
         Width           =   825
      End
      Begin VB.CommandButton cmdAddProdPack 
         Caption         =   "&Add"
         Height          =   315
         Left            =   8340
         TabIndex        =   22
         Top             =   1890
         Width           =   825
      End
      Begin VB.ComboBox cmbProdPack 
         BackColor       =   &H00EAFDFF&
         Height          =   315
         Left            =   6930
         TabIndex        =   21
         Top             =   4050
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox cmbCat 
         BackColor       =   &H00EAFDFF&
         Height          =   315
         ItemData        =   "frmProdEntry.frx":05A6
         Left            =   1650
         List            =   "frmProdEntry.frx":05A8
         TabIndex        =   18
         Top             =   1890
         Width           =   2625
      End
      Begin VB.TextBox txtProdCode 
         Height          =   315
         Left            =   1650
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1020
         Width           =   2655
      End
      Begin VB.ComboBox cmbPack 
         BackColor       =   &H00EAFDFF&
         Height          =   315
         Left            =   8160
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00F8F8F8&
         Caption         =   "Acti&ve"
         Height          =   255
         Left            =   390
         TabIndex        =   9
         Top             =   3840
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtSupPrice 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1050
         Width           =   1635
      End
      Begin VB.TextBox txtSRPrice 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1440
         Width           =   1635
      End
      Begin VB.TextBox txtProdDescription 
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
         Left            =   1650
         MaxLength       =   100
         TabIndex        =   3
         ToolTipText     =   "Customer Name"
         Top             =   1410
         Width           =   3615
      End
      Begin VB.TextBox txtProdID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F5F5F5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   1935
      End
      Begin b8Controls4.LynxGrid3 listProdPack 
         Height          =   2355
         Left            =   5760
         TabIndex        =   25
         Top             =   2250
         Width           =   4995
         _ExtentX        =   8811
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
         BorderStyle     =   0
         FocusRectColor  =   33023
         AllowUserResizing=   4
         Editable        =   -1  'True
         Striped         =   -1  'True
         SBackColor1     =   16056319
         SBackColor2     =   14940667
      End
      Begin b8Controls4.b8GradLine b8GradLine2 
         Height          =   240
         Left            =   0
         TabIndex        =   29
         Top             =   3480
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "   Record Properties"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin b8Controls4.b8GradLine b8GradLine4 
         Height          =   240
         Left            =   0
         TabIndex        =   30
         Top             =   240
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "   General"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin b8Controls4.b8GradLine b8GradLine6 
         Height          =   240
         Left            =   5730
         TabIndex        =   31
         Top             =   240
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "   Main Package"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin b8Controls4.b8GradLine b8GradLine5 
         Height          =   240
         Left            =   5730
         TabIndex        =   28
         Top             =   1890
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "   Sub Packages"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00808080&
         Height          =   2415
         Left            =   5730
         Top             =   2220
         Width           =   5055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   366
         X2              =   366
         Y1              =   12
         Y2              =   318
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Category:"
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   1890
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Main Package:"
         Height          =   195
         Left            =   5940
         TabIndex        =   19
         Top             =   600
         Width           =   1035
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
         TabIndex        =   17
         Top             =   4050
         Width           =   45
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
         TabIndex        =   16
         Top             =   3870
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Supplier Price:"
         Height          =   195
         Left            =   5940
         TabIndex        =   5
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SRP:"
         Height          =   195
         Left            =   5940
         TabIndex        =   7
         Top             =   1500
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Description:"
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
         Left            =   330
         TabIndex        =   2
         Top             =   1470
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         Height          =   195
         Left            =   330
         TabIndex        =   0
         Top             =   1020
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product &ID"
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
         Left            =   330
         TabIndex        =   14
         Top             =   660
         Width           =   900
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
      TabIndex        =   15
      Top             =   0
      Width           =   6975
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
         Left            =   600
         TabIndex        =   32
         Top             =   390
         Width           =   3900
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmProdEntry.frx":05AA
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Products"
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
         TabIndex        =   26
         Top             =   60
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmProdEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mFormState As String

Dim curProd As tProd
Dim newProd As tProd

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isOn As Boolean

Public Function ShowAdd(Optional ByVal sProdDescription As String = "") As Boolean
    
    'set form state
    mFormState = "add"
    
    'set parameter
    newProd.ProdDescription = sProdDescription
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function


Public Function ShowAddRetID(ByRef lProdID As Long, Optional ByVal sProdDescription As String = "") As Boolean
    
    'set form state
    mFormState = "add"
    
    'set parameter
    newProd.ProdDescription = sProdDescription
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAddRetID = mShowAdd
    lProdID = newProd.ProdID
    
End Function

Public Function ShowEdit(ByVal lProdID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curProd.ProdID = lProdID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function






Private Sub cmbCat_GotFocus()
    Dim tmpS As String
    
    tmpS = cmbCat.Text
    
    'load Category List
    modRSCat.FillCatToCMB cmbCat
    
    cmbCat.Text = tmpS
End Sub

Private Sub cmbPack_GotFocus()
    
    Dim tmpS As String
    
    tmpS = cmbPack.Text
    
    'load Package List
    modRSPack.FillPackToCMB cmbPack

    cmbPack.Text = tmpS
    
End Sub

Private Sub cmbPack_LostFocus()
        
    If listProdPack.FindItem(cmbPack.Text, 0, lgSMEqual, False) = 0 Then
        'duplicate
        MsgBox "The package   '" & cmbPack.Text & "'   that you have entered was already in the Other Package list.", vbExclamation
        cmbPack.SetFocus
    End If
    
End Sub

Private Sub cmbProdPack_GotFocus()

    Dim tmpS As String
    
    tmpS = cmbProdPack.Text
    
    'load package list for other package
    modRSPack.FillPackToCMB cmbProdPack

    cmbProdPack.Text = tmpS
    
End Sub

Private Sub cmdSave_Click()

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

Private Sub cmdDeleteProdPack_Click()
    If listProdPack.RowCount > 0 Then
        listProdPack.RemoveItem listProdPack.Row
    End If
End Sub

Private Sub cmdEditProdPack_Click()
    Call listProdPack_DblClick
End Sub




Private Sub Form_Activate()
    
    Me.AutoRedraw = False
    
    'make mouse pointer bussy
    Me.MousePointer = vbHourglass
    
    
    
    
    Select Case mFormState
        Case "add"
                        
            'set form caption
            Me.Caption = "Add New Product Entry"
            
            'generate new Prod ID
            txtProdID.Text = modFunction.ComNumZ(modRSProd.GetNewProdID, 10)

            
        Case "edit"
        
            Dim vPack As tPack
            Dim vCat As tCat
    
            'set form caption
            Me.Caption = "Edit Product Entry"
            
            'get product info
            If GetProdByID(curProd.ProdID, curProd) = False Then
                WriteErrorLog Me.Name, "Form_Activate", "Failed on: 'GetProdByID(curProd.ProdID) = False'"
                'fatal error so we need to close this form
                Unload Me
                GoTo RAE
            End If
            
            
            'set form fields
            With curProd
                txtProdID.Text = modFunction.ComNumZ(.ProdID, 10)
                txtProdCode.Text = .ProdCode
                txtProdDescription.Text = .ProdDescription
                
                'get package
                If modRSPack.GetPackByID(.FK_PackID, vPack) = False Then
                    WriteErrorLog Me.Name, "Form_Activate", "Failed on: 'modRSPack.GetPackByID(.FK_PackID, vPack) = False'"
                    'fatal error so we need to close this form
                    Unload Me
                    GoTo RAE
                End If
                'get category
                If modRSCat.GetCatByID(.FK_CatID, vCat) = False Then
                    WriteErrorLog Me.Name, "Form_Activate", "Failed on: 'modRSCat.GetCatByID(.FK_CatID, vCat) = False'"
                    'fatal error so we need to close this form
                    Unload Me
                    GoTo RAE
                End If
                
                cmbPack.Text = vPack.PackTitle
                cmbCat.Text = vCat.CatTitle
                
                txtSupPrice.Text = .SupPrice
                txtSRPrice.Text = .SRPrice
                
                'fill other packages
                RefreshProdPack .ProdID
                
                chkActive.Value = IIf(.Active = True, vbChecked, vbUnchecked)
                lblRC.Caption = "Created:  " & .RC & "    By: " & .RCU
                If Not IsEmpty(.RMU) Then
                    lblRM.Caption = "Modified: " & .RM & "    By: " & .RMU
                End If
                
            End With
            
    End Select
    
    
RAE:
    'restoremouse pointer tonormal
    Me.MousePointer = vbNormal
    Me.AutoRedraw = True
End Sub


Private Sub Form_Load()

    isOn = False
    PaintGrad bgHeader, &HEDEBE9, &HFFFFFF, 0
  
    
    'set list column
    With listProdPack
        .AddColumn "Package", 90
        .AddColumn "PackID", 0
        .AddColumn "Qty/Unit", 80, lgAlignRightCenter
        .AddColumn "Sup. Price", 80, lgAlignRightCenter
        .AddColumn "SRP", 80, lgAlignRightCenter
        .BindControl 0, cmbProdPack, lgBCLeft Or lgBCTop Or lgBCWidth
        .RowHeightMin = 21
        .ImageList = ilList
    End With
    
End Sub

Private Sub cmbProdPack_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
    
        
        listProdPack.ToggleEdit
        listProdPack.Refresh
    End If
End Sub

Private Sub cmdAddProdPack_Click()

    'validate
    'main package
    If IsEmpty(cmbPack.Text) Then
        MsgBox "Please enter main Package first.", vbExclamation
        cmbPack.SetFocus
        Exit Sub
    End If
    
    'check prices
    If Not (GetTxtVal(txtSupPrice.Text) > 0) Then
        MsgBox "Please enter valid Supplier Price.", vbExclamation
        HLTxt txtSupPrice
        Exit Sub
    End If
    If Not (GetTxtVal(txtSRPrice.Text) > 0) Then
        MsgBox "Please enter valid SRP.", vbExclamation
        HLTxt txtSRPrice
        Exit Sub
    End If
    If GetTxtVal(txtSupPrice.Text) > GetTxtVal(txtSRPrice.Text) Then
        MsgBox "Please enter valid Supplier Price or SRP price." & vbNewLine & _
            "SRP price must be greater or equal to Supplier Price.", vbExclamation
        HLTxt txtSRPrice
        Exit Sub
    End If
    
    'validate prev prod pack item
    If listProdPack.RowCount > 0 Then
        
        
        With listProdPack
            'check package
            If IsEmpty(.CellText(.RowCount - 1, 0)) Then
                MsgBox "Please enter valid Package on previous item first.", vbExclamation
                Exit Sub
            End If
            'check qty
            If Not (GetTxtVal(.CellText(.RowCount - 1, 2)) > 0) Then
                MsgBox "Please enter valid Quantity on previous item first.", vbExclamation
                Exit Sub
            End If
            'check supplier price
            If Not (GetTxtVal(.CellText(.RowCount - 1, 3)) > 0) Then
                MsgBox "Please enter valid Supplier Price on previous item first.", vbExclamation
                Exit Sub
            End If
            'check SRP
            If Not (GetTxtVal(.CellText(.RowCount - 1, 4)) > 0) Then
                MsgBox "Please enter valid SRP on previous item first.", vbExclamation
                Exit Sub
            End If
            
        End With
    End If
    
    

    With listProdPack
        .AddItem ""
        .ItemImage(listProdPack.RowCount - 1) = 1
        .EditCell listProdPack.RowCount - 1, 0
    End With
    
End Sub

Private Sub listProdPack_DblClick()
    If listProdPack.RowCount > 0 Then
        listProdPack.EditCell listProdPack.Row, listProdPack.Col
    End If
End Sub

Private Sub listProdPack_RequestEdit(Row As Long, Col As Long, Cancel As Boolean)
    
    Select Case Col
        Case 0
            cmbProdPack.Text = listProdPack.CellText(Row, Col)
    End Select
    
End Sub

Private Sub listProdPack_RequestUpdate(Row As Long, Col As Long, NewValue As String, Cancel As Boolean)
    
    Dim vPack As tPack
    
    
    Select Case Col
        Case 0 'pack
  
            If IsEmpty(cmbProdPack.Text) = True Then
                Exit Sub
            End If
            
            'check pack duplication
            If LCase(Trim(cmbProdPack.Text)) = LCase(Trim(cmbPack.Text)) Then
                MsgBox "The package   '" & cmbProdPack.Text & "'   that you have entered was already used on Main Package.", vbExclamation
                cmbProdPack.Text = ""
                Cancel = True
                Exit Sub
            End If
            If LCase(Trim(listProdPack.CellText(Row, Col))) <> LCase(Trim(cmbProdPack.Text)) Then
                If listProdPack.FindItem(cmbProdPack.Text, 0, lgSMEqual, False) = 0 Then
                    'duplicate
                    MsgBox "The package   '" & cmbProdPack.Text & "'   that you have entered was already in the Other Package list.", vbExclamation
                    cmbProdPack.Text = ""
                    Cancel = True
                    Exit Sub
                End If
            End If
            
            'add pack if new
            If modRSPack.AddPack(cmbProdPack.Text) = True Then
                NewValue = cmbProdPack.Text
            Else
                WriteErrorLog Me.Name, "listProdPack_RequestUpdate", "Falied on: 'modRSPack.AddPack(cmbProdPack.Text) = True'"
            End If
            
            'add id
            If modRSPack.GetPackByTitle(NewValue, vPack) = True Then
                listProdPack.CellText(Row, 1) = vPack.PackID
            Else
                WriteErrorLog Me.Name, "listProdPack_RequestUpdate", "Falied on: 'modRSPack.GetPackByTitle(NewValue, vPack) = True'"
            End If
            
        
        Case 2 'qty
            If Not (GetTxtVal(NewValue) > 0) Then
                MsgBox "Please enter valid Quantity.", vbExclamation
                Cancel = True
                Exit Sub
            End If
            
            'generate supplier price
            listProdPack.CellText(Row, 3) = FormatNumber(GetTxtVal(NewValue) * GetTxtVal(txtSupPrice.Text), 2)
            'generate SRP
            listProdPack.CellText(Row, 4) = FormatNumber(GetTxtVal(NewValue) * GetTxtVal(txtSRPrice.Text), 2)

            
        Case 3 'supplier price
            If Not (GetTxtVal(NewValue) > 0) Then
                MsgBox "Please enter valid Supplier Price.", vbExclamation
                Cancel = True
            End If
            
        Case 4 'srp
            If Not (GetTxtVal(NewValue) > 0) Then
                MsgBox "Please enter valid SRP.", vbExclamation
                Cancel = True
            End If
            
    End Select
    
End Sub







Private Sub SaveAdd()
    
    Dim tmpProd As tProd
    Dim vPack As tPack
    Dim vCat As tCat
    
    
    
    'validate
    
    'check description
    If IsEmpty(txtProdDescription.Text) Then
        MsgBox "Please enter Description.", vbExclamation
        HLTxt txtProdDescription
        Exit Sub
    End If
    'check description duplication
    If modRSProd.GetProdByDescription(txtProdDescription.Text, tmpProd) = True Then
        MsgBox "The Product Description that you have entered was already existed. Please enter another description.", vbExclamation
        HLTxt txtProdDescription
        Exit Sub
    End If
    
    'check code duplication
    If Not IsEmpty(txtProdCode.Text) Then
        If modRSProd.GetProdByCode(txtProdCode.Text, tmpProd) = True Then
            MsgBox "The Product Code that you have entered was already existed. Please enter another code.", vbExclamation
            HLTxt txtProdCode
            Exit Sub
        End If
    End If
    
    'check package
    If cmbPack.ListIndex < 0 Then
        If IsEmpty(cmbPack.Text) Then
            MsgBox "Please enter valid Package.", vbExclamation
            cmbPack.SetFocus
            Exit Sub
        Else
            'add new package
            modRSPack.AddPack Trim(cmbPack.Text)
        End If
    End If
    
    'set package
    If modRSPack.GetPackByTitle(cmbPack.Text, vPack) = False Then
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: modRSPack.GetPackByTitle(cmbPack.Text, vPack) = False'  |  PackTitle: " & cmbPack.Text
        Exit Sub
    End If
    
    'check Category
    If cmbCat.ListIndex < 0 Then
        If IsEmpty(cmbCat.Text) Then
            MsgBox "Please enter valid Category.", vbExclamation
            cmbCat.SetFocus
            Exit Sub
        Else
            'add new Category
            modRSCat.AddCat Trim(cmbCat.Text)
        End If
    End If
    
    'set Category
    If modRSCat.GetCatByTitle(cmbCat.Text, vCat) = False Then
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: modRSCat.GetCatByTitle(cmbCat.Text, vCat) = False'  |  CatTitle: " & cmbCat.Text
        Exit Sub
    End If
    
    
    'check prices
    If Not (GetTxtVal(txtSupPrice.Text) > 0) Then
        MsgBox "Please enter valid Supplier Price.", vbExclamation
        HLTxt txtSupPrice
        Exit Sub
    End If
    If Not (GetTxtVal(txtSRPrice.Text) > 0) Then
        MsgBox "Please enter valid SRP.", vbExclamation
        HLTxt txtSRPrice
        Exit Sub
    End If
    If GetTxtVal(txtSupPrice.Text) > GetTxtVal(txtSRPrice.Text) Then
        MsgBox "Please enter valid Supplier Price or SRP price." & vbNewLine & _
            "SRP price must be greater or equal to Supplier Price.", vbExclamation
        HLTxt txtSRPrice
        Exit Sub
    End If
    
    'check other packages (last item)
    If listProdPack.RowCount > 0 Then
        With listProdPack
            'check package
            If IsEmpty(.CellText(.RowCount - 1, 0)) Then
                MsgBox "Please enter valid other Package on last item.", vbExclamation
                Exit Sub
            End If
            'check qty
            If Not (GetTxtVal(.CellText(.RowCount - 1, 2)) > 0) Then
                MsgBox "Please enter valid other Quantity on last item.", vbExclamation
                Exit Sub
            End If
            'check supplier price
            If Not (GetTxtVal(.CellText(.RowCount - 1, 3)) > 0) Then
                MsgBox "Please enter valid other Supplier Price on last item.", vbExclamation
                Exit Sub
            End If
            'check SRP
            If Not (GetTxtVal(.CellText(.RowCount - 1, 4)) > 0) Then
                MsgBox "Please enter valid other SRP on last item.", vbExclamation
                Exit Sub
            End If
            
        End With
    End If
    
    
    
    'set new product
    With newProd
        .ProdID = GetTxtVal(txtProdID)

        .ProdCode = Trim(txtProdCode.Text)
        .ProdDescription = Trim(txtProdDescription.Text)
    
        .FK_PackID = vPack.PackID
        .FK_CatID = vCat.CatID
            
        .BegInvStock = 0
        
        .SupPrice = FormatNumber(GetTxtVal(txtSupPrice.Text), 2)
        .SRPrice = FormatNumber(GetTxtVal(txtSRPrice.Text), 2)
        
        .Active = IIf(chkActive.Value = vbChecked, True, False)
        .RC = Now
        .RM = Now
        .RCU = CurrentUser.UserID
        .RMU = ""
        
    End With
    
    'save
    
    If modRSProd.AddProd(newProd) = True Then
        
        'save other packages
        Dim li As Long
        Dim vProdPack As tProdPack
        
        If listProdPack.RowCount > 0 Then
        
            vProdPack.FK_ProdID = newProd.ProdID
            For li = 0 To listProdPack.RowCount - 1
                vProdPack.FK_PackID = GetTxtVal(listProdPack.CellText(li, 1))
                vProdPack.Qty = GetTxtVal(listProdPack.CellText(li, 2))
                vProdPack.SupPrice = GetTxtVal(listProdPack.CellText(li, 3))
                vProdPack.SRPrice = GetTxtVal(listProdPack.CellText(li, 4))
            
                'write
                If modRSProdPack.AddProdPack(vProdPack) = False Then
                    WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSProdPack.AddProdPack(vProdPack) = False'"
                End If
            Next
        End If
        
        'set flag
        mShowAdd = True
        
        'unload this form
        Unload Me
    Else
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSProd.AddProd(newProd) = True'"
    End If
    
End Sub


Private Sub SaveEdit()
    
    Dim tmpProd As tProd
    Dim vPack As tPack
    Dim vCat As tCat
    
    'validate
    
    'check description
    If IsEmpty(txtProdDescription.Text) Then
        MsgBox "Please enter Description.", vbExclamation
        HLTxt txtProdDescription
        Exit Sub
    End If
    
    If LCase(Trim(curProd.ProdDescription)) <> LCase(Trim(txtProdDescription.Text)) Then
        'check description duplication
        If modRSProd.GetProdByDescription(txtProdDescription.Text, tmpProd) = True Then
            MsgBox "The Product Description that you have entered was already existed. Please enter another description.", vbExclamation
            HLTxt txtProdDescription
            Exit Sub
        End If
    End If
    
    'check code duplication
    If Not IsEmpty(txtProdCode.Text) Then
        If LCase(Trim(curProd.ProdCode)) <> LCase(Trim(txtProdCode.Text)) Then
            If modRSProd.GetProdByCode(txtProdCode.Text, tmpProd) = True Then
                MsgBox "The Product Code that you have entered was already existed. Please enter another code.", vbExclamation
                HLTxt txtProdCode
                Exit Sub
            End If
        End If
    End If
    
    'check package
    If cmbPack.ListIndex < 0 Then
        If IsEmpty(cmbPack.Text) Then
            MsgBox "Please enter valid Package.", vbExclamation
            cmbPack.SetFocus
            Exit Sub
        Else
            'add new package
            modRSPack.AddPack Trim(cmbPack.Text)
        End If
    End If
    
    'set package
    If modRSPack.GetPackByTitle(cmbPack.Text, vPack) = False Then
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: modRSPack.GetPackByTitle(cmbPack.Text, vPack) = False'  |  PackTitle: " & cmbPack.Text
        Exit Sub
    End If
    
    'check Category
    If cmbCat.ListIndex < 0 Then
        If IsEmpty(cmbCat.Text) Then
            MsgBox "Please enter valid Category.", vbExclamation
            cmbCat.SetFocus
            Exit Sub
        Else
            'add new Category
            modRSCat.AddCat Trim(cmbCat.Text)
        End If
    End If
    
    'set Category
    If modRSCat.GetCatByTitle(cmbCat.Text, vCat) = False Then
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: modRSCat.GetCatByTitle(cmbCat.Text, vCat) = False'  |  CatTitle: " & cmbCat.Text
        Exit Sub
    End If
    
    
    'check prices
    If Not (GetTxtVal(txtSupPrice.Text) > 0) Then
        MsgBox "Please enter valid Supplier Price.", vbExclamation
        HLTxt txtSupPrice
        Exit Sub
    End If
    If Not (GetTxtVal(txtSRPrice.Text) > 0) Then
        MsgBox "Please enter valid SRP.", vbExclamation
        HLTxt txtSRPrice
        Exit Sub
    End If
    
    If GetTxtVal(txtSupPrice.Text) > GetTxtVal(txtSRPrice.Text) Then
        MsgBox "Please enter valid Supplier Price or SRP price." & vbNewLine & _
            "SRP price must be greater or equal to Supplier Price.", vbExclamation
        HLTxt txtSRPrice
        Exit Sub
    End If
    
    
    'check other packages (last item)
    If listProdPack.RowCount > 0 Then
        With listProdPack
            'check package
            If IsEmpty(.CellText(.RowCount - 1, 0)) Then
                MsgBox "Please enter valid other Package on last item.", vbExclamation
                Exit Sub
            End If
            'check qty
            If Not (GetTxtVal(.CellText(.RowCount - 1, 2)) > 0) Then
                MsgBox "Please enter valid other Quantity on last item.", vbExclamation
                Exit Sub
            End If
            'check supplier price
            If Not (GetTxtVal(.CellText(.RowCount - 1, 3)) > 0) Then
                MsgBox "Please enter valid other Supplier Price on last item.", vbExclamation
                Exit Sub
            End If
            'check SRP
            If Not (GetTxtVal(.CellText(.RowCount - 1, 4)) > 0) Then
                MsgBox "Please enter valid other SRP on last item.", vbExclamation
                Exit Sub
            End If
            
        End With
    End If
    
    
    
    
    'set cur product
    With curProd
        '.ProdID = GetTxtVal(txtProdID)

        .ProdCode = Trim(txtProdCode.Text)
        .ProdDescription = Trim(txtProdDescription.Text)
    
        .FK_PackID = vPack.PackID
        .FK_CatID = vCat.CatID
            
        .BegInvStock = 0
        
        .SupPrice = FormatNumber(GetTxtVal(txtSupPrice.Text), 2)
        .SRPrice = FormatNumber(GetTxtVal(txtSRPrice.Text), 2)
        
        .Active = IIf(chkActive.Value = vbChecked, True, False)
        '.RC = Now
        .RM = Now
        '.RCU = CurrentUser.UserID
        .RMU = CurrentUser.UserID
        
    End With
    
    'save
    
    If modRSProd.EditProd(curProd) = True Then
        
        'delete all other packages
        If modRSProdPack.DeleteAllProdPack(curProd.ProdID) = False Then
            WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSProdPack.DeleteAllProdPack(curProd.ProdID) = False'"
        End If
        
        'save other packages
        Dim li As Long
        Dim vProdPack As tProdPack
        
        If listProdPack.RowCount > 0 Then
        
            vProdPack.FK_ProdID = curProd.ProdID
            For li = 0 To listProdPack.RowCount - 1
                vProdPack.FK_PackID = GetTxtVal(listProdPack.CellText(li, 1))
                vProdPack.Qty = GetTxtVal(listProdPack.CellText(li, 2))
                vProdPack.SupPrice = GetTxtVal(listProdPack.CellText(li, 3))
                vProdPack.SRPrice = GetTxtVal(listProdPack.CellText(li, 4))
            
                'write
                If modRSProdPack.AddProdPack(vProdPack) = False Then
                    WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSProdPack.AddProdPack(vProdPack) = False'"
                End If
            Next
        End If
        
        'set flag
        mShowEdit = True
        
        'unload this form
        Unload Me
    Else
        WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'modRSProd.EditProd(newProd) = True'"
    End If
            
End Sub



Public Sub RefreshProdPack(ByVal lFK_ProdID As Long)
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long
    
    
    listProdPack.Redraw = False
    listProdPack.Clear
    
    
    sSQL = " SELECT tblProdPack.FK_PackID, tblPack.PackTitle, tblProdPack.Qty, tblProdPack.SupPrice, tblProdPack.SRPrice" & _
            " FROM tblPack INNER JOIN tblProdPack ON tblPack.PackID = tblProdPack.FK_PackID" & _
            " Where (((tblProdPack.FK_ProdID) = " & lFK_ProdID & "))" & _
            " ORDER BY tblPack.PackTitle;"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "RefreshProdPack", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    'fill
    vRS.MoveFirst
    While vRS.EOF = False
        With listProdPack
            li = .AddItem(ReadField(vRS.Fields("PackTitle")))
            .ItemImage(li) = 1
            .CellText(li, 1) = ReadField(vRS.Fields("FK_PackID"))
            .CellText(li, 2) = ReadField(vRS.Fields("Qty"))
            .CellText(li, 3) = FormatNumber(ReadField(vRS.Fields("SupPrice")), 2)
            .CellText(li, 4) = FormatNumber(ReadField(vRS.Fields("SRPrice")), 2)
        End With
        vRS.MoveNext
    Wend
    
    
RAE:
    Set vRS = Nothing
    listProdPack.Redraw = True
    listProdPack.Refresh
End Sub
