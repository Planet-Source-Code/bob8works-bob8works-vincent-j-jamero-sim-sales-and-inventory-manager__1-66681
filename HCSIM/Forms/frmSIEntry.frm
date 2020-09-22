VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSIEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Invoice"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSIEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   775
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
      Left            =   8580
      TabIndex        =   0
      Top             =   6960
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   10140
      TabIndex        =   1
      Top             =   6960
      Width           =   1395
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   6885
      Left            =   0
      ScaleHeight     =   459
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   775
      TabIndex        =   2
      Top             =   540
      Width           =   11625
      Begin MSComctlLib.ImageList ilList 
         Left            =   3960
         Top             =   4710
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
               Picture         =   "frmSIEntry.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtRemarks 
         Height          =   765
         Left            =   8190
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         Top             =   5400
         Width           =   3165
      End
      Begin VB.PictureBox Picture1 
         Height          =   2985
         Left            =   990
         ScaleHeight     =   195
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   457
         TabIndex        =   48
         Top             =   3180
         Width           =   6915
         Begin b8Controls4.LynxGrid3 listSIProd 
            Height          =   2925
            Left            =   0
            TabIndex        =   49
            Top             =   30
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   5159
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
            Striped         =   -1  'True
            SBackColor1     =   16056319
            SBackColor2     =   14940667
         End
      End
      Begin VB.CommandButton cmdEditSIProd 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   345
         Left            =   6630
         TabIndex        =   47
         Top             =   2820
         Width           =   645
      End
      Begin VB.TextBox txtRefNum 
         Height          =   315
         Left            =   5730
         MaxLength       =   20
         TabIndex        =   45
         Top             =   120
         Width           =   2205
      End
      Begin VB.CommandButton cmdNewProd 
         Height          =   345
         Left            =   7530
         Picture         =   "frmSIEntry.frx":05A6
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2220
         Width           =   375
      End
      Begin VB.TextBox txtSIBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F4FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   9330
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   42
         Text            =   "0.00"
         Top             =   4560
         Width           =   1995
      End
      Begin VB.TextBox txtPayAmtOnDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F4FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   9330
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   40
         Text            =   "0.00"
         Top             =   3810
         Width           =   2025
      End
      Begin VB.TextBox txtTotalAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F4FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   9300
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   1980
         Width           =   2025
      End
      Begin VB.ComboBox cmbFP 
         Height          =   315
         Left            =   9720
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3420
         Width           =   1635
      End
      Begin VB.ComboBox cmbCA 
         Height          =   315
         Left            =   9720
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3000
         Width           =   1635
      End
      Begin VB.CommandButton cmdAddSIProd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5970
         TabIndex        =   35
         Top             =   2820
         Width           =   645
      End
      Begin VB.CommandButton cmdDeleteSIProd 
         Caption         =   "&Del"
         Enabled         =   0   'False
         Height          =   345
         Left            =   7290
         TabIndex        =   31
         Top             =   2820
         Width           =   615
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F4FFFF&
         Height          =   315
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   29
         Top             =   2850
         Width           =   975
      End
      Begin VB.TextBox txtQtyPrice 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   3450
         MaxLength       =   50
         TabIndex        =   27
         Top             =   2850
         Width           =   825
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F4FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   990
         MaxLength       =   50
         TabIndex        =   25
         Top             =   2850
         Width           =   855
      End
      Begin VB.ComboBox cmbPackTitle 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2850
         Width           =   1515
      End
      Begin VB.CommandButton cmdNewSup 
         Height          =   345
         Left            =   7530
         Picture         =   "frmSIEntry.frx":0B30
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   900
         Width           =   375
      End
      Begin b8Controls4.b8Line b8Line1 
         Height          =   30
         Left            =   -30
         TabIndex        =   14
         Top             =   480
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin MSComCtl2.DTPicker dtpSIDate 
         Height          =   315
         Left            =   9810
         TabIndex        =   12
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM - dd - yyyy"
         Format          =   118947843
         CurrentDate     =   38961
      End
      Begin b8Controls4.b8DataPicker b8DPCust 
         Height          =   360
         Left            =   990
         TabIndex        =   9
         Top             =   900
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
         DropWinWidth    =   6210
      End
      Begin b8Controls4.b8DataPicker b8DPProd 
         Height          =   360
         Left            =   990
         TabIndex        =   8
         Top             =   2220
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
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00EAFDFF&
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   6915
      End
      Begin VB.TextBox txtSIID 
         BackColor       =   &H00F5F5F5&
         Height          =   285
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   120
         Width           =   1635
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8GradLine b8GradLine1 
         Height          =   240
         Left            =   0
         TabIndex        =   10
         Top             =   540
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "      Customer"
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
      Begin b8Controls4.b8GradLine b8GradLine2 
         Height          =   240
         Left            =   0
         TabIndex        =   21
         Top             =   1860
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "      Purchased Products"
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
      Begin b8Controls4.b8Line b8Line3 
         Height          =   30
         Left            =   8220
         TabIndex        =   32
         Top             =   4410
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line4 
         Height          =   30
         Left            =   0
         TabIndex        =   33
         Top             =   1800
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line5 
         Height          =   30
         Left            =   0
         TabIndex        =   34
         Top             =   6300
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line6 
         Height          =   30
         Left            =   8130
         TabIndex        =   39
         Top             =   2550
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line7 
         Height          =   30
         Left            =   8190
         TabIndex        =   52
         Top             =   4980
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Remarks:"
         Height          =   195
         Left            =   8190
         TabIndex        =   51
         Top             =   5190
         Width           =   675
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. #:"
         Height          =   195
         Left            =   5130
         TabIndex        =   46
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Balance"
         Height          =   195
         Left            =   8220
         TabIndex        =   43
         Top             =   4590
         Width           =   555
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Amount:"
         Height          =   195
         Left            =   8250
         TabIndex        =   41
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Less Payment:"
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
         Left            =   8130
         TabIndex        =   38
         Top             =   2670
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Total Amount:"
         Height          =   195
         Left            =   8220
         TabIndex        =   37
         Top             =   2010
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form of Payment:"
         Height          =   195
         Left            =   8220
         TabIndex        =   19
         Top             =   3420
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charge Account:"
         Height          =   195
         Left            =   8220
         TabIndex        =   16
         Top             =   3030
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Amount"
         Height          =   195
         Left            =   4320
         TabIndex        =   30
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Unit Price"
         Height          =   195
         Left            =   3450
         TabIndex        =   28
         Top             =   2640
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Qty:"
         Height          =   195
         Left            =   990
         TabIndex        =   26
         Top             =   2640
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
         Height          =   195
         Left            =   1920
         TabIndex        =   24
         Top             =   2640
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product:"
         Height          =   195
         Left            =   300
         TabIndex        =   23
         Top             =   2250
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Purchased:"
         Height          =   195
         Left            =   8280
         TabIndex        =   13
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   1380
         Width           =   645
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
         Left            =   9480
         TabIndex        =   5
         Top             =   3000
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         Height          =   195
         Left            =   690
         TabIndex        =   4
         Top             =   120
         Width           =   225
      End
   End
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   6
      Top             =   0
      Width           =   10305
      Begin VB.Label Label18 
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
         Left            =   660
         TabIndex        =   54
         Top             =   330
         Width           =   3900
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmSIEntry.frx":10BA
         Top             =   30
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Invoice"
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
         Left            =   660
         TabIndex        =   53
         Top             =   0
         Width           =   1905
      End
   End
End
Attribute VB_Name = "frmSIEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mFormState As String

Dim ProdPackList() As tProdPack

Dim curSI As tSI
Dim newSI As tSI

Dim curCustPay As tCustPay

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isOn As Boolean


Public Function ShowAdd(Optional ByVal dSIDate As Date = 0, Optional ByVal lCustID As Long = 0) As Boolean
    
    'set form state
    mFormState = "add"
    
    'evaluate param
    If dSIDate = 0 Then
        newSI.SIDate = Now
    Else
        newSI.SIDate = dSIDate
    End If
    newSI.FK_CustID = lCustID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function

Public Function ShowEdit(ByVal lSIID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curSI.SIID = lSIID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function



Private Sub cmdCancel_Click()
    
    Select Case mFormState
        Case "add"
            mShowAdd = False
        Case "edit"
            mShowEdit = False
    End Select

    
    Unload Me
End Sub


Private Sub cmdEditSIProd_Click()
    Call listSIProd_DblClick
End Sub

Private Sub cmdNewProd_Click()
        
    Dim lProdID As Long
    Dim vProd As tProd
    
    If frmProdEntry.ShowAddRetID(lProdID) = False Then
        Exit Sub
    End If
    
    If GetProdByID(lProdID, vProd) = False Then
        Exit Sub
    End If
    
    
    b8DPProd.BoundData = lProdID
    b8DPProd.DisplayData = vProd.ProdDescription
    
    
    If RefeshCurSIProd(lProdID) = False Then
        Exit Sub
    End If

    'set focused control
    HLTxt txtQty
End Sub

Private Sub cmdNewSup_Click()
    
    Dim lCustID As Long
    Dim vCust As tCust
    
    If frmCustEntry.ShowAddRetID(lCustID) = False Then
        Exit Sub
    End If
    
    If GetCustByID(lCustID, vCust) = False Then
        Exit Sub
    End If
    
    b8DPCust.BoundData = vCust.CustID
    b8DPCust.DisplayData = vCust.CustName
    
    Call b8DPCust_Change
    
End Sub

Private Sub cmdSave_Click()

    Select Case mFormState
        Case "add"
            SaveAdd
        Case "edit"
            SaveEdit
    End Select
    
End Sub

Private Sub dtpSIDate_Change()
    Form_UseThisCust CLng(GetTxtVal(b8DPCust.BoundData))
End Sub

Private Sub Form_Activate()
        
    
    If isOn = True Then
        Exit Sub
    End If
    isOn = True

    DoEvents: DoEvents: DoEvents
    
    'make mouse Pointer bussy
    Me.MousePointer = vbHourglass
   
    Select Case mFormState
        Case "add"
        
            Me.Caption = "Add New Sales InvoiceEntry"
            
            'add form of payment list
            Form_RefreshFP
            
            'add charge account list
            Form_RefreshCA
                        
            'set form fields
            txtSIID.Text = modFunction.ComNumZ(modRSSI.GetNewSIID, 10)
            dtpSIDate.Value = newSI.SIDate
            
            If newSI.FK_CustID > 0 Then
                Form_UseThisCust newSI.FK_CustID
            End If
            
            '
            CAFPChange
            
            
        Case "edit"
        
            Me.Caption = "Edit Sales InvoiceEntry"
       
            If GetSIByID(curSI.SIID, curSI) = False Then
                'WriteErrorLog Me.Name, "Form_Activate", "Failed on: 'GetSIByID(curSI.SIID, vSI) = False'"
                Unload Me
                GoTo RAE
            End If
            
            txtSIID.Text = modFunction.ComNumZ(curSI.SIID, 10)
            txtRefNum.Text = curSI.RefNum
            dtpSIDate.Value = curSI.SIDate
            'set form ui
            Form_UseThisCust curSI.FK_CustID
            
            'load products
            LoadProducts curSI.SIID
            
            'load Customer payment if is has
            If modRSCustPay.GetCustPayByID(curSI.OptFK_CustPayID, curCustPay) = True Then
                       
                'add form of payment list
                Form_RefreshFP curCustPay.FP
                           
                'add charge account list
                Form_RefreshCA IIf(curCustPay.Amount = curSI.TotalAmt, "Full Payment", "Partial Payment")
    
                'reasign FP
                Form_RefreshFP curCustPay.FP
                
                txtPayAmtOnDate.Text = FormatNumber(curCustPay.Amount, 2)
            Else
                
                'add form of payment list
                Form_RefreshFP "other"
                
                Form_RefreshCA "not paid"
                
            End If
            
            txtRemarks.Text = curSI.Remarks
            
            'calculate
            Call Form_CalTotalAmount
            
            
    End Select
    
    
RAE:
    'restoremouse Pointer tonormal
    Me.MousePointer = vbNormal
End Sub


Private Sub Form_Load()
    
    isOn = False
    
    PaintGrad bgHeader, &HEDEBE9, &HFFFFFF, 0

    'set SI SIrd column headers
    With listSIProd
    
        .AddColumn "Qty.", 70, lgAlignRightCenter '0
        .AddColumn "InvQty", 0 '1
        .AddColumn "FK_PackID", 0 '2
        .AddColumn "Unit", 80 '3
        .AddColumn "Product ID", 0 '4
        .AddColumn "Articles", 120 '5
        .AddColumn "Unit Price", 70, lgAlignRightCenter '6
        .AddColumn "Amount", 90, lgAlignRightCenter '7
        
        .RowHeightMin = 21
        .ImageList = ilList
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    isOn = False
End Sub



Private Sub SaveAdd()
        
    Dim lCustPayID As Long
    Dim newCustPay As tCustPay
    
    Dim dNewAmount As Double
        
    'default

    
    'validate
    'reference number
    If Len(Trim(txtRefNum.Text)) < 1 Then
        MsgBox "Please enter 'Reference Number'.", vbExclamation
        HLTxt txtRefNum
        Exit Sub
    End If
    
    'Customer
    If Len(Trim(b8DPCust.BoundData)) < 1 Then
        MsgBox "Please enter 'Customer'.", vbExclamation
        b8DPCust.FocusedDropButton
        Exit Sub
    End If
    
    'products
    If Not (GetTxtVal(txtTotalAmount.Text) > 0) Then
        MsgBox "Enter some purchased Product first.", vbExclamation
        b8DPProd.FocusedDropButton
        Exit Sub
    End If

  
    Select Case cmbCA.ListIndex
    
        Case 0 To 1 'full or partial

            Dim sRemarks As String
            sRemarks = "Payment for Sales Invoice with Ref # " & Trim(txtRefNum.Text)

            If frmCustPayEntry.ShowAdd(dtpSIDate.Value, CLng(GetTxtVal(b8DPCust.BoundData)), cmbFP.Text, GetTxtVal(txtPayAmtOnDate.Text), GetTxtVal(txtTotalAmount.Text), sRemarks, lCustPayID) = False Then
                Exit Sub
            End If
            
            If GetCustPayByID(lCustPayID, newCustPay) = False Then
                WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'GetCustPayByID(lCustPayID, newCustPay) = False'"
                Exit Sub
            End If

            With newSI
                .OptFK_CustPayID = lCustPayID
                .TotalAmt = GetTxtVal(txtTotalAmount.Text)
                .Remarks = txtRemarks.Text
            End With
            

        Case 2 'not paid
            cmbFP.ListIndex = 2

            With newSI
                .TotalAmt = GetTxtVal(txtTotalAmount.Text)
                .Remarks = txtRemarks.Text
            End With
            
    End Select
    

    'set remaining new SI info
    With newSI
        .SIID = CLng(GetTxtVal(txtSIID.Text))
        .RefNum = Trim(txtRefNum.Text)
        .FK_CustID = CLng(GetTxtVal(b8DPCust.BoundData))
        
        ' + 1 second
        .SIDate = modFunction.GetRSec(dtpSIDate.Value) + (1 / 86400)
                
        .RC = Now
        'RM
        .RCU = CurrentUser.UserID
        'RMU
    End With
    
    'write new SI
    If modRSSI.AddSI(newSI) = True Then
        
        'add SI Items(Products)
        Dim newSIProd As tSIProd
        Dim li As Long
        
        For li = 0 To listSIProd.RowCount - 1
            With newSIProd
                .FK_SIID = newSI.SIID
                .FK_ProdID = Val(listSIProd.CellText(li, 4))
                .Qty = GetTxtVal(listSIProd.CellText(li, 0))
                .InvQty = GetTxtVal(listSIProd.CellText(li, 1))
                .FK_PackID = GetTxtVal(listSIProd.CellText(li, 2))
                .UnitPrice = GetTxtVal(listSIProd.CellText(li, 6))
                .Amount = GetTxtVal(listSIProd.CellText(li, 7))
            End With
            
            If modRSSIProd.AddSIProd(newSIProd, newSI) = False Then
                WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSSIProd.AddSIProd(newSIProd, newSI) = False'"
            End If

        Next
                
        'set flag
        mShowAdd = True
        'close this form
        Unload Me
        
    Else
    
        'delete saved pts
        If modRSCustPay.DeleteCustPay(lCustPayID) = False Then
            WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSCustPay.DeleteCustPay(lCustPayID) = False'"
        End If
        
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSSI.AddSI(newSI) = True'"
    End If
        
    
End Sub


Private Sub SaveEdit()
    
    Dim lCustPayID As Long
    Dim curCustPay As tCustPay
    Dim dNewAmount As Double
    Dim tmpCustPay As tCustPay
    
    
    'default
    
    lCustPayID = -1
    
    
    'validate
    'reference number
    If Len(Trim(txtRefNum.Text)) < 1 Then
        MsgBox "Please enter 'Reference Number'.", vbExclamation
        HLTxt txtRefNum
        Exit Sub
    End If
    
    'Customer
    If Len(Trim(b8DPCust.BoundData)) < 1 Then
        MsgBox "Please enter 'Customer'.", vbExclamation
        b8DPCust.FocusedDropButton
        Exit Sub
    End If
    
    'products
    If Not (GetTxtVal(txtTotalAmount.Text) > 0) Then
        MsgBox "Enter some purchased Product first.", vbExclamation
        b8DPProd.FocusedDropButton
        Exit Sub
    End If

  
    Select Case cmbCA.ListIndex
    
        Case 0 To 1 'full or partial

            Dim sRemarks As String
            sRemarks = "Payment for Sales Invoice with Ref # " & Trim(txtRefNum.Text)

            If GetCustPayByID(curSI.OptFK_CustPayID, tmpCustPay) = True Then
                If frmCustPayEntry.ShowEdit(curSI.OptFK_CustPayID) = False Then
                    Exit Sub
                End If
            Else
                If frmCustPayEntry.ShowAdd(dtpSIDate.Value, CLng(GetTxtVal(b8DPCust.BoundData)), cmbFP.Text, GetTxtVal(txtPayAmtOnDate.Text), GetTxtVal(txtTotalAmount.Text), sRemarks, curSI.OptFK_CustPayID) = False Then
                    Exit Sub
                End If
            End If
            
            
            If GetCustPayByID(curSI.OptFK_CustPayID, curCustPay) = False Then
                WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'GetCustPayByID(curSI.OptFK_CustPayID, curCustPay) = False'"
                Exit Sub
            End If
            
            With curSI
                
                .TotalAmt = GetTxtVal(txtTotalAmount.Text)
                
                If curCustPay.FP = "check" Then
                    If curCustPay.Cleared = True Then
                        dNewAmount = curCustPay.Amount
                    Else
                        dNewAmount = 0
                    End If
                Else
                    dNewAmount = curCustPay.Amount
                End If
              .Remarks = txtRemarks.Text
                
            End With
            

        Case 2 'not paid
            cmbFP.ListIndex = 2

            With curSI

                .TotalAmt = GetTxtVal(txtTotalAmount.Text)
                .Remarks = txtRemarks.Text
            End With
            
            'Delete pts
            If GetCustPayByID(curSI.OptFK_CustPayID, tmpCustPay) = True Then
                modRSCustPay.DeleteCustPay curSI.OptFK_CustPayID
            End If
            
    End Select

    'set remaining new SI info
    With curSI
        .SIID = CLng(GetTxtVal(txtSIID.Text))
        .RefNum = Trim(txtRefNum.Text)
        .FK_CustID = CLng(GetTxtVal(b8DPCust.BoundData))
        
        ' + 1 second
        '.SIDate = modFunction.GetRSec(dtpSIDate.Value) + (1 / 86400)
                
        '.RC = Now
        .RM = Now
        '.RCU
        .RMU = CurrentUser.UserID
    End With
    
    'write new SI
    If modRSSI.EditSI(curSI) = True Then
        
        'add SI Items(Products)
        Dim curSIProd As tSIProd
        Dim li As Long
        
        'delete old SI prod
        DeleteAllSIProd curSI.SIID, curSI
        
        For li = 0 To listSIProd.RowCount - 1
            With curSIProd
                .FK_SIID = curSI.SIID
                .FK_ProdID = Val(listSIProd.CellText(li, 4))
                .Qty = GetTxtVal(listSIProd.CellText(li, 0))
                .InvQty = GetTxtVal(listSIProd.CellText(li, 1))
                .FK_PackID = GetTxtVal(listSIProd.CellText(li, 2))
                .UnitPrice = GetTxtVal(listSIProd.CellText(li, 6))
                .Amount = GetTxtVal(listSIProd.CellText(li, 7))
            End With
            
            If modRSSIProd.AddSIProd(curSIProd, curSI) = False Then
                WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'modRSSIProd.AddSIProd(curSIProd, curSI) = False'"
            End If

        Next
                
        'set flag
        mShowEdit = True
        'close this form
        Unload Me
        
    Else
    
        'delete saved pts
        If modRSCustPay.DeleteCustPay(curSI.OptFK_CustPayID) = False Then
            WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'modRSCustPay.DeleteCustPay(lCustPayID) = False'"
        End If
        
        WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'modRSSI.AddSI(curSI) = True'"
    End If
    
End Sub









'---------------------------------------------------------------
'Customer Info Procedures
'---------------------------------------------------------------
Private Sub Form_UseThisCust(ByVal lCustID As Long)
    
    Dim vCust As tCust
    
    txtAddress.Text = ""
    
    If modRSCust.GetCustByID(lCustID, vCust) = True Then
    
        b8DPCust.BoundData = vCust.CustID
        b8DPCust.DisplayData = vCust.CustName
        
        txtAddress.Text = IIf(IsEmpty(vCust.AddrStreet), " __, ", vCust.AddrStreet & ", ") & _
                            IIf(IsEmpty(vCust.AddrBrgy), " __, ", vCust.AddrBrgy & ", ") & _
                            IIf(IsEmpty(vCust.AddrCity), " __, ", vCust.AddrCity & ", ") & _
                            IIf(IsEmpty(vCust.AddrProvince), " __ ", vCust.AddrProvince)
        
        
    End If
    
End Sub


Private Sub b8DPProd_Change()

    If RefeshCurSIProd(CLng(GetTxtVal(b8DPProd.BoundData))) = False Then
        Exit Sub
    End If

    'set focused control
    HLTxt txtQty
    
End Sub

Private Sub b8DPCust_Change()
    
    Form_UseThisCust CLng(GetTxtVal(b8DPCust.BoundData))

End Sub



'---------------------------------------------------------------
'>>> END Customer Info Procedures
'---------------------------------------------------------------





'---------------------------------------------------------------
'Product Info Procedures
'---------------------------------------------------------------
Private Sub LoadProducts(ByVal lSIID As Long)
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long



    sSQL = "SELECT tblSIProd.Qty, tblSIProd.InvQty, tblProd.FK_PackID, tblPack.PackTitle, tblSIProd.FK_ProdID, tblProd.ProdDescription, tblSIProd.UnitPrice, tblSIProd.Amount, tblSIProd.FK_SIID" & _
            " FROM tblPack INNER JOIN (tblProd INNER JOIN tblSIProd ON tblProd.ProdID = tblSIProd.FK_ProdID) ON tblPack.PackID = tblProd.FK_PackID" & _
            " WHERE tblSIProd.FK_SIID=" & lSIID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "LoadProducts", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    listSIProd.Redraw = False
    listSIProd.Clear

    vRS.MoveFirst
    While vRS.EOF = False
        
        With listSIProd
        li = .AddItem(CStr(ReadField(vRS.Fields("Qty"))))
        .ItemImage(li) = 1
        .CellText(li, 1) = ReadField(vRS.Fields("InvQty"))
        .CellText(li, 2) = ReadField(vRS.Fields("FK_PackID"))
        .CellText(li, 3) = ReadField(vRS.Fields("PackTitle"))
        .CellText(li, 4) = ReadField(vRS.Fields("FK_ProdID"))
        .CellText(li, 5) = ReadField(vRS.Fields("ProdDescription"))
        .CellText(li, 6) = ReadField(vRS.Fields("UnitPrice"))
        .CellText(li, 7) = ReadField(vRS.Fields("Amount"))
        End With
        
        vRS.MoveNext
    Wend
    
RAE:
    listSIProd.Redraw = True
    listSIProd.Refresh
    Set vRS = Nothing

End Sub

Private Sub cmdDeleteSIProd_Click()
    If listSIProd.RowCount > 0 Then
        listSIProd.RemoveItem listSIProd.Row
    
        'calculate total amount
        Call Form_CalTotalAmount
    
    End If
End Sub


Private Function RefeshCurSIProd(ByVal lProdID As Long, Optional dQty As Double = 0, Optional lPackID As Long = 0, Optional dPrice As Double = 0) As Boolean
    
    Dim i As Integer
    Dim vProd As tProd

    'default
    RefeshCurSIProd = False
    
    'clear & Disable
    cmbPackTitle.Clear
    txtQty.Text = dQty
    txtQtyPrice.Text = dPrice
    txtAmount.Text = ""
    
    txtQty.Enabled = False
    cmbPackTitle.Enabled = False
    txtQtyPrice.Enabled = False
    
    cmdAddSIProd.Enabled = False
    
    If GetProdByID(lProdID, vProd) = False Then
        Exit Function
    End If
    
    'fill packages
    If modRSProdPack.FillProdPackToTypeArray(lProdID, ProdPackList) = False Then
        WriteErrorLog Me.Name, "RefeshCurSIProd", "Failed on: 'modRSProdPack.FillProdPackToTypeArray(lProdID, prodpacklis) = False'"
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
        
    txtQty.Enabled = True
    txtQtyPrice.Enabled = True
    
    'return sucess
    RefeshCurSIProd = True
    
End Function

Private Sub cmbPackTitle_Change()

    cmdAddSIProd.Enabled = False
    
    If cmbPackTitle.ListIndex >= 0 Then
        txtQtyPrice.Text = FormatNumber(ProdPackList(cmbPackTitle.ListIndex).SRPrice, 2)
    Else
        Exit Sub
    End If
    
    'generate Amount
    If GetTxtVal(txtQty.Text) > 0 Then
        txtAmount.Text = FormatNumber(GetTxtVal(txtQty.Text) * GetTxtVal(txtQtyPrice.Text), 2)
    Else
        Exit Sub
    End If
    
    If Not GetTxtVal(txtAmount.Text) > 0 Then
        Exit Sub
    End If
    
    'sucess
    'enable Add
    cmdAddSIProd.Enabled = True
    
End Sub

Private Sub cmbPackTitle_Click()
    Call cmbPackTitle_Change
End Sub

Private Sub listSIProd_DblClick()

    With listSIProd
    
    If .RowCount > 0 Then

        b8DPProd.BoundData = CLng(GetTxtVal(.CellText(.Row, 4)))
        b8DPProd.DisplayData = .CellText(.Row, 5)
        
        RefeshCurSIProd CLng(GetTxtVal(.CellText(.Row, 4))), GetTxtVal(.CellText(.Row, 0)), CLng(GetTxtVal(.CellText(.Row, 2))), _
                        GetTxtVal(.CellText(.Row, 6))
    End If
    
    End With
End Sub

Private Sub listSIProd_ItemCountChanged()

    If listSIProd.RowCount > 0 Then
        cmdDeleteSIProd.Enabled = True
        cmdEditSIProd.Enabled = True
    Else
        cmdDeleteSIProd.Enabled = False
        cmdEditSIProd.Enabled = False
        txtTotalAmount.Text = "0.00"
    End If
End Sub



Private Sub txtQty_Change()
    'generate amount
    Call cmbPackTitle_Change
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'add item
        Call cmdAddSIProd_Click
    End If
End Sub

Private Sub txtQtyPrice_Change()

    cmdAddSIProd.Enabled = False
    
    If Not (cmbPackTitle.ListIndex >= 0) Then
        Exit Sub
    End If
    
    'generate Amount
    If GetTxtVal(txtQty.Text) > 0 Then
        txtAmount.Text = FormatNumber(GetTxtVal(txtQty.Text) * GetTxtVal(txtQtyPrice.Text), 2)
    Else
        Exit Sub
    End If
    
    If Not GetTxtVal(txtAmount.Text) > 0 Then
        Exit Sub
    End If
    
    'sucess
    'enable Add
    cmdAddSIProd.Enabled = True
    
End Sub

Private Sub cmdAddSIProd_Click()

    Dim li As Long
    Dim lProdID As Long
    Dim dupli As Long
    
    'validate
    If Not (GetTxtVal(txtQty.Text) > 0) Then
        MsgBox "Please enter valid 'Quantity'", vbExclamation
        HLTxt txtQty
        Exit Sub
    End If
    
    If Not (GetTxtVal(txtQtyPrice.Text) > 0) Then
        MsgBox "Please enter valid 'Unit Price'", vbExclamation
        HLTxt txtQtyPrice
        Exit Sub
    End If
    
    'check if the product is already in the list
    lProdID = CLng(GetTxtVal(b8DPProd.BoundData))
    dupli = listSIProd.FindItem(CStr(lProdID), 4, lgSMEqual, False)
    
    If dupli >= 0 Then
        If MsgBox("The Product that you have added is already in the list." & vbNewLine & vbNewLine & _
            "Do you want to replace it?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            'the answer is YES
            'remove old
            listSIProd.RemoveItem dupli
        Else
            'the answer is NO
            Exit Sub
        End If
    End If
            
    
    With listSIProd
        .Redraw = False
        
        li = .AddItem(CStr(GetTxtVal(txtQty.Text)))
        .ItemImage(li) = 1
        .CellText(li, 1) = ProdPackList(cmbPackTitle.ListIndex).Qty * GetTxtVal(txtQty.Text)
        .CellText(li, 2) = ProdPackList(cmbPackTitle.ListIndex).FK_PackID
        .CellText(li, 3) = cmbPackTitle.Text
        .CellText(li, 4) = lProdID
        .CellText(li, 5) = b8DPProd.DisplayData
        .CellText(li, 6) = FormatNumber(GetTxtVal(txtQtyPrice.Text), 2)
        .CellText(li, 7) = FormatNumber(GetTxtVal(txtAmount.Text), 2)
        
        .Redraw = True
        .Refresh
    End With
    
    
    'calculate total amount
    Call Form_CalTotalAmount
    
    'clear & Disable
    cmbPackTitle.Clear
    txtQty.Text = ""
    txtQtyPrice.Text = ""
    txtAmount.Text = ""
    
    b8DPProd.ClearCurData
    
    txtQty.Enabled = False
    cmbPackTitle.Enabled = False
    txtQtyPrice.Enabled = False
    
    cmdAddSIProd.Enabled = False
    
    'set focused on next control
    b8DPProd.FocusedDropButton
    

End Sub

Private Sub Form_CalTotalAmount()

    Dim li As Long
    Dim dTA As Double
    
    'clear
    txtSIBalance.Text = "0.00"
    
    dTA = 0
    For li = 0 To listSIProd.RowCount - 1
        dTA = dTA + GetTxtVal(listSIProd.CellText(li, 7))
    Next
    
    txtTotalAmount.Text = FormatNumber(dTA, 2)
    
    If GetTxtVal(txtPayAmtOnDate.Text) < 0 Then
        Exit Sub
    End If
    
    txtSIBalance.Text = FormatNumber(GetTxtVal(txtTotalAmount.Text) - GetTxtVal(txtPayAmtOnDate.Text), 2)
    
End Sub
'---------------------------------------------------------------
' >>> END Product Info Procedures
'---------------------------------------------------------------



'---------------------------------------------------------------
'Payment Info Procedures
'---------------------------------------------------------------

Private Sub txtPayAmtOnDate_Change()

    txtSIBalance.Text = "0.00"
    
    If GetTxtVal(txtTotalAmount.Text) < 0 Then
        Exit Sub
    End If
    
    If GetTxtVal(txtPayAmtOnDate.Text) < 0 Then
        Exit Sub
    End If
    
    txtSIBalance.Text = FormatNumber(GetTxtVal(txtTotalAmount.Text) - GetTxtVal(txtPayAmtOnDate.Text), 2)
    
End Sub

Private Sub cmbCA_Change()
    
    'disable affected controls
    cmbFP.Enabled = False
    
    Select Case cmbCA.ListIndex
        Case 0 'full
            cmbFP.Enabled = True
            'set FA to cash
            cmbFP.ListIndex = 0
        Case 1 'partial
            cmbFP.Enabled = True
            'set FA to cash
            cmbFP.ListIndex = 0
        Case 2 'no paid
            cmbFP.ListIndex = 2
            
    End Select
    
    
    
End Sub

Private Sub cmbCA_Click()
    
    Call cmbCA_Change
    
    Call CAFPChange
    
End Sub

Private Sub cmbFP_Change()
    Call CAFPChange
End Sub

Private Sub cmbFP_Click()
    Call CAFPChange
End Sub

Private Sub CAFPChange()

    Select Case cmbCA.ListIndex
        Case 0 'full
            Select Case cmbFP.ListIndex
                Case 0 'cash
                    txtPayAmtOnDate.Text = FormatNumber(GetTxtVal(txtTotalAmount.Text), 2)
                    txtPayAmtOnDate.Locked = True
                    
                Case 1 'check
                    txtPayAmtOnDate.Locked = True
                    
                Case 2 'other
                    txtPayAmtOnDate.Locked = False
                    
            End Select
            
            
        Case 1 'partial
            Select Case cmbFP.ListIndex
                Case 0 'cash
                    txtPayAmtOnDate.Locked = False
                Case 1 'check
                    txtPayAmtOnDate.Locked = True
                    
                Case 2 'other
                    txtPayAmtOnDate.Locked = False
            End Select
        Case 2 'not paid
            txtPayAmtOnDate.Locked = True
            txtPayAmtOnDate.Text = "0.00"
    End Select

End Sub

Private Sub txtTotalAmount_Change()
    CAFPChange
End Sub

Private Sub Form_RefreshCA(Optional sCA As String = "Full Payment")
    
    Dim i As Integer
    
    cmbCA.Clear
    
    
    cmbCA.AddItem "Full Payment"
    cmbCA.AddItem "Partial Payment"
    cmbCA.AddItem "Not Paid"
    
    For i = 0 To cmbCA.ListCount - 1
        If LCase(Trim(cmbCA.List(i))) = LCase(Trim(sCA)) Then
            cmbCA.ListIndex = i
            Exit For
        End If
    Next
    
End Sub

Private Sub Form_RefreshFP(Optional sFP As String = "Cash")
    
    cmbFP.Clear
    
    cmbFP.AddItem "Cash"
    cmbFP.AddItem "Check"
    cmbFP.AddItem "Other"

    Dim i As Integer
    

    For i = 0 To cmbFP.ListCount - 1
        If LCase(Trim(cmbFP.List(i))) = LCase(Trim(sFP)) Then
            cmbFP.ListIndex = i
            Exit Sub
        End If
    Next

    cmbFP.ListIndex = 2
End Sub

'---------------------------------------------------------------
'Customer Info Procedures
'---------------------------------------------------------------

