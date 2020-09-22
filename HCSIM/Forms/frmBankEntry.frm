VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmBankEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bank"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBankEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
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
      Left            =   2550
      TabIndex        =   0
      Top             =   2820
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4110
      TabIndex        =   1
      Top             =   2820
      Width           =   1395
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   2685
      Left            =   0
      ScaleHeight     =   179
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   379
      TabIndex        =   5
      Top             =   600
      Width           =   5685
      Begin VB.TextBox txtAddress 
         Height          =   675
         Left            =   1440
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1140
         Width           =   3855
      End
      Begin VB.TextBox txtBankName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   14
         Top             =   690
         Width           =   3855
      End
      Begin VB.TextBox txtBankID 
         BackColor       =   &H00F5F5F5&
         Height          =   285
         Left            =   3660
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   90
         Width           =   1635
      End
      Begin b8Controls4.b8Line b8Line1 
         Height          =   30
         Left            =   0
         TabIndex        =   7
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
         TabIndex        =   8
         Top             =   0
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line3 
         Height          =   30
         Left            =   60
         TabIndex        =   9
         Top             =   2070
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* &Name:"
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
         Left            =   270
         TabIndex        =   15
         Top             =   720
         Width           =   675
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   4050
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Address:"
         Height          =   195
         Left            =   420
         TabIndex        =   11
         Top             =   1140
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         Height          =   195
         Left            =   3360
         TabIndex        =   10
         Top             =   90
         Width           =   225
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
      ScaleWidth      =   403
      TabIndex        =   2
      Top             =   0
      Width           =   6045
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmBankEntry.frx":000C
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
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
         Width           =   705
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
Attribute VB_Name = "frmBankEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mFormState As String

Dim curBank As tBank
Dim newBank As tBank

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isOn As Boolean

Public Function ShowAdd(Optional sBankName As String = "") As Boolean
    
    'set form state
    mFormState = "add"
    
    'set parameter
    newBank.BankName = sBankName
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function

Public Function ShowEdit(ByVal lBankID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curBank.BankID = lBankID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function


Private Sub cmdSave_Click()

    'Add/Edit validations
    If IsEmpty(txtBankName.Text) Then
        MsgBox "Please enter Name.", vbExclamation
        HLTxt txtBankName
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
            Me.Caption = "New Bank Entry"
            
            txtBankID.Text = modFunction.ComNumZ(modRSBank.GetNewBankID, 10)
            txtBankName.Text = newBank.BankName
            
            'set focused ctl
            txtBankName.SetFocus
            
        Case "edit"
        
            'set form caption
            Me.Caption = "Edit Bank Entry"
            
            If GetBankByID(curBank.BankID, curBank) = False Then
                WriteErrorLog Me.Name, "Form_Activate", "Failed on: 'GetBankByID(curBank.BankID, curBank) = False'"
                Unload Me
                GoTo RAE
            End If
            
            txtBankID.Text = modFunction.ComNumZ(curBank.BankID, 10)
            txtBankName.Text = curBank.BankName
            txtAddress.Text = curBank.Address
            
            'set focused ctl
            txtBankName.SetFocus
            
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
  
End Sub


Private Sub SaveAdd()
    
    Dim tmpBank As tBank

    'validate
    'check duplication
    If GetBankByName(Trim(txtBankName.Text), tmpBank) = True Then
        MsgBox "The Bank Name that you have entered was already existed.", vbExclamation
        HLTxt txtBankName
        Exit Sub
    End If
    
    'set new Bank info
    With newBank
        .BankID = CLng(GetTxtVal(txtBankID.Text))
        .BankName = Trim(txtBankName.Text)
        .Address = Trim(txtAddress.Text)
    End With
    
    'write
    If modRSBank.AddBank(newBank.BankName, newBank.Address) = True Then
        'set flag
        mShowAdd = True
        'close this form
        Unload Me
    Else
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSBank.AddBank(newBank) = True'"
    End If
    
End Sub


Private Sub SaveEdit()

    Dim tmpBank As tBank
        
    'validate
    'check duplication
    If LCase(Trim(txtBankName.Text)) <> LCase(Trim(curBank.BankName)) Then
        If GetBankByName(Trim(txtBankName.Text), tmpBank) = True Then
            MsgBox "The Bank Name that you have entered was already existed.", vbExclamation
            HLTxt txtBankName
            Exit Sub
        End If
    End If
    
    'set cur Bank info
    With curBank
        '.BankID = CLng(GetTxtVal(txtBankID.Text))
        .BankName = Trim(txtBankName.Text)
        .Address = Trim(txtAddress.Text)
    End With
    
    'write
    If modRSBank.EditBank(curBank) = True Then
        'set flag
        mShowEdit = True
        'close this form
        Unload Me
    Else
        WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'modRSBank.EditBank(curBank) = True'"
    End If
    
End Sub



