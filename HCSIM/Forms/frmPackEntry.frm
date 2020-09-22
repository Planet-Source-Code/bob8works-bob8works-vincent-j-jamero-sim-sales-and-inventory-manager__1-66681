VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmPackEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Package"
   ClientHeight    =   2625
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
   Icon            =   "frmPackEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   175
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
      Top             =   2040
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4110
      TabIndex        =   1
      Top             =   2040
      Width           =   1395
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   2025
      Left            =   0
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   379
      TabIndex        =   5
      Top             =   600
      Width           =   5685
      Begin VB.TextBox txtPackTitle 
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
         MaxLength       =   20
         TabIndex        =   13
         Top             =   690
         Width           =   3855
      End
      Begin VB.TextBox txtPackID 
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
         Left            =   0
         TabIndex        =   9
         Top             =   1290
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* &Title:"
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
         TabIndex        =   14
         Top             =   720
         Width           =   570
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   4050
         Width           =   45
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
         Picture         =   "frmPackEntry.frx":000C
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Package"
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
         Width           =   1170
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
Attribute VB_Name = "frmPackEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mFormState As String

Dim curPack As tPack
Dim newPack As tPack

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isOn As Boolean

Public Function ShowAdd(Optional sPackTitle As String = "") As Boolean
    
    'set form state
    mFormState = "add"
    
    'set parameter
    newPack.PackTitle = sPackTitle
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function

Public Function ShowEdit(ByVal lPackID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curPack.PackID = lPackID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function


Private Sub cmdSave_Click()

    'Add/Edit validations
    If IsEmpty(txtPackTitle.Text) Then
        MsgBox "Please enter Title.", vbExclamation
        HLTxt txtPackTitle
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
            Me.Caption = "New Package Entry"
            
            txtPackID.Text = modFunction.ComNumZ(modRSPack.GetNewPackID, 10)
            txtPackTitle.Text = newPack.PackTitle
            
            'set focused ctl
            txtPackTitle.SetFocus
            
        Case "edit"
        
            'set form caption
            Me.Caption = "Edit Package Entry"
            
            If GetPackByID(curPack.PackID, curPack) = False Then
                WriteErrorLog Me.Name, "Form_Activate", "Failed on: 'GetPackByID(curPack.PackID, curPack) = False'"
                Unload Me
                GoTo RAE
            End If
            
            txtPackID.Text = modFunction.ComNumZ(curPack.PackID, 10)
            txtPackTitle.Text = curPack.PackTitle

            'set focused ctl
            txtPackTitle.SetFocus
            
    End Select
    
    
RAE:
    'restoremouse pointer tonormal
    Me.MousePointer = vbNormal
    Me.AutoRedraw = True
End Sub


Private Sub Form_Load()

    isOn = False
    PaintGrad bgHeader, &HEDEBE9, &HFFFFFF, 0
  
End Sub


Private Sub SaveAdd()
    
    Dim tmpPack As tPack

    'validate
    'check dupliPackion
    If GetPackByTitle(Trim(txtPackTitle.Text), tmpPack) = True Then
        MsgBox "The Package Title that you have entered was already existed.", vbExclamation
        HLTxt txtPackTitle
        Exit Sub
    End If
    
    'set new Pack info
    With newPack
        .PackID = CLng(GetTxtVal(txtPackID.Text))
        .PackTitle = Trim(txtPackTitle.Text)
    End With
    
    'write
    If modRSPack.AddPack(newPack.PackTitle) = True Then
        'set flag
        mShowAdd = True
        'close this form
        Unload Me
    Else
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSPack.AddPack(newPack) = True'"
    End If
    
End Sub


Private Sub SaveEdit()

    Dim tmpPack As tPack
        
    'validate
    'check dupliPackion
    If LCase(Trim(txtPackTitle.Text)) <> LCase(Trim(curPack.PackTitle)) Then
        If GetPackByTitle(Trim(txtPackTitle.Text), tmpPack) = True Then
            MsgBox "The Package Title that you have entered was already existed.", vbExclamation
            HLTxt txtPackTitle
            Exit Sub
        End If
    End If
    
    'set cur Pack info
    With curPack
        '.PackID = CLng(GetTxtVal(txtPackID.Text))
        .PackTitle = Trim(txtPackTitle.Text)
    End With
    
    'write
    If modRSPack.EditPack(curPack) = True Then
        'set flag
        mShowEdit = True
        'close this form
        Unload Me
    Else
        WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'modRSPack.EditPack(curPack) = True'"
    End If
    
End Sub



