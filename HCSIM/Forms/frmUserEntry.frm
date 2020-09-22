VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmUserEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loading..."
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2130
      Width           =   1395
   End
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
      Left            =   1920
      TabIndex        =   0
      Top             =   2130
      Width           =   1395
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   2085
      Left            =   0
      ScaleHeight     =   139
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   5
      Top             =   600
      Width           =   5055
      Begin VB.TextBox txtUserID 
         Height          =   315
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   11
         Top             =   270
         Width           =   3405
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1365
         MaxLength       =   20
         PasswordChar    =   "="
         TabIndex        =   10
         Top             =   810
         Width           =   3405
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   30
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   1380
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "* &User Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "* &Password:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   855
         Width           =   885
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   3870
         Width           =   45
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
      ScaleWidth      =   339
      TabIndex        =   2
      Top             =   0
      Width           =   5085
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmUserEntry.frx":000C
         Top             =   30
         Width           =   480
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
         TabIndex        =   4
         Top             =   420
         Width           =   3900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
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
         TabIndex        =   3
         Top             =   60
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmUserEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mFormState As String

Dim curUser As tUser

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Public Function ShowAdd() As Boolean
    
    'check current user
    If LCase(CurrentUser.UserID) <> "administrator" Then
        MsgBox "You are not permitted to access user entries.", vbExclamation
        Unload Me
        Exit Function
    End If
    
    'set form state
    mFormState = "add"
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function

Public Function ShowAddAdmin() As Boolean
    
    'set form state
    mFormState = "addadmin"
    
    txtUserID.Text = "Administrator"
    txtUserID.Enabled = False
    'show form
    Me.Show vbModal
    
    'return
    ShowAddAdmin = mShowAdd
    
End Function

Public Function ShowEdit(sUserID As String) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curUser.UserID = sUserID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function



Private Sub cmdCancel_Click()
    
    Select Case mFormState
        Case "add"
            mShowAdd = False
        Case "addadmin"
            mShowAdd = False
        Case "edit"
            mShowEdit = False
    End Select
    
    Unload Me
End Sub

Private Sub cmdSave_Click()

    Select Case mFormState
        Case "add"
            SaveAdd
        Case "addadmin"
            SaveAdd
        Case "edit"
            SaveEdit
    End Select
    
End Sub

Private Function SaveEdit()

    Dim newUser As tUser
    Dim oldUser As tUser
    
    'check form field
    If IsEmpty(txtUserID.Text) Then
        MsgBox "Please enter User ID", vbExclamation
        HLTxt txtUserID
        Exit Function
    End If
    
    If IsEmpty(txtPassword.Text) Then
        MsgBox "Please enter Password", vbExclamation
        HLTxt txtPassword
        Exit Function
    End If
        
    'check duplication
    If LCase(curUser.UserID) <> LCase(txtUserID.Text) Then
        If GetUserByID(txtUserID.Text, oldUser) = True Then
            MsgBox "The User ID that you have entered is already exist." & vbNewLine & vbNewLine & _
                "Please enter different value.", vbExclamation
            
            HLTxt txtUserID
            Exit Function
        End If
    End If
    
    'set new user
    curUser.UserID = txtUserID.Text
    curUser.Password = txtPassword.Text
    
    'try
    'add new user
    If EditUser(curUser) = True Then
        MsgBox "User entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update User entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function

Private Function SaveAdd()
    Dim newUser As tUser
    Dim oldUser As tUser
    
    
    'check form field
    If IsEmpty(txtUserID.Text) Then
        MsgBox "Please enter User ID", vbExclamation
        HLTxt txtUserID
        Exit Function
    End If
    
    If mFormState = "add" Then
        If LCase(txtUserID.Text) = "administrator" Then
            MsgBox "User ID cannot be 'Administrator'.", vbExclamation
            HLTxt txtUserID
            Exit Function
        End If
    End If
    
    If IsEmpty(txtPassword.Text) Then
        MsgBox "Please enter Password", vbExclamation
        HLTxt txtPassword
        Exit Function
    End If
        
    'check duplication
    If GetUserByID(txtUserID.Text, oldUser) = True Then
        MsgBox "The User ID that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtUserID
        Exit Function
    End If
    
    'set new user
    newUser.UserID = txtUserID
    newUser.Password = txtPassword
    
    'try
    'add new user
    If AddUser(newUser) = True Then
        MsgBox "New User entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to add new User entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function



Private Sub Form_Activate()
    
    Select Case mFormState
        Case "add"
        
            'set caption
            Me.Caption = "Add User"
            Me.cmdSave.Caption = "&Save"
            
        Case "edit"
            'get info
            If GetUserByID(curUser.UserID, curUser) = False Then
                'show failed
                MsgBox "User entry with User ID : '" & curUser.UserID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
            txtUserID.Text = curUser.UserID
            txtPassword.Text = curUser.Password
            
            'set caption
            Me.Caption = "Edit User"
            Me.cmdSave.Caption = "&Update"
            
            txtUserID.Enabled = False
            
    End Select
    
End Sub

Private Sub Form_Load()
    'PaintGrad Me, &H8000000F, &H80000014, 135
End Sub
