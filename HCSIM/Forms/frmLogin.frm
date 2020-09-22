VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00EDEBE9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User's Login"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   2970
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2970
      Width           =   1395
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1740
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
      Left            =   1200
      MaxLength       =   20
      PasswordChar    =   "="
      TabIndex        =   3
      Top             =   2280
      Width           =   3405
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   0
      Picture         =   "frmLogin.frx":058A
      Top             =   0
      Width           =   4830
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   1785
      Width           =   840
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mShowForm As Boolean
Dim dFailedCount As Integer

Public Function ShowForm() As Boolean

    'show form
    Me.Show vbModal
    
    ShowForm = mShowForm
End Function

Private Sub cmdCancel_Click()
    mShowForm = False
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    
    'check form field
    If IsEmpty(txtUserID.Text) Then
        MsgBox "Please enter User ID", vbExclamation
        HLTxt txtUserID
        Exit Sub
    End If
    
    If IsEmpty(txtPassword.Text) Then
        MsgBox "Please enter Password", vbExclamation
        HLTxt txtPassword
        Exit Sub
    End If
    
    'check user
    If GetUserByID(txtUserID.Text, CurrentUser) = False Then
        MsgBox "User does not exist.", vbExclamation
        HLTxt txtUserID
        Exit Sub
    End If
    
    'check password
    If txtPassword.Text <> CurrentUser.Password Then
    
        If dFailedCount >= 5 Then
            WriteErrorLog Me.Name, "cmdLogin_Click", "Err: 0000x000FF"
            Unload Me
            Exit Sub
        End If
    
        MsgBox "Invalid Password.", vbExclamation
        HLTxt txtPassword
        
        'increment counter
        dFailedCount = dFailedCount + 1
        
        Exit Sub
    End If
    
    
    'set current user
    If GetUserByID(Trim(txtUserID.Text), CurrentUser) = False Then
        WriteErrorLog Me.Name, "cmdLogin_Click", "GetUserByID(Trim(txtUserID.Text), CurrentUser) = False"
        Unload Me
    End If
    
    
    'success
    'write to log
    'temp
    
    'set flag
    mShowForm = True
    'close this form
    Unload Me
    
End Sub


Private Sub Form_Load()
    'default
    dFailedCount = 0
    
    txtUserID.Text = GetSetting(App.EXEName, "TextBox", txtUserID.Name, "")
    PaintGrad Me, &HEDEBE9, &HF5F5F5, 135
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    SaveSetting App.EXEName, "TextBox", txtUserID.Name, txtUserID.Text
End Sub
