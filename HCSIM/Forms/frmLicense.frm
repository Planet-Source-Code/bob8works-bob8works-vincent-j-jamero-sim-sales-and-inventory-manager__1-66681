VERSION 5.00
Begin VB.Form frmLicense 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DMIS Key"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   750
      Width           =   1035
   End
   Begin VB.TextBox txtKey 
      Height          =   345
      Left            =   150
      TabIndex        =   0
      Top             =   330
      Width           =   3555
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter Key:"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   1365
   End
End
Attribute VB_Name = "frmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mKey As String

Public Function ShowForm() As String
    
    Me.Show vbModal
    
    ShowForm = mKey
End Function


Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mKey = txtKey.Text
End Sub
