VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "Welcome"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image4 
      Height          =   150
      Left            =   4380
      Picture         =   "frmWelcome.frx":0000
      Top             =   4680
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   4350
      Picture         =   "frmWelcome.frx":083A
      Top             =   4410
      Width           =   3795
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   4620
      Picture         =   "frmWelcome.frx":3DEC
      Stretch         =   -1  'True
      Top             =   4140
      Width           =   15360
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   12000
      Left            =   0
      Top             =   4380
      Width           =   15360
   End
   Begin VB.Image Image1 
      Height          =   3465
      Left            =   0
      Picture         =   "frmWelcome.frx":3E7E
      Top             =   990
      Width           =   4635
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowForm()


    mdiMain.AddChild Me
    
End Sub

Private Sub Form_Activate()
    mdiMain.ActivateChild Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'pass keyinfo to mdiMain
    mdiMain.AFForm_KeyDown KeyCode, Shift
End Sub
