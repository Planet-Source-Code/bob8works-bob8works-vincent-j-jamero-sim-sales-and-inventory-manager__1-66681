VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAboutAuthor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   397
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin b8Controls4.b8GradLine b8GradLine1 
      Height          =   885
      Left            =   30
      TabIndex        =   1
      Top             =   5010
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1561
      Color1          =   14737632
      Angle           =   90
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
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4665
      Left            =   30
      TabIndex        =   0
      Top             =   390
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   8229
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Appearance      =   0
      FileName        =   "D:\Projects\HCSIM\Code\My Message.rtf"
      TextRTF         =   $"frmAboutAuthor.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin b8Controls4.b8GradLine b8GradLine2 
      Height          =   375
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Color1          =   14737632
      Angle           =   270
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press any key to Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   6360
      TabIndex        =   3
      Top             =   60
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5895
      Left            =   3420
      Picture         =   "frmAboutAuthor.frx":03B1
      Top             =   30
      Width           =   4425
   End
End
Attribute VB_Name = "frmAboutAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API for Top Most form
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Const HWND_NOTOPMOST = -2

Public Function ShowForm()
    
    'show form
    SetWindowPos Me.hWnd, HWND_TOPMOST, _
    0, 0, 0, 0, FLAGS
    Me.Show vbModal
    
    DoEvents
    DoEvents
    DoEvents

End Function


Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or vbKeyEscape Then
        Unload Me
    End If
End Sub


