VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMSG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Author's Message"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9300
   Icon            =   "frmMSG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   7290
      TabIndex        =   1
      Top             =   6300
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6135
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   10821
      _Version        =   393217
      ScrollBars      =   2
      FileName        =   "D:\Projects\HCSIM\Code\Document.rtf"
      TextRTF         =   $"frmMSG.frx":000C
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ShowForm()
    Me.Show vbModal
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
