VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPref 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HCSIM - Preferences"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPref.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ilList 
      Left            =   6450
      Top             =   390
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
            Picture         =   "frmPref.frx":1CFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin b8Controls4.b8WinTabs b8WT 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   7290
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.PictureBox bgApp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6795
      Left            =   0
      ScaleHeight     =   453
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   675
      TabIndex        =   13
      Top             =   450
      Width           =   10125
      Begin b8Controls4.b8SideTab b8STApp 
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   609
         Caption         =   "System Settings"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         Begin VB.CheckBox chkAutoBackup 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Auto Backup at startup"
            Height          =   345
            Left            =   540
            TabIndex        =   40
            Top             =   4110
            Width           =   2115
         End
         Begin b8Controls4.b8GradLine b8GradLine5 
            Height          =   30
            Left            =   0
            TabIndex        =   37
            Top             =   3930
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   53
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
         Begin b8Controls4.b8GradLine b8GradLine4 
            Height          =   30
            Left            =   0
            TabIndex        =   34
            Top             =   900
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   53
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
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "SIM settings"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00926747&
            Height          =   165
            Left            =   150
            TabIndex        =   39
            Top             =   3750
            Width           =   780
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Application Settings"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D0AD73&
            Height          =   270
            Left            =   120
            TabIndex        =   38
            Top             =   3450
            Width           =   2250
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Double click cell to modify."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00926747&
            Height          =   165
            Left            =   150
            TabIndex        =   36
            Top             =   720
            Width           =   1665
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Business Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D0AD73&
            Height          =   270
            Left            =   120
            TabIndex        =   35
            Top             =   420
            Width           =   2445
         End
      End
   End
   Begin VB.PictureBox bgMisc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6855
      ScaleWidth      =   10095
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   10095
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   330
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   609
         Caption         =   "Beginning Accounts Payable and AccountsReceivable"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         Begin b8Controls4.b8GradLine b8GradLine1 
            Height          =   30
            Left            =   0
            TabIndex        =   29
            Top             =   870
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   53
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
         Begin b8Controls4.LynxGrid3 listAR 
            Height          =   4905
            Left            =   5130
            TabIndex        =   10
            Top             =   1200
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   8652
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
            FocusRectMode   =   2
            FocusRectColor  =   33023
            AllowUserResizing=   4
            Editable        =   -1  'True
            Striped         =   -1  'True
            SBackColor1     =   16056319
            SBackColor2     =   14940667
         End
         Begin b8Controls4.LynxGrid3 listAP 
            Height          =   4905
            Left            =   90
            TabIndex        =   9
            Top             =   1200
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   8652
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
            FocusRectMode   =   2
            FocusRectColor  =   33023
            AllowUserResizing=   4
            Editable        =   -1  'True
            Striped         =   -1  'True
            SBackColor1     =   16056319
            SBackColor2     =   14940667
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Double click cell to modify."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00926747&
            Height          =   165
            Left            =   150
            TabIndex        =   28
            Top             =   690
            Width           =   1665
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Beginning Accounts Receivable:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5100
            TabIndex        =   8
            Top             =   930
            Width           =   2280
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00808080&
            Height          =   4965
            Left            =   5100
            Top             =   1170
            Width           =   4965
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Beginning Accounts Apayable:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   7
            Top             =   930
            Width           =   2175
         End
         Begin VB.Shape shpLBorder 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00808080&
            Height          =   4965
            Left            =   60
            Top             =   1170
            Width           =   4965
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Beginning Accounts Payable / Accounts Receivable"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D0AD73&
            Height          =   270
            Left            =   120
            TabIndex        =   6
            Top             =   390
            Width           =   5820
         End
      End
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   609
         Caption         =   "Categories and Packges"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         Begin b8Controls4.b8GradLine b8GradLine2 
            Height          =   30
            Left            =   0
            TabIndex        =   30
            Top             =   720
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   53
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
         Begin b8Controls4.b8Line b8Line2 
            Height          =   30
            Left            =   5130
            TabIndex        =   27
            Top             =   1110
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   53
            BorderColor1    =   15592425
            BorderColor2    =   16777215
         End
         Begin VB.CommandButton cmdPackDelete 
            Caption         =   "&Delete"
            Height          =   285
            Left            =   9330
            TabIndex        =   26
            Top             =   810
            Width           =   705
         End
         Begin VB.CommandButton cmdPackEdit 
            Caption         =   "&Edit"
            Height          =   285
            Left            =   8610
            TabIndex        =   25
            Top             =   810
            Width           =   705
         End
         Begin VB.CommandButton cmdPackAdd 
            Caption         =   "&Add"
            Height          =   285
            Left            =   7890
            TabIndex        =   24
            Top             =   810
            Width           =   705
         End
         Begin b8Controls4.b8Line b8Line1 
            Height          =   30
            Left            =   90
            TabIndex        =   23
            Top             =   1110
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   53
            BorderColor1    =   15592425
            BorderColor2    =   16777215
         End
         Begin VB.CommandButton cmdCatDelete 
            Caption         =   "&Delete"
            Height          =   285
            Left            =   4290
            TabIndex        =   22
            Top             =   810
            Width           =   705
         End
         Begin VB.CommandButton cmdCatEdit 
            Caption         =   "&Edit"
            Height          =   285
            Left            =   3570
            TabIndex        =   21
            Top             =   810
            Width           =   705
         End
         Begin VB.CommandButton cmdCatAdd 
            Caption         =   "&Add"
            Height          =   285
            Left            =   2850
            TabIndex        =   20
            Top             =   810
            Width           =   705
         End
         Begin b8Controls4.LynxGrid3 listPack 
            Height          =   4995
            Left            =   5130
            TabIndex        =   18
            Top             =   1110
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   8811
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
            FocusRectMode   =   2
            FocusRectColor  =   33023
            AllowUserResizing=   4
            Editable        =   -1  'True
            Striped         =   -1  'True
            SBackColor1     =   16056319
            SBackColor2     =   14940667
         End
         Begin b8Controls4.LynxGrid3 listCat 
            Height          =   4965
            Left            =   90
            TabIndex        =   15
            Top             =   1140
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   8758
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
            FocusRectMode   =   2
            FocusRectColor  =   33023
            AllowUserResizing=   4
            Editable        =   -1  'True
            Striped         =   -1  'True
            SBackColor1     =   16056319
            SBackColor2     =   14940667
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Packages"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5190
            TabIndex        =   19
            Top             =   840
            Width           =   675
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00F5F5F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00808080&
            Height          =   5355
            Left            =   5100
            Top             =   780
            Width           =   4965
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Categories"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   840
            Width           =   780
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00F5F5F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00808080&
            Height          =   5355
            Left            =   60
            Top             =   780
            Width           =   4965
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Categories and Packges"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D0AD73&
            Height          =   270
            Left            =   120
            TabIndex        =   16
            Top             =   420
            Width           =   2730
         End
      End
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   5745
         Index           =   2
         Left            =   0
         TabIndex        =   2
         Top             =   660
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   10134
         Caption         =   "Begining Product Inventory"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         Begin b8Controls4.b8GradLine b8GradLine3 
            Height          =   30
            Left            =   -30
            TabIndex        =   31
            Top             =   840
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   53
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
         Begin b8Controls4.LynxGrid3 listProdInv 
            Height          =   4995
            Left            =   90
            TabIndex        =   5
            Top             =   1110
            Width           =   9945
            _ExtentX        =   17542
            _ExtentY        =   8811
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
            FocusRectMode   =   2
            FocusRectColor  =   33023
            AllowUserResizing=   4
            Editable        =   -1  'True
            Striped         =   -1  'True
            SBackColor1     =   16056319
            SBackColor2     =   14940667
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Begining Product Inventory"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D0AD73&
            Height          =   270
            Left            =   90
            TabIndex        =   33
            Top             =   360
            Width           =   3090
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Double click cell to modify."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00926747&
            Height          =   165
            Left            =   120
            TabIndex        =   32
            Top             =   660
            Width           =   1665
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00808080&
            Height          =   5055
            Left            =   60
            Top             =   1080
            Width           =   10005
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Products:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   870
            Width           =   690
         End
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preferences"
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
      Left            =   150
      TabIndex        =   11
      Top             =   30
      Width           =   1665
   End
End
Attribute VB_Name = "frmPref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const iCatPackTabIndex As Integer = 0
Private Const iAPARTabIndex As Integer = 1
Private Const iProdInvTabIndex As Integer = 2

Private Const sF_App = "app"
Private Const sF_Rec = "rec"

Dim bCatLoaded As Boolean
Dim bPackLoaded As Boolean
Dim bAPLoaded As Boolean
Dim bARLoaded As Boolean
Dim bPILoaded As Boolean

Dim iCurTab As Integer
Dim iCurWin As Integer

Dim isOn As Boolean

Public Sub ShowForm(Optional iTab As Integer = -1, Optional iWin As Integer = 0)
    
    
    'defaults
    isOn = False
    
    bCatLoaded = False
    bPackLoaded = False
    bAPLoaded = False
    bARLoaded = False
    bPILoaded = False
    
    iCurTab = iTab
    iCurWin = iWin
    
    Me.Show vbModal
End Sub



Private Sub b8STApp_BeforeExpand(Index As Integer)
    b8STApp(0).MaxHeight = bgApp.Height
    RefreshAppSetting
End Sub

Private Sub b8WT_Change(sFormName As String, Index As Integer)
    Select Case Index
        Case 0
            bgMisc.Visible = False
            bgApp.Visible = True
            
        Case 1
            bgApp.Visible = False
            bgMisc.Visible = True
        
    End Select
    
    b8WT.SetForm sFormName
End Sub


Private Sub b8WT_Click(sFormName As String, Index As Integer)
    Call b8WT_Change(sFormName, Index)
End Sub

Private Sub chkAutoBackup_Click()
    modApp.SetAutoBackup IIf(chkAutoBackup.Value = vbChecked, True, False)
End Sub

Private Sub cmdCatAdd_Click()
    
    If frmCatEntry.ShowAdd() = True Then
        RefreshCat True
    End If

End Sub



Private Sub cmdCatDelete_Click()
    Dim vID As Variant
    
    If listCat.RowCount < 1 Then
        Exit Sub
    End If
    
    vID = listCat.CellText(listCat.Row, 0)
    If MsgBox("Are you sure you want to delete this Category '" & listCat.CellText(listCat.Row, 1) & "'?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
        If modRSCat.DeleteCat(CLng(vID)) = True Then
            RefreshCat True
        End If
    End If
End Sub

Private Sub cmdCatEdit_Click()
    
    Dim vID As Variant
    
    If listCat.RowCount < 1 Then
        Exit Sub
    End If
    
    vID = listCat.CellText(listCat.Row, 0)
    
    If frmCatEntry.ShowEdit(CLng(vID)) = True Then
        RefreshCat True
    End If
    
End Sub











Private Sub cmdPackAdd_Click()
    
    If frmPackEntry.ShowAdd() = True Then
        RefreshPack True
    End If

End Sub



Private Sub cmdPackDelete_Click()
    Dim vID As Variant
    
    If listPack.RowCount < 1 Then
        Exit Sub
    End If
    
    vID = listPack.CellText(listPack.Row, 0)
    If MsgBox("Are you sure you want to delete this Package '" & listPack.CellText(listPack.Row, 1) & "'?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
        If modRSPack.DeletePack(CLng(vID)) = True Then
            RefreshPack True
        End If
    End If
End Sub

Private Sub cmdPackEdit_Click()
    
    Dim vID As Variant
    
    If listPack.RowCount < 1 Then
        Exit Sub
    End If
    
    vID = listPack.CellText(listPack.Row, 0)
    
    If frmPackEntry.ShowEdit(CLng(vID)) = True Then
        RefreshPack True
    End If
    
End Sub











Private Sub Form_Activate()
    
    If isOn = True Then
        Exit Sub
    End If
    isOn = True
    
    DoEvents: DoEvents: DoEvents
    
    Select Case iCurWin
    
        Case 0 'app
        
            b8WT.SetForm sF_App
            RefreshAppSetting
            
        Case 1 'Records
        
            b8WT.SetForm sF_Rec
            
            If iCurTab >= 0 Then
                b8ST(iCurTab).Expanded = True
            End If
    End Select

End Sub

Private Sub Form_Load()


    b8WT.AddForm sF_App, "System Settings"
    b8WT.AddForm sF_Rec, "Misc Records"
    b8WT.SetForm sF_App
    
    'set colums
    With listCat
        .AddColumn "ID", 80, lgAlignRightCenter
        .AddColumn "Title", 160
        .AddColumn "Description", 66
        .RowHeightMin = 21
        .ImageList = ilList
    End With
    
    With listPack
        .AddColumn "ID", 80, lgAlignRightCenter
        .AddColumn "Title", 226
        .RowHeightMin = 21
        .ImageList = ilList
    End With
    
    With listAP
        .AddColumn "Supplier", 218
        .AddColumn "Supplier ID", 0
        .AddColumn "A/P", 90, lgAlignRightCenter
        .RowHeightMin = 21
        .ImageList = ilList
    End With
    
    With listAR
        .AddColumn "Customer", 218
        .AddColumn "Customer ID", 0
        .AddColumn "A/R", 90, lgAlignRightCenter
        .RowHeightMin = 21
        .ImageList = ilList
    End With
    
    With listProdInv
        .AddColumn "Description", 218
        .AddColumn "Prod ID", 0
        .AddColumn "Code", 130
        .AddColumn "Package", 100
        .AddColumn "Category", 100
        .AddColumn "Unit/s", 90, lgAlignRightCenter
        .RowHeightMin = 21
        .ImageList = ilList
    End With
    
    'arrange tabs
    Dim i As Integer
    For i = 0 To b8ST.UBound
            If b8ST(i).AutoContract = True Then
                b8ST(i).Expanded = False
            End If
    Next
    
End Sub



Private Sub b8ST_BeforeExpand(Index As Integer)
    
    Dim i As Integer
    Dim tHeight As Integer
    
    'load data
    Select Case Index
        Case iCatPackTabIndex
            RefreshCat
            RefreshPack
        Case iAPARTabIndex
            RefreshAP
            RefreshAR
        Case iProdInvTabIndex
            RefreshProdInv
            
    End Select
    
    tHeight = (b8ST.UBound + 1) * 345
    b8ST(Index).MaxHeight = bgMisc.Height - tHeight
    For i = 0 To b8ST.UBound
        If Index <> i Then
            If b8ST(i).AutoContract = True Then
                b8ST(i).Expanded = False
            End If
        End If
    Next
    
    
End Sub

Private Sub b8ST_Resize(Index As Integer)
    
    Dim i As Integer
    
    For i = 1 To b8ST.UBound
        b8ST(i).Move b8ST(i).Left, (b8ST(i - 1).Top + b8ST(i - 1).Height) - 15
    Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bCatLoaded = False
    bPackLoaded = False
    bAPLoaded = False
    bARLoaded = False
    bPILoaded = False
End Sub

Private Sub listAP_DblClick()
    
    Dim lSupID As Long
    
    If listAP.RowCount < 1 Then
        Exit Sub
    End If
    
    If listAP.Col = 2 Then
        listAP.EditCell listAP.Row, listAP.Col
        Exit Sub
    End If
    
    lSupID = GetTxtVal(listAP.CellText(listAP.Row, 1))
    
    If frmSupEntry.ShowEdit(lSupID) = True Then
        RefreshAP True
    End If
    
End Sub

Private Sub listAP_RequestEdit(Row As Long, Col As Long, Cancel As Boolean)
    If Col <> 2 Then
        Cancel = True
    End If
End Sub


Private Sub listAP_RequestUpdate(Row As Long, Col As Long, NewValue As String, Cancel As Boolean)
    
    'default
    Cancel = True
    
    'validate
    If GetTxtVal(NewValue) < 0 Then
        MsgBox "Please enter valid value. It must be greater or equalto 0", vbExclamation
        Exit Sub
    End If
    
    NewValue = FormatNumber(GetTxtVal(NewValue), 2)
    
    If modRSSup.SetSupBegAP(GetTxtVal(listAP.CellText(Row, 1)), FormatNumber(GetTxtVal(NewValue), 2)) = True Then
        listAP.CellText(Row, Col) = FormatNumber(GetTxtVal(NewValue), 2)
        listAP.Refresh
    Else
        WriteErrorLog Me.Name, "listAP_RequestUpdate", "Failed on: 'modRSSup.SetSupBegAP(GetTxtVal(listAP.celltext(Row, 1)), GetTxtVal(NewValue)) = True'"
        'refresh list
        RefreshAP True
    End If
    
    Cancel = False
    
End Sub

Private Sub listAR_DblClick()
    Dim lCustID As Long
    
    If listAR.RowCount < 1 Then
        Exit Sub
    End If
    
    If listAR.Col = 2 Then
        listAR.EditCell listAR.Row, listAR.Col
        Exit Sub
    End If
    
    
    lCustID = GetTxtVal(listAR.CellText(listAR.Row, 1))
    
    If frmCustEntry.ShowEdit(lCustID) = True Then
        RefreshAR True
    End If
End Sub

Private Sub listAR_RequestEdit(Row As Long, Col As Long, Cancel As Boolean)
    If Col <> 2 Then
        Cancel = True
    End If
End Sub

Private Sub listAR_RequestUpdate(Row As Long, Col As Long, NewValue As String, Cancel As Boolean)
    
    'default
    Cancel = True
    
    'validate
    If GetTxtVal(NewValue) < 0 Then
        MsgBox "Please enter valid value. It must be greater or equalto 0", vbExclamation
        Exit Sub
    End If
    
    NewValue = FormatNumber(GetTxtVal(NewValue), 2)
    
    If modRSCust.SetCustBegAR(GetTxtVal(listAR.CellText(Row, 1)), FormatNumber(GetTxtVal(NewValue), 2)) = True Then
        listAR.CellText(Row, Col) = FormatNumber(GetTxtVal(NewValue), 2)
        listAR.Refresh
    Else
        WriteErrorLog Me.Name, "listAR_RequestUpdate", "Failed on: 'modRSSup.SetSupBegAP(GetTxtVal(listAR.celltext(Row, 1)), GetTxtVal(NewValue)) = True'"
        'refresh list
        RefreshAR True
    End If
    
    Cancel = False
End Sub




Private Sub listCat_DblClick()
    Call cmdCatEdit_Click
End Sub

Private Sub listCat_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF2
            Call cmdCatAdd_Click
        Case vbKeyF3
            Call cmdCatEdit_Click
        Case vbKeyDelete
            Call cmdCatDelete_Click
                    
    End Select
    
End Sub

Private Sub listPack_DblClick()
    Call cmdPackEdit_Click
End Sub

Private Sub listPack_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF2
            Call cmdPackAdd_Click
        Case vbKeyF3
            Call cmdPackEdit_Click
        Case vbKeyDelete
            Call cmdPackDelete_Click
                    
    End Select
End Sub

Private Sub listProdInv_DblClick()
    Dim lProdID As Long
    
    If listProdInv.RowCount < 1 Then
        Exit Sub
    End If
    
    If listProdInv.Col = 5 Then
        listProdInv.EditCell listProdInv.Row, listProdInv.Col
        Exit Sub
    End If
    
    lProdID = GetTxtVal(listProdInv.CellText(listProdInv.Row, 1))
    
    If frmProdEntry.ShowEdit(lProdID) = True Then
        RefreshProdInv True
    End If
End Sub

Private Sub listProdInv_RequestEdit(Row As Long, Col As Long, Cancel As Boolean)
    If Col <> 5 Then
        Cancel = True
    End If
End Sub

Private Sub listProdInv_RequestUpdate(Row As Long, Col As Long, NewValue As String, Cancel As Boolean)
    'default
    Cancel = True
    
    'validate
    If GetTxtVal(NewValue) < 0 Then
        MsgBox "Please enter valid value. It must be greater or equalto 0", vbExclamation
        Exit Sub
    End If
    
    If modRSProd.SetProdBegInvStock(GetTxtVal(listProdInv.CellText(Row, 1)), GetTxtVal(NewValue)) = True Then
        listProdInv.CellText(Row, Col) = GetTxtVal(NewValue)
        listProdInv.Refresh
    Else
        WriteErrorLog Me.Name, "listProdInv_RequestUpdate", "Failed on: 'modRSSup.SetSupBegAP(GetTxtVal(listProdInv.celltext(Row, 1)), GetTxtVal(NewValue)) = True'"
        'refresh list
        RefreshProdInv True
    End If
    
    Cancel = False
End Sub





Private Sub RefreshAppSetting()

    chkAutoBackup.Value = IIf(modApp.GetAutoBackup = True, vbChecked, vbUnchecked)
End Sub







Private Sub RefreshCat(Optional ByVal bForceLoad As Boolean = False)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long

    'exit if this procedure was already execute
    If bCatLoaded = True And bForceLoad = False Then
        GoTo RAE
    End If
    bCatLoaded = True

    listCat.Redraw = False
    listCat.Clear
    
    
    sSQL = "SELECT tblCat.CatID, tblCat.CatTitle, tblCat.Description" & _
            " From tblCat" & _
            " ORDER BY tblCat.CatTitle"


    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "RefreshCat", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    vRS.MoveFirst
    While vRS.EOF = False
        With listCat
            li = .AddItem(ReadField(vRS.Fields("CatID")))
            .CellText(li, 1) = ReadField(vRS.Fields("CatTitle"))
            .CellText(li, 2) = ReadField(vRS.Fields("Description"))
            .ItemImage(li) = 1
        End With
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    listCat.Redraw = True
    listCat.Refresh
    
End Sub


Private Sub RefreshPack(Optional bForceLoad As Boolean = False)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long
    
    'exit if this procedure was already execute
    'temp 101
'    If bPackLoaded = True And bForceLoad = False Then
'        GoTo RAE
'    End If
'    bPackLoaded = True

    listPack.Redraw = False
    listPack.Clear
    
    
    sSQL = "SELECT tblPack.PackID, tblPack.PackTitle" & _
            " From tblPack" & _
            " ORDER BY tblPack.PackTitle"



    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "RefreshPack", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    vRS.MoveFirst
    While vRS.EOF = False
        With listPack
            li = .AddItem(ReadField(vRS.Fields("PackID")))
            .CellText(li, 1) = ReadField(vRS.Fields("PackTitle"))
            .ItemImage(li) = 1
        End With
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    listPack.Redraw = True
    listPack.Refresh
    
End Sub



Private Sub RefreshAP(Optional bForceLoad As Boolean = False)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long
    
    'exit if this procedure was already execute
    If bAPLoaded = True And bForceLoad = False Then
        GoTo RAE
    End If
    bPackLoaded = True

    listAP.Redraw = False
    listAP.Clear
    
    
    sSQL = "SELECT tblSup.SupName, tblSup.SupID, tblSup.BegAP" & _
            " From tblSup" & _
            " ORDER BY tblSup.SupName"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "RefreshAP", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    vRS.MoveFirst
    While vRS.EOF = False
        With listAP
            li = .AddItem(ReadField(vRS.Fields("SupName")))
            .CellText(li, 1) = ReadField(vRS.Fields("SupID"))
            .CellText(li, 2) = FormatNumber(ReadField(vRS.Fields("BegAP")), 2)
            .ItemImage(li) = 1
        End With
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    listAP.Redraw = True
    listAP.Refresh
    
End Sub

Private Sub RefreshAR(Optional bForceLoad As Boolean = False)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long
    
    'exit if this procedure was already execute
    
    If bARLoaded = True And bForceLoad = False Then
        GoTo RAE
    End If
    bARLoaded = True

    listAR.Redraw = False
    listAR.Clear
    
    
    sSQL = "SELECT tblCust.CustName, tblCust.CustID, tblCust.BegAR" & _
            " From tblCust" & _
            " ORDER BY tblCust.CustName"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "RefreshAR", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    vRS.MoveFirst
    While vRS.EOF = False
        With listAR
            li = .AddItem(ReadField(vRS.Fields("CustName")))
            .CellText(li, 1) = ReadField(vRS.Fields("CustID"))
            .CellText(li, 2) = FormatNumber(ReadField(vRS.Fields("BegAR")), 2)
            .ItemImage(li) = 1
        End With
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    listAR.Redraw = True
    listAR.Refresh
    
End Sub



Private Sub RefreshProdInv(Optional bForceLoad As Boolean = False)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long
    
    'exit if this procedure was already execute
    If bPILoaded = True And bForceLoad = False Then
        GoTo RAE
    End If
    bPILoaded = True

    listProdInv.Redraw = False
    listProdInv.Clear
    
    
    sSQL = "SELECT tblProd.ProdID, tblProd.ProdCode, tblProd.ProdDescription, tblPack.PackTitle, tblCat.CatTitle, tblProd.BegInvStock" & _
            " FROM tblCat INNER JOIN (tblPack INNER JOIN tblProd ON tblPack.PackID = tblProd.FK_PackID) ON tblCat.CatID = tblProd.FK_CatID" & _
            " ORDER BY tblProd.ProdDescription"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "RefreshProdInv", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    vRS.MoveFirst
    While vRS.EOF = False
        With listProdInv
            li = .AddItem(ReadField(vRS.Fields("ProdDescription")))
            .CellText(li, 1) = ReadField(vRS.Fields("ProdID"))
            .CellText(li, 2) = ReadField(vRS.Fields("ProdCode"))
            .CellText(li, 3) = ReadField(vRS.Fields("PackTitle"))
            .CellText(li, 4) = ReadField(vRS.Fields("CatTitle"))
            .CellText(li, 5) = FormatNumber(ReadField(vRS.Fields("BegInvStock")), 0)
            .ItemImage(li) = 1
        End With
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    listProdInv.Redraw = True
    listProdInv.Refresh
    
End Sub

