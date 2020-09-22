VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00F5F5F5&
   Caption         =   "Sales and Inventory Manager"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer timeUpdateDate 
      Interval        =   1000
      Left            =   5040
      Top             =   1860
   End
   Begin b8Controls4.b8ClientWin b8CW 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   6150
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   661
      Begin VB.PictureBox bgSystemBot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   0
         ScaleHeight     =   405
         ScaleWidth      =   3495
         TabIndex        =   33
         Top             =   0
         Width           =   3495
         Begin VB.CommandButton cmdVote 
            Caption         =   "V O T E"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   34
            Top             =   30
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Please don't forget to"
            Height          =   285
            Left            =   120
            TabIndex        =   35
            Top             =   60
            Width           =   1605
         End
         Begin VB.Image Image3 
            Height          =   360
            Left            =   0
            Picture         =   "mdiMain.frx":1CFA
            Stretch         =   -1  'True
            Top             =   0
            Width           =   19995
         End
      End
   End
   Begin b8Controls4.b8SBCenter b8SBC 
      Align           =   3  'Align Left
      Height          =   5070
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   8943
      MinWidth        =   180
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   510
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         Caption         =   "Quick Launch         [ Ctrl + Q ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         MaxHeight       =   3285
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         Begin MSComctlLib.ImageList ilQL 
            Left            =   1080
            Top             =   660
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
                  Picture         =   "mdiMain.frx":1DD0
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView listQL 
            Height          =   3075
            Left            =   30
            TabIndex        =   32
            Top             =   360
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   5424
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDropMode     =   1
            _Version        =   393217
            Icons           =   "ilQL"
            SmallIcons      =   "ilQL"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDropMode     =   1
            NumItems        =   0
         End
      End
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   1
         Left            =   60
         TabIndex        =   3
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         Caption         =   "Search Item          [ Ctrl + S ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         MaxHeight       =   2245
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         AutoContract    =   0   'False
         Begin VB.TextBox txtSearchWhat 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   6
            Top             =   690
            Width           =   3165
         End
         Begin VB.ComboBox cmbLookIn 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1470
            Width           =   3135
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   4
            Top             =   1950
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Search What:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   450
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Look In:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   7
            Top             =   1260
            Width           =   585
         End
      End
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   2
         Left            =   60
         TabIndex        =   11
         Top             =   1170
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         Caption         =   "Filter By Date        [ Ctrl + D ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         MaxHeight       =   2865
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         AutoContract    =   0   'False
         Begin b8Controls4.b8DatePicker b8DateP 
            Height          =   2415
            Left            =   120
            TabIndex        =   12
            Top             =   420
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   4260
            BackColor       =   16777215
            MinDate         =   38968
            MaxDate         =   38968
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today is "
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
         Left            =   150
         TabIndex        =   31
         Top             =   270
         Width           =   585
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   780
         TabIndex        =   30
         Top             =   255
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome,"
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
         Left            =   150
         TabIndex        =   29
         Top             =   60
         Width           =   600
      End
      Begin VB.Label lblCurrentUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   28
         Top             =   45
         Width           =   180
      End
   End
   Begin VB.PictureBox bgHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   0
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   0
      Top             =   0
      Width           =   7515
      Begin VB.PictureBox bgRecOpt 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   3600
         ScaleHeight     =   51
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1024
         TabIndex        =   18
         Top             =   330
         Width           =   15360
         Begin b8Controls4.b8ToolButton b8RecOpt 
            Height          =   615
            Index           =   0
            Left            =   30
            TabIndex        =   19
            Top             =   60
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":1EE3
            BackColor       =   -2147483643
            Caption         =   "New"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   0
            DisabledPicture =   "mdiMain.frx":27BD
            BgColorDown     =   12632256
            BorderColor     =   12632256
         End
         Begin b8Controls4.b8ToolButton b8RecOpt 
            Height          =   615
            Index           =   1
            Left            =   1050
            TabIndex        =   20
            Top             =   60
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":3097
            BackColor       =   -2147483643
            Caption         =   "Edit"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   4210752
            DisabledPicture =   "mdiMain.frx":3971
            BorderColor     =   12632256
         End
         Begin b8Controls4.b8ToolButton b8RecOpt 
            Height          =   615
            Index           =   2
            Left            =   2070
            TabIndex        =   21
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":424B
            BackColor       =   -2147483643
            Caption         =   "Delete"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   4210752
            DisabledPicture =   "mdiMain.frx":4B25
            BorderColor     =   12632256
         End
         Begin b8Controls4.b8ToolButton b8RecOpt 
            Height          =   615
            Index           =   3
            Left            =   3270
            TabIndex        =   22
            Top             =   60
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":53FF
            BackColor       =   -2147483643
            Caption         =   "Refresh"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   0
            DisabledPicture =   "mdiMain.frx":5CD9
            BorderColor     =   12632256
         End
         Begin b8Controls4.b8ToolButton b8RecOpt 
            Height          =   615
            Index           =   4
            Left            =   4560
            TabIndex        =   23
            Top             =   60
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":65B3
            BackColor       =   -2147483643
            Caption         =   "Print"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   0
            DisabledPicture =   "mdiMain.frx":6E8D
            BorderColor     =   12632256
         End
         Begin b8Controls4.b8Line b8Line2 
            Height          =   30
            Left            =   0
            TabIndex        =   26
            Top             =   720
            Width           =   15720
            _ExtentX        =   27728
            _ExtentY        =   53
            BorderColor1    =   14737632
            BorderColor2    =   16777215
         End
      End
      Begin VB.PictureBox bgHeaderMenu 
         BackColor       =   &H00EDEBE9&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   0
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1024
         TabIndex        =   13
         Top             =   0
         Width           =   15360
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   15
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&System"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&System"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   1
            Left            =   810
            TabIndex        =   15
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Records"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Records"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   2
            Left            =   1710
            TabIndex        =   16
            Top             =   15
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Monitoring"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Monitoring"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   3
            Left            =   2670
            TabIndex        =   24
            Top             =   15
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Tools"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Tools"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   4
            Left            =   3300
            TabIndex        =   25
            Top             =   15
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Help"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Help"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Line b8Line1 
            Height          =   30
            Left            =   0
            TabIndex        =   17
            Top             =   300
            Width           =   15720
            _ExtentX        =   27728
            _ExtentY        =   53
            BorderColor1    =   16119285
            BorderColor2    =   14737632
         End
      End
      Begin b8Controls4.b8Line b8LLogoB 
         Height          =   30
         Left            =   0
         TabIndex        =   27
         Top             =   1050
         Visible         =   0   'False
         Width           =   15720
         _ExtentX        =   27728
         _ExtentY        =   53
         BorderColor1    =   14737632
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8SBtop b8SBT 
         Height          =   945
         Left            =   0
         TabIndex        =   1
         Top             =   330
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1667
         MinWidth        =   180
         Begin VB.Image Image1 
            Height          =   480
            Left            =   150
            Picture         =   "mdiMain.frx":7767
            Top             =   150
            Width           =   480
         End
         Begin VB.Image Image2 
            Height          =   540
            Left            =   690
            Picture         =   "mdiMain.frx":8031
            Top             =   150
            Width           =   1710
         End
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "&System"
      Visible         =   0   'False
      Begin VB.Menu mnuAddUser 
         Caption         =   "&Add New User"
      End
      Begin VB.Menu mnuManageUser 
         Caption         =   "&Manage Users"
      End
      Begin VB.Menu mnuS01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "&Log Off"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "&Records"
      Visible         =   0   'False
      Begin VB.Menu mnuAddPO 
         Caption         =   "&New P.O. Entry"
      End
      Begin VB.Menu mnuAddSales 
         Caption         =   "New &Sales Entry"
      End
      Begin VB.Menu mnus09 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProducts 
         Caption         =   "&Products"
         Begin VB.Menu mnuAddProduct 
            Caption         =   "&Add New Product"
         End
         Begin VB.Menu mnuManageProducts 
            Caption         =   "&Manage Products"
         End
         Begin VB.Menu mnuS02 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCategories 
            Caption         =   "Categories"
            Begin VB.Menu mnuAddCateroy 
               Caption         =   "&Add New Category Entry"
            End
            Begin VB.Menu mnuManageCategories 
               Caption         =   "&Manage Categories"
            End
         End
         Begin VB.Menu mnuPackages 
            Caption         =   "&Packages"
            Begin VB.Menu mnuAddPackage 
               Caption         =   "&Add New Package Entry"
            End
            Begin VB.Menu mnuManagePackages 
               Caption         =   "&Manage Packages"
            End
         End
      End
      Begin VB.Menu mnuSupplier 
         Caption         =   "Supplier Entries"
         Begin VB.Menu mnuAddSupplier 
            Caption         =   "&Add New Supplier Entry"
         End
         Begin VB.Menu mnuManageSupplier 
            Caption         =   "&Manage Supplier Entries"
         End
      End
      Begin VB.Menu mnuCustomers 
         Caption         =   "&Customers"
         Begin VB.Menu mnuAddCustomer 
            Caption         =   "&Add New Customer Entry"
         End
         Begin VB.Menu mnuManageCustomer 
            Caption         =   "&Manage Customers"
         End
      End
      Begin VB.Menu mnuS05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBank 
         Caption         =   "&Bank Entries"
         Begin VB.Menu mnuAddBank 
            Caption         =   "&New Bank Entry"
         End
         Begin VB.Menu mnuManageBank 
            Caption         =   "&Manage Bank Entries"
         End
      End
   End
   Begin VB.Menu mnuMonitoring 
      Caption         =   "&Monitoring"
      Visible         =   0   'False
      Begin VB.Menu mnuPPM 
         Caption         =   "&Purchase/Payment Monitoring"
      End
      Begin VB.Menu mnuSICPM 
         Caption         =   "&Sales/Customer Payments Monitoring"
      End
      Begin VB.Menu mnuAllVoid 
         Caption         =   "&Void Products Monitoring"
      End
      Begin VB.Menu mnuStockInvMon 
         Caption         =   "Stock &Inventory Monitoring"
      End
      Begin VB.Menu mnuManagePTSDueCheck 
         Caption         =   "Manage Due Checks (Payment to Supplier)"
      End
      Begin VB.Menu mnuManageCustPayDueCheck 
         Caption         =   "Manage Due Checks (Customer Payments)"
      End
      Begin VB.Menu mnuS04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuManageStockInv 
         Caption         =   "&Manage Beg. Stock Inv."
      End
      Begin VB.Menu mnuManageARAP 
         Caption         =   "&Manage Beg.  AP / AR"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuPreferences 
         Caption         =   "&Preferences"
      End
      Begin VB.Menu mnuDatabaseUtilities 
         Caption         =   "&Database Utilities"
         Begin VB.Menu mnuBackupDatabase 
            Caption         =   "&Backup Database"
         End
         Begin VB.Menu mnuDatabaseRestore 
            Caption         =   "Database Restore"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuS03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutApp 
         Caption         =   "About &SIM"
      End
      Begin VB.Menu mnuAboutAuthor 
         Caption         =   "&About The Author"
      End
      Begin VB.Menu mnuvisithome 
         Caption         =   "&Visit www.bob8works.cjb.net"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_TabShowQuickLaunch = 0
Private Const m_TabSearch = 1
Private Const m_TabFilterDate = 2

'Flag for User log
Public bUserLoggedOn As Boolean


Public Function ShowForm()
    
    'default
    bUserLoggedOn = False
    
    'show form
    Me.WindowState = vbMaximized
    Me.Show
    DoEvents
    
    'show weclome
    frmWelcome.ShowForm
    
    'unload splash
    frmSplash.UnloadSplash
    
BeginLogin:
    If AnyUserExist = False Then
        If frmUserEntry.ShowAddAdmin = False Then
            Unload Me
        Else
            GoTo BeginLogin
        End If
    Else
        If frmLogin.ShowForm = False Then
            Unload Me
            Exit Function
        End If
    End If
    
    'set log flag
    bUserLoggedOn = True
    
      
    'set UI
    'current user info
    lblCurrentUser.Caption = CurrentUser.UserID
    
    'set date
    timeUpdateDate_Timer
    
    
    'if the current user is not the administrator,
    'disable user related menus
    If LCase(Trim(CurrentUser.UserID)) <> "administrator" Then
        mnuAddUser.Enabled = False
        mnuManageUser.Enabled = False
    Else
        mnuAddUser.Enabled = True
        mnuManageUser.Enabled = True
    End If
    

End Function



'Control Procedures
'-----------------------------------------------------------
Private Sub b8CW_FormTabClick(ByVal sFormName As String, ByVal Index As Integer)
    modFuncChild.ActivateMDIChildForm sFormName
End Sub

Private Sub b8DateP_Change()
    Call Form_DateChange
End Sub



Private Sub b8RecOpt_Click(Index As Integer)
    Select Case Index
        Case 0 'add
            Form_Add
        Case 1 'edit
            Form_Edit
        Case 2 'delete
            Form_Delete
        Case 3 'refresh
            Form_Refresh
        Case 4 'print
            Form_Print
    End Select
End Sub



Private Sub b8SBC_BeforeResize(ByVal NewWidth As Integer)
    ResizeFb8SBC NewWidth
End Sub

Private Sub ResizeFb8SBC(ByVal NewWidth As Integer)
    
    'resize top side bar
    b8SBT.Width = NewWidth / Screen.TwipsPerPixelX
    bgSystemBot.Width = NewWidth
    'resize quick tabs
    Dim i As Integer
    For i = 0 To b8ST.UBound
        b8ST(i).Left = 60
        b8ST(i).Width = NewWidth - 120
    Next
    
    'resize window tab
    If b8SBC.Visible = True Then
        b8CW.SBWidth = NewWidth / Screen.TwipsPerPixelX
    Else
        b8CW.SBWidth = 0
    End If
    
    'call mdi resize to resize all opened mdi childs
    MDIForm_Resize
    
End Sub

Private Sub b8SBC_Resize()
    ResizeFb8SBC b8SBC.Width
End Sub

Private Sub b8SBT_Resize()
    b8SBC.Width = b8SBT.Width * Screen.TwipsPerPixelX
End Sub

Private Sub b8SBT_SizeChange(ByVal newSizeState As b8Controls4.eSizeState)
    
    If newSizeState = ssContracted Then
        b8CW.SBWidth = b8SBC.Width / Screen.TwipsPerPixelX
        b8SBC.Visible = True
        bgSystemBot.Visible = True
        b8LLogoB.Visible = False
    Else
        b8CW.SBWidth = 0
        b8SBC.Visible = False
        bgSystemBot.Visible = False
        b8LLogoB.Visible = True
    End If
    
    'call mdi resize to resize all opened child forms
    Call MDIForm_Resize
    
End Sub

Private Sub b8ST_BeforeExpand(Index As Integer)

    'resize contained controlsbeofre expanding
    Select Case Index
        Case m_TabSearch 'search
            'resize
            txtSearchWhat.Move 150, txtSearchWhat.Top, b8ST(Index).Width - 300
            cmbLookIn.Move 150, cmbLookIn.Top, txtSearchWhat.Width
            cmdSearch.Move b8ST(Index).Width - cmdSearch.Width - 150
        Case m_TabFilterDate 'filter date
            b8DateP.Move 150, b8DateP.Top, b8ST(Index).Width - 300
        
        Case m_TabShowQuickLaunch
            listQL.Move 150, listQL.Top, b8ST(Index).Width - 300

    End Select

End Sub

Private Sub b8ST_CompleteExpand(Index As Integer)
    Dim i As Integer
    
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
    
    If b8ST(Index).Expanded = True Then
        Select Case Index
            Case m_TabSearch 'search
                'resize
                txtSearchWhat.Move 150, txtSearchWhat.Top, b8ST(Index).Width - 300
                cmbLookIn.Move 150, cmbLookIn.Top, txtSearchWhat.Width
                cmdSearch.Move b8ST(Index).Width - cmdSearch.Width - 150
            
            Case m_TabFilterDate 'filter date
                b8DateP.Move 150, b8DateP.Top, b8ST(Index).Width - 300
            
            Case m_TabShowQuickLaunch
                listQL.Move 150, listQL.Top, b8ST(Index).Width - 300
                
        End Select
    End If

End Sub





Private Sub cmbLookIn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdSearch_Click
    End If
End Sub



Private Sub cmdVote_Click()
    modFunction.OpenURL "http://www.bob8works.cjb.net/prod_sim.htm", Me.hWnd
End Sub

Private Sub listQL_DblClick()

    Dim selItem As ListItem
    
    On Error GoTo RAE
    
    Set selItem = listQL.SelectedItem
    
    Select Case selItem.Key
        Case "prod" 'Manage Products"
            frmAllProduct.ShowForm
        Case "supp" 'Manage Supliers"
            frmAllSupplier.ShowForm
        Case "cust" 'Manage Customers"
            frmAllCustomer.ShowForm
        Case "poad" 'New P.O."
            frmPOEntry.ShowAdd
        Case "sale" 'New Sales Entry"
            frmSIEntry.ShowAdd
        Case "ppm" 'Purchases/Payments Mon."
            frmAllPPM.ShowForm
        Case "sicpm" 'Sales/Cust.Payments Mon."
            frmAllSICPM.ShowForm
        Case "void" 'Void Products Mon."
            frmAllVoid.ShowForm
        Case "stock" 'Stock Inventory"
            frmAllStockInv.ShowForm
        Case "checkcust" 'Manage Due Checks (Cust.)"
            frmAllCustPayDueCheck.ShowForm
        Case "checksupp" 'Manage Due Checks (Supp.)"
            frmAllPTSDueCheck.ShowForm
    End Select

RAE:
    Set selItem = Nothing
End Sub

Private Sub MDIForm_Load()
    
    'set menus
    Set b8Menus(0).Menu = Me.mnuSystem
    Set b8Menus(1).Menu = Me.mnuRecords
    Set b8Menus(2).Menu = Me.mnuMonitoring
    Set b8Menus(3).Menu = Me.mnuTools
    Set b8Menus(4).Menu = Me.mnuHelp
    
    
    'add quick launch items
    listQL.ListItems.Add , "prod", "Manage Products", 1, 1
    listQL.ListItems.Add , "supp", "Manage Supliers", 1, 1
    listQL.ListItems.Add , "cust", "Manage Customers", 1, 1
    listQL.ListItems.Add , "poad", "New P.O.", 1, 1
    listQL.ListItems.Add , "sale", "New Sales Entry", 1, 1

    listQL.ListItems.Add , "ppm", "Purchases/Payments Mon.", 1, 1
    listQL.ListItems.Add , "sicpm", "Sales/Cust.Payments Mon.", 1, 1
    listQL.ListItems.Add , "void", "Void Products Mon.", 1, 1
    listQL.ListItems.Add , "stock", "Stock Inventory", 1, 1
    
    listQL.ListItems.Add , "checkcust", "Manage Due Checks (Cust.)", 1, 1
    listQL.ListItems.Add , "checksupp", "Manage Due Checks (Supp.)", 1, 1
    

End Sub

Private Sub mnuAboutApp_Click()
    frmSplash.ShowForm
    modFunction.OpenURL "http://www.bob8works.cjb.net/prod_sim.htm", Me.hWnd
End Sub

Private Sub mnuAboutAuthor_Click()
    frmAboutAuthor.ShowForm
End Sub

Private Sub mnuAddBank_Click()
    frmBankEntry.ShowAdd
End Sub

Private Sub mnuAddCateroy_Click()
    frmCatEntry.ShowAdd
End Sub

Private Sub mnuAddCustomer_Click()
    If frmCustEntry.ShowAdd = True Then
        Me.Form_Refresh
    End If
End Sub

Private Sub mnuAddPackage_Click()
    frmPackEntry.ShowAdd
End Sub

Private Sub mnuAddPO_Click()
    If frmPOEntry.ShowAdd = True Then
        Me.Form_Refresh
    End If
End Sub

Private Sub mnuAddProduct_Click()
    If frmProdEntry.ShowAdd = True Then
        Me.Form_Refresh
    End If
End Sub

Private Sub mnuAddSales_Click()
    If frmSIEntry.ShowAdd = True Then
        Me.Form_Refresh
    End If
End Sub

Private Sub mnuAddSupplier_Click()
    If frmSupEntry.ShowAdd = True Then
        Me.Form_Refresh
    End If
End Sub

Private Sub mnuAddUser_Click()
    frmUserEntry.ShowAdd
End Sub

Private Sub mnuAllVoid_Click()
    frmAllVoid.ShowForm
End Sub

Private Sub mnuBackupDatabase_Click()
    frmDBBackup.ShowForm
End Sub

Private Sub mnuDatabaseRestore_Click()
    frmRestore.ShowForm
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuLogOff_Click()
    Me.ShowForm
End Sub

Private Sub mnuManageARAP_Click()
    frmPref.ShowForm 1, 1
End Sub

Private Sub mnuManageBank_Click()
    frmAllBank.ShowForm
End Sub

Private Sub mnuManageCategories_Click()
    frmPref.ShowForm 0, 1
End Sub

Private Sub mnuManageCustomer_Click()
    frmAllCustomer.ShowForm
End Sub

Private Sub mnuManageCustPayDueCheck_Click()
    frmAllCustPayDueCheck.ShowForm
End Sub



Private Sub mnuManagePackages_Click()
    frmPref.ShowForm 0, 1
End Sub

Private Sub mnuManageProducts_Click()
    frmAllProduct.ShowForm
End Sub

Private Sub mnuManagePTSDueCheck_Click()
    frmAllPTSDueCheck.ShowForm
End Sub

Private Sub mnuManageStockInv_Click()
    frmPref.ShowForm 2, 1
End Sub

Private Sub mnuManageSupplier_Click()
    frmAllSupplier.ShowForm
End Sub

Private Sub mnuManageUser_Click()
    frmAllUser.ShowForm
End Sub


Private Sub mnuPPM_Click()
    frmAllPPM.ShowForm
End Sub

Private Sub mnuPreferences_Click()
    frmPref.ShowForm 0
End Sub

Private Sub mnuSICPM_Click()
    frmAllSICPM.ShowForm
End Sub

Private Sub mnuStockInvMon_Click()
    frmAllStockInv.ShowForm
End Sub

Private Sub mnuvisithome_Click()
    modFunction.OpenURL "http://www.bob8works.cjb.net", Me.hWnd
End Sub

Private Sub timeUpdateDate_Timer()
    lblDate.Caption = FormatDateTime(Now, vbGeneralDate)
End Sub

Private Sub txtSearchWhat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdSearch_Click
    End If
End Sub

Private Sub cmdSearch_Click()
    Form_Search
End Sub

'-----------------------------------------------------------
' end Control Procedures


' MDI Form procedures
'-----------------------------------------------------------
Private Sub MDIForm_Resize()
        
    Dim frm As Form
    
    
    On Error Resume Next
    
    'resize header menus bg
    'bgHeaderMenu.Left = b8SBC.Width / Screen.TwipsPerPixelX
    
    'resize bg Record Opt
    bgRecOpt.Move b8SBC.Width / Screen.TwipsPerPixelX
    
    
    'resize childs
    If GetActiveChildCount > 0 Then
        For Each frm In Forms
        If frm.Name <> Me.Name Then
            If frm.MDIChild = True Then
                If frm.Name = Me.ActiveForm.Name Then
                    ResizeMdiChildForm frm
                Else
                    frm.Visible = False
                End If
            End If
        End If
        
        Next
        
    End If
    
    Set frm = Nothing
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    'close database
    CloseDB modMain.PrimeDB
End Sub
'Get Opened MDI Child Forms Count
Public Function GetActiveChildCount() As Integer
    
    Dim frm As Form
    Dim iCount As Integer
    
    iCount = 0
    
    For Each frm In Forms
        If frm.Name <> Me.Name Then
            If frm.MDIChild = True Then
                iCount = iCount + 1
            End If
        End If
    Next
    
    GetActiveChildCount = iCount
    Set frm = Nothing
    
End Function

'-----------------------------------------------------------
' >> End MDI Form procedures
'------------------------------------------------------------



'------------------------------------------------------------
' Parent To Child procedures
'------------------------------------------------------------

Public Sub AddChild(ByRef CFrm As Form)

    'load form
    modFuncChild.LoadForm CFrm
    
End Sub

Public Sub ActivateChild(ByRef CFrm As Form)

    'activate form
    Me.b8CW.SetActiveWindow CFrm.Name
    
    'refresh record operation buttons
    Form_CanAdd
    Form_CanEdit
    Form_CanDelete
    Form_CanRefresh
    Form_CanPrint
    Form_CanSearch
    Form_SetSearch

    
End Sub

Public Sub RemoveChild(ByVal sFormName As String)
    
    'remove form
    Me.b8CW.RemoveChildWindow sFormName
    
End Sub



'Record Operation

Public Function Form_CanAdd() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanAdd
    
    b8RecOpt(0).Enabled = bReturn

    Form_CanAdd = bReturn
    
    Err.Clear
    
End Function
Public Function Form_Add()
    
    If Form_CanAdd Then
        Me.ActiveForm.Form_Add
    End If

End Function


Public Function Form_CanEdit() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanEdit
    
    b8RecOpt(1).Enabled = bReturn

    Form_CanEdit = bReturn
    
    Err.Clear
    
End Function
Public Function Form_Edit()
    
    If Form_CanEdit Then
        Me.ActiveForm.Form_Edit
    End If

End Function


Public Function Form_CanDelete() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanDelete
    
    b8RecOpt(2).Enabled = bReturn

    Form_CanDelete = bReturn
    
    Err.Clear
    
End Function


Public Function Form_Delete()
    
    If Form_CanDelete Then
        Me.ActiveForm.Form_Delete
    End If

End Function


Public Function Form_CanRefresh() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanRefresh
    
    b8RecOpt(3).Enabled = bReturn

    Form_CanRefresh = bReturn
    
    Err.Clear
    
End Function


Public Function Form_Refresh()
    
    If Form_CanRefresh Then
        Me.ActiveForm.Form_Refresh
    End If

End Function



Public Function Form_CanPrint() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanPrint
    
    b8RecOpt(4).Enabled = bReturn

    Form_CanPrint = bReturn
    
    Err.Clear
    
End Function


Public Function Form_Print()
    
    If Form_CanPrint Then
        Me.ActiveForm.Form_Print
    End If

End Function


Public Function Form_CanSearch() As Boolean

    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanSearch
    
    Form_CanSearch = bReturn
    
    Err.Clear
    
End Function



Public Function Form_ShowQuickLaunch()

    'expand side bar
    If b8SBT.SizeState <> ssContracted Then
        b8SBT.SizeState = ssContracted
    End If

    'expand search tab
    If b8ST(m_TabShowQuickLaunch).Expanded = False Then
        b8ST(m_TabShowQuickLaunch).Expanded = True
    End If
    
    On Error Resume Next
    b8ST(m_TabShowQuickLaunch).SetFocus
    'HLTxt txtSearchWhat
    Err.Clear
    
End Function

Public Function Form_ShowSearch()

    'expand side bar
    If b8SBT.SizeState <> ssContracted Then
        b8SBT.SizeState = ssContracted
    End If

    'expand search tab
    If b8ST(m_TabSearch).Expanded = False Then
        b8ST(m_TabSearch).Expanded = True
    End If
    
    On Error Resume Next
    b8ST(m_TabSearch).SetFocus
    HLTxt txtSearchWhat
    Err.Clear
    
End Function


Public Function Form_ShowDateFilter()

    'expand side bar
    If b8SBT.SizeState <> ssContracted Then
        b8SBT.SizeState = ssContracted
    End If

    'expand search tab
    If b8ST(m_TabFilterDate).Expanded = False Then
        b8ST(m_TabFilterDate).Expanded = True
    End If
    
    On Error Resume Next
    b8ST(m_TabFilterDate).SetFocus
    b8DateP.SetFocus
    Err.Clear
    
End Function


Public Function Form_SetSearch()
    Dim bReturn As Boolean
    Dim sFields() As String
    Dim i  As Integer
    
    'clear
    txtSearchWhat.Text = ""
    cmbLookIn.Clear
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_SetSearch(sFields)

    txtSearchWhat.Enabled = bReturn
    cmbLookIn.Enabled = bReturn
    cmdSearch.Enabled = bReturn
    
    If bReturn = True Then
        cmbLookIn.AddItem "All Fields"
        cmbLookIn.ListIndex = 0
        If UBound(sFields) >= 0 Then
            For i = 0 To UBound(sFields)
                cmbLookIn.AddItem sFields(i)
            Next
        End If
    Else
        'contract search tab if it was expanded
        If b8ST(m_TabSearch).Expanded = True Then
            b8ST(m_TabSearch).Expanded = False
        End If
        
    End If
    
    Form_SetSearch = bReturn
    
    Err.Clear
End Function


Public Function Form_Search()
        
    Dim bResult As Boolean
    
    'default
    bResult = False
    
    
    On Error GoTo errh
    
    If txtSearchWhat.Text = "Enter text here" Then
        txtSearchWhat.Text = ""
    End If
    
    If Len(Trim(txtSearchWhat.Text)) <= 0 Then
        MsgBox "Please enter text to search.", vbExclamation
        txtSearchWhat.Text = "Enter text here"
        HLTxt txtSearchWhat
        GoTo errh
    End If
    
    If Len(Trim(cmbLookIn.Text)) <= 0 Then
        MsgBox "Please enter valid field.", vbExclamation
        cmbLookIn.SetFocus
        GoTo errh
    End If
    

    bResult = Me.ActiveForm.Form_Search(Trim(txtSearchWhat.Text), Trim(cmbLookIn.Text))

    If bResult = False Then
        MsgBox "Cannot find '" & txtSearchWhat.Text & "'", vbExclamation
        HLTxt txtSearchWhat
    End If
    
errh:
    Err.Clear
    
End Function

Public Sub Form_DateChange()

    On Error GoTo errh
    Me.ActiveForm.Form_DateChange
errh:
End Sub

Public Function Form_StartBussy()
    Me.MousePointer = vbHourglass
End Function

Public Function Form_EndBussy()
    Me.MousePointer = vbDefault
End Function

Public Sub AFForm_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 83 And Shift = 4 Then
        b8Menus(0).ShowPopUp
    ElseIf KeyCode = 82 And Shift = 4 Then
        b8Menus(1).ShowPopUp
    ElseIf KeyCode = 77 And Shift = 4 Then
        b8Menus(2).ShowPopUp
    ElseIf KeyCode = 84 And Shift = 4 Then
        b8Menus(3).ShowPopUp
    ElseIf KeyCode = 72 And Shift = 4 Then
        b8Menus(4).ShowPopUp
        
    ElseIf KeyCode = 81 And Shift = 2 Then
        'Ctrl + Q
        Me.Form_ShowQuickLaunch
    ElseIf KeyCode = 68 And Shift = 2 Then
        'Ctrl + D
        Me.Form_ShowDateFilter
    End If
    
    'MsgBox KeyCode & " - " & Shift
End Sub
'------------------------------------------------------------
' >>> Parent To Child procedures


'Member variables property
Public Property Get TabSearchIndex() As Integer
    TabSearchIndex = m_TabSearch
End Property


