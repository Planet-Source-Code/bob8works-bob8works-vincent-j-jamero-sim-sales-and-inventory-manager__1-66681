VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmCustEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Entry"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   462
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
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
      Left            =   3870
      TabIndex        =   18
      Top             =   6450
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5430
      TabIndex        =   19
      Top             =   6450
      Width           =   1395
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   6315
      Left            =   0
      ScaleHeight     =   421
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   20
      Top             =   630
      Width           =   7215
      Begin VB.ComboBox cmbAddrProvince 
         Height          =   315
         Left            =   4350
         TabIndex        =   12
         Text            =   "cmbAddrProvince"
         Top             =   3060
         Width           =   2355
      End
      Begin VB.ComboBox cmbAddrCity 
         Height          =   315
         Left            =   1950
         TabIndex        =   10
         Text            =   "cmbAddrCity"
         Top             =   3060
         Width           =   2295
      End
      Begin VB.TextBox txtAddrStreet 
         Height          =   645
         Left            =   1950
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2130
         Width           =   2715
      End
      Begin VB.ComboBox cmbAddrBrgy 
         Height          =   315
         ItemData        =   "frmCustEntry.frx":000C
         Left            =   4770
         List            =   "frmCustEntry.frx":000E
         TabIndex        =   8
         Text            =   "cmbAddrBrgy"
         Top             =   2130
         Width           =   1935
      End
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00F8F8F8&
         Caption         =   "Acti&ve"
         Height          =   255
         Left            =   1950
         TabIndex        =   17
         Top             =   5220
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtCPName 
         Height          =   315
         Left            =   1950
         MaxLength       =   100
         TabIndex        =   14
         Top             =   3990
         Width           =   4755
      End
      Begin VB.TextBox txtCPPosition 
         Height          =   315
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   16
         Top             =   4380
         Width           =   4755
      End
      Begin VB.TextBox txtContactNumber 
         Height          =   315
         Left            =   1950
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1500
         Width           =   2745
      End
      Begin VB.TextBox txtCustName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1020
         Width           =   4755
      End
      Begin VB.TextBox txtCustID 
         BackColor       =   &H00F5F5F5&
         Height          =   315
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   630
         Width           =   1635
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8GradLine b8GradLine1 
         Height          =   240
         Left            =   0
         TabIndex        =   29
         Top             =   240
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "   General"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin b8Controls4.b8GradLine b8GradLine4 
         Height          =   240
         Left            =   0
         TabIndex        =   26
         Top             =   3600
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "   Contact Person"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin b8Controls4.b8GradLine b8GradLine2 
         Height          =   240
         Left            =   0
         TabIndex        =   27
         Top             =   4920
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "   Record Properties"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Province:"
         Height          =   195
         Left            =   4350
         TabIndex        =   11
         Top             =   2820
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Address:"
         Height          =   195
         Left            =   330
         TabIndex        =   4
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&City:"
         Height          =   195
         Left            =   1950
         TabIndex        =   9
         Top             =   2820
         Width           =   345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Brgy.:"
         Height          =   195
         Left            =   4770
         TabIndex        =   7
         Top             =   1890
         Width           =   450
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
         Left            =   2880
         TabIndex        =   25
         Top             =   5430
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
         Left            =   2880
         TabIndex        =   24
         Top             =   5250
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Na&me:"
         Height          =   195
         Left            =   330
         TabIndex        =   13
         Top             =   4050
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posi&tion:"
         Height          =   195
         Left            =   330
         TabIndex        =   15
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Street:"
         Height          =   195
         Left            =   1950
         TabIndex        =   5
         Top             =   1890
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Contact No."
         Height          =   195
         Left            =   330
         TabIndex        =   2
         Top             =   1500
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Name:"
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
         Left            =   330
         TabIndex        =   0
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer &ID"
         Height          =   195
         Left            =   330
         TabIndex        =   22
         Top             =   660
         Width           =   900
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
      ScaleWidth      =   465
      TabIndex        =   23
      Top             =   0
      Width           =   6975
      Begin VB.Label Label8 
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
         Left            =   600
         TabIndex        =   31
         Top             =   390
         Width           =   3900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         TabIndex        =   30
         Top             =   60
         Width           =   1380
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmCustEntry.frx":0010
         Top             =   60
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmCustEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mFormState As String

Dim curCust As tCust
Dim newCust As tCust

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isOn As Boolean

Public Function ShowAdd(Optional sCustName As String = "") As Boolean
    
    'set form state
    mFormState = "add"
    
    'set parameter
    newCust.CustName = sCustName
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function


Public Function ShowAddRetID(ByRef lCUstID As Long, Optional sCustName As String = "") As Boolean
    
    'set form state
    mFormState = "add"
    
    'set parameter
    newCust.CustName = sCustName
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAddRetID = mShowAdd
    lCUstID = newCust.CustID
    
End Function

Public Function ShowEdit(ByVal lCUstID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curCust.CustID = lCUstID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function


Private Sub cmdCancel_Click()
    
    Select Case mFormState
        Case "add"
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
        Case "edit"
            SaveEdit
    End Select
    
    
    
End Sub

Private Sub Form_Activate()
    
    'make mouse pointer bussy
    Me.MousePointer = vbHourglass
    
    'load addresses list
    modRSAddress.FillBrgyToCMB cmbAddrBrgy
    modRSAddress.FillCityToCMB cmbAddrCity
    modRSAddress.FillProvinceToCMB cmbAddrProvince
            
    Select Case mFormState
        Case "add"
                        
            'set form caption
            Me.Caption = "Add New Customer Entry"
            
            'generate new Cust ID
            txtCustID.Text = modFunction.ComNumZ(modRSCust.GetNewCustID, 10)
        
            'get parameter
            txtCustName.Text = newCust.CustName
            
            'set first focused item
            txtCustName.SetFocus
            
        Case "edit"
        
            'set form caption
            Me.Caption = "Edit Customer Entry"
        
            'get Cust info
            If GetCustByID(curCust.CustID, curCust) = False Then
                WriteErrorLog Me.Name, "Form_Activate", "Failed on: 'GetCustByID(curCust.CustID, curCust) = False'"
                'fatal error, close this form
                Unload Me
                GoTo RAE
            End If
            
            'set form's fields
            With curCust
                txtCustID.Text = modFunction.ComNumZ(.CustID, 10)
                txtCustName.Text = .CustName
                txtContactNumber.Text = .ContactNumber
                
                txtAddrStreet.Text = .AddrStreet
                cmbAddrBrgy.Text = .AddrBrgy
                cmbAddrCity.Text = .AddrCity
                cmbAddrProvince.Text = .AddrProvince
                
                txtCPName.Text = .CPName
                txtCPPosition.Text = .CPPosition
                chkActive.Value = IIf(.Active = True, vbChecked, vbUnchecked)
                lblRC.Caption = "Created:  " & .RC & "    By: " & .RCU
                If Not IsEmpty(.RMU) Then
                    lblRM.Caption = "Modified: " & .RM & "    By: " & .RMU
                End If
            End With
            
            
    End Select
    
    
RAE:
    'restoremouse pointer tonormal
    Me.MousePointer = vbNormal
End Sub


Private Sub Form_Load()

    isOn = False
    'PaintGrad bgMain, &HF8F8F8, &HFFFFFF, 90
    PaintGrad bgHeader, &HEDEBE9, &HFFFFFF, 0
    
End Sub



Private Sub SaveAdd()
    
    Dim tmpCust As tCust
    
    
    'validate
    
    'check Customer name
    If IsEmpty(txtCustName.Text) Then
        MsgBox "Please enter Customer Name.", vbExclamation
        HLTxt txtCustName
        Exit Sub
    End If
    
    'check Customer name duplication
    If GetCustByName(txtCustName.Text, tmpCust) = True Then
        MsgBox "The Customer named   '" & txtCustName.Text & "'   was already existed." & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtCustName
        Exit Sub
    End If
    
    
    'set new Customer info
    With newCust
        .CustID = GetTxtVal(txtCustID.Text)
        .CustName = Trim(txtCustName.Text)
        .ContactNumber = Trim(txtContactNumber.Text)
        
        .AddrStreet = txtAddrStreet.Text
        .AddrBrgy = cmbAddrBrgy.Text
        .AddrCity = cmbAddrCity.Text
        .AddrProvince = cmbAddrProvince.Text
                
        .CPName = Trim(txtCPName.Text)
        .CPPosition = Trim(txtCPPosition.Text)
        .BegAR = 0
        .Active = IIf(chkActive.Value = vbChecked, True, False)
        .RC = Now
        .RM = Now
        .RCU = CurrentUser.UserID
        .RMU = ""
    End With
    
    'save
    If AddCust(newCust) = True Then
        
        'success
        
        'save new addresses
        modRSAddress.AddBrgy cmbAddrBrgy.Text
        modRSAddress.AddCity cmbAddrCity.Text
        modRSAddress.AddProvince cmbAddrProvince.Text
    
        'set flag
        mShowAdd = True
        
        'close this form
        Unload Me
    
    Else
        WriteErrorLog Me.Name, "SaveAdd", "Failed on : 'AddCust(newCust) = True'"
    End If
    
End Sub


Private Sub SaveEdit()
    
    Dim tmpCust As tCust
    
    'validate
    
    'check Customer name
    If IsEmpty(txtCustName.Text) Then
        MsgBox "Please enter Customer Name.", vbExclamation
        HLTxt txtCustName
        Exit Sub
    End If
    
    If LCase(Trim(curCust.CustName)) <> LCase(Trim(txtCustName.Text)) Then
        'check Customer name duplication
        If GetCustByName(txtCustName.Text, tmpCust) = True Then
            MsgBox "The Customer named   '" & txtCustName.Text & "'   was already existed." & vbNewLine & _
                "Please enter different value.", vbExclamation
            HLTxt txtCustName
            Exit Sub
        End If
    End If
    
    
    'set new Customer info
    With curCust
        '.CustID = GetTxtVal(txtCustID.Text)
        .CustName = Trim(txtCustName.Text)
        .ContactNumber = Trim(txtContactNumber.Text)
        
        .AddrStreet = txtAddrStreet.Text
        .AddrBrgy = cmbAddrBrgy.Text
        .AddrCity = cmbAddrCity.Text
        .AddrProvince = cmbAddrProvince.Text
        
        .CPName = Trim(txtCPName.Text)
        .CPPosition = Trim(txtCPPosition.Text)
        '.BegAP = 0
        .Active = IIf(chkActive.Value = vbChecked, True, False)
        '.RC = Now
        .RM = Now
        '.RCU = CurrentUser.UserID
        .RMU = CurrentUser.UserID
    End With
    
    'save
    If EditCust(curCust) = True Then
        
        'success
        
        'save new addresses
        modRSAddress.AddBrgy cmbAddrBrgy.Text
        modRSAddress.AddCity cmbAddrCity.Text
        modRSAddress.AddProvince cmbAddrProvince.Text
        
        'set flag
        mShowEdit = True
        
        'close this form
        Unload Me
    
    Else
        WriteErrorLog Me.Name, "SaveAdd", "Failed on : 'EditCust(curCust) = True'"
    End If
        
End Sub


