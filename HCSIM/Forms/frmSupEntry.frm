VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmSupEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supplier Entry"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSupEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
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
      Left            =   3840
      TabIndex        =   11
      Top             =   6180
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   6180
      Width           =   1395
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   409
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   13
      Top             =   630
      Width           =   6915
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00F8F8F8&
         Caption         =   "Acti&ve"
         Height          =   255
         Left            =   1950
         TabIndex        =   10
         Top             =   4620
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtCPName 
         Height          =   315
         Left            =   1950
         MaxLength       =   100
         TabIndex        =   7
         Top             =   3270
         Width           =   4635
      End
      Begin VB.TextBox txtCPPosition 
         Height          =   315
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   9
         Top             =   3660
         Width           =   4635
      End
      Begin VB.TextBox txtAddress 
         Height          =   630
         Left            =   1950
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1920
         Width           =   4635
      End
      Begin VB.TextBox txtContactNumber 
         Height          =   315
         Left            =   1950
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1530
         Width           =   2835
      End
      Begin VB.TextBox txtSupName 
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
         Top             =   1050
         Width           =   4635
      End
      Begin VB.TextBox txtSupID 
         BackColor       =   &H00F5F5F5&
         Height          =   315
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   660
         Width           =   1635
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   240
         Width           =   5625
         _ExtentX        =   9922
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
         TabIndex        =   22
         Top             =   2850
         Width           =   5625
         _ExtentX        =   9922
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
         TabIndex        =   23
         Top             =   4260
         Width           =   5625
         _ExtentX        =   9922
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
         Left            =   1950
         TabIndex        =   18
         Top             =   5070
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
         Left            =   1950
         TabIndex        =   17
         Top             =   4890
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Na&me:"
         Height          =   195
         Left            =   330
         TabIndex        =   6
         Top             =   3330
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Position:"
         Height          =   195
         Left            =   330
         TabIndex        =   8
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Address:"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   4
         Top             =   1860
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Contact No."
         Height          =   195
         Left            =   330
         TabIndex        =   2
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* &Name:"
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
         Top             =   1050
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Supplier ID"
         Height          =   195
         Left            =   330
         TabIndex        =   15
         Top             =   690
         Width           =   780
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
      TabIndex        =   16
      Top             =   0
      Width           =   6975
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
         Left            =   570
         TabIndex        =   24
         Top             =   420
         Width           =   3900
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmSupEntry.frx":000C
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         Index           =   1
         Left            =   570
         TabIndex        =   19
         Top             =   30
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmSupEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mFormState As String

Dim curSup As tSup
Dim newSup As tSup

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isOn As Boolean

Public Function ShowAdd(Optional sSupName As String = "") As Boolean
    
    'set form state
    mFormState = "add"
    
    'set parameter
    newSup.SupName = sSupName
    
    'show form
    Me.Show vbModal
    
    'return
    
    ShowAdd = mShowAdd
    
End Function

Public Function ShowAddRetID(Optional sSupName As String = "") As Long
    
    'set form state
    mFormState = "add"
    
    'set parameter
    newSup.SupName = sSupName
    
    'show form
    Me.Show vbModal
    
    'return
    If mShowAdd = True Then
        ShowAddRetID = newSup.SupID
    Else
        ShowAddRetID = -1
    End If
    
End Function

Public Function ShowEdit(ByVal lSupID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curSup.SupID = lSupID
    
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
    
    
    Select Case mFormState
        Case "add"
        
            'set form caption
            Me.Caption = "Add New Supplier Entry"
            
            'generate new Sup ID
            txtSupID.Text = modFunction.ComNumZ(modRSSup.GetNewSupID, 10)
        
            'get parameter
            txtSupName.Text = newSup.SupName
            
            'set first focused item
            txtSupName.SetFocus
            
        Case "edit"
        
            'set form caption
            Me.Caption = "Edit Supplier Entry"
            
            'get sup info
            If GetSupByID(curSup.SupID, curSup) = False Then
                WriteErrorLog Me.Name, "Form_Activate", "Failed on: 'GetSupByID(curSup.SupID, curSup) = False'"
                'fatal error, close this form
                Unload Me
                GoTo RAE
            End If
            
            'set form's fields
            With curSup
                txtSupID.Text = modFunction.ComNumZ(.SupID, 10)
                txtSupName.Text = .SupName
                txtContactNumber.Text = .ContactNumber
                txtAddress.Text = .Address
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
    
    Dim tmpSup As tSup
    
    
    'validate
    
    'check supplier name
    If IsEmpty(txtSupName.Text) Then
        MsgBox "Please enter Supplier Name.", vbExclamation
        HLTxt txtSupName
        Exit Sub
    End If
    
    'check supplier name duplication
    If GetSupByName(txtSupName.Text, tmpSup) = True Then
        MsgBox "The Supplier named   '" & txtSupName.Text & "'   was already existed." & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtSupName
        Exit Sub
    End If
    
    
    'set new supplier info
    With newSup
        .SupID = GetTxtVal(txtSupID.Text)
        .SupName = Trim(txtSupName.Text)
        .ContactNumber = Trim(txtContactNumber.Text)
        .Address = Trim(txtAddress.Text)
        .CPName = Trim(txtCPName.Text)
        .CPPosition = Trim(txtCPPosition.Text)
        .BegAP = 0
        .Active = IIf(chkActive.Value = vbChecked, True, False)
        .RC = Now
        .RM = Now
        .RCU = CurrentUser.UserID
        .RMU = ""
    End With
    
    'save
    If AddSup(newSup) = True Then
        
        'success
        
        'set flag
        mShowAdd = True
        
        'close this form
        Unload Me
    
    Else
        WriteErrorLog Me.Name, "SaveAdd", "Failed on : 'AddSup(newSup) = True'"
    End If
    
End Sub


Private Sub SaveEdit()
    
    Dim tmpSup As tSup
    
    'validate
    
    'check supplier name
    If IsEmpty(txtSupName.Text) Then
        MsgBox "Please enter Supplier Name.", vbExclamation
        HLTxt txtSupName
        Exit Sub
    End If
    
    If LCase(Trim(curSup.SupName)) <> LCase(Trim(txtSupName.Text)) Then
        'check supplier name duplication
        If GetSupByName(txtSupName.Text, tmpSup) = True Then
            MsgBox "The Supplier named   '" & txtSupName.Text & "'   was already existed." & vbNewLine & _
                "Please enter different value.", vbExclamation
            HLTxt txtSupName
            Exit Sub
        End If
    End If
    
    
    'set new supplier info
    With curSup
        '.SupID = GetTxtVal(txtSupID.Text)
        .SupName = Trim(txtSupName.Text)
        .ContactNumber = Trim(txtContactNumber.Text)
        .Address = Trim(txtAddress.Text)
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
    If EditSup(curSup) = True Then
        
        'success
        
        'set flag
        mShowEdit = True
        
        'close this form
        Unload Me
    
    Else
        WriteErrorLog Me.Name, "SaveAdd", "Failed on : 'EditSup(curSup) = True'"
    End If
    
    
End Sub
