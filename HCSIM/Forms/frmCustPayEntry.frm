VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustPayEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payment Entry"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustPayEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   5055
      Left            =   0
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   613
      TabIndex        =   0
      Top             =   540
      Width           =   9195
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00F5F5F5&
         BorderStyle     =   0  'None
         Height          =   1395
         Left            =   4290
         ScaleHeight     =   1395
         ScaleWidth      =   4185
         TabIndex        =   30
         Top             =   3030
         Width           =   4185
         Begin VB.TextBox txtRemarks 
            Height          =   1155
            Left            =   1350
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   60
            Width           =   2835
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Remarks:"
            Height          =   195
            Left            =   0
            TabIndex        =   32
            Top             =   30
            Width           =   675
         End
      End
      Begin VB.PictureBox bgCheck 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1605
         Left            =   210
         ScaleHeight     =   1605
         ScaleWidth      =   8325
         TabIndex        =   14
         Top             =   1890
         Width           =   8325
         Begin VB.CheckBox chkCleared 
            BackColor       =   &H00F5F5F5&
            Caption         =   "Cleared"
            Height          =   255
            Left            =   5430
            TabIndex        =   33
            Top             =   810
            Width           =   885
         End
         Begin VB.TextBox txtAccountName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5430
            MaxLength       =   50
            TabIndex        =   24
            Top             =   0
            Width           =   2835
         End
         Begin VB.TextBox txtAccountNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   22
            Top             =   390
            Width           =   2085
         End
         Begin MSComCtl2.DTPicker dtpDateDue 
            Height          =   315
            Left            =   1410
            TabIndex        =   21
            Top             =   1200
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMM - dd - yyyy"
            Format          =   181010435
            CurrentDate     =   38961
         End
         Begin MSComCtl2.DTPicker dtpDateIssued 
            Height          =   315
            Left            =   1410
            TabIndex        =   20
            Top             =   810
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMM - dd - yyyy"
            Format          =   181010435
            CurrentDate     =   38961
         End
         Begin VB.TextBox txtCheckNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   15
            Top             =   0
            Width           =   2085
         End
         Begin b8Controls4.b8DataPicker b8dpBankName 
            Height          =   345
            Left            =   5430
            TabIndex        =   26
            Top             =   360
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   609
            SQLWhereSeparator=   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextLocked      =   0   'False
            DropWinWidth    =   4485
            Locked          =   0   'False
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Bank:"
            Height          =   195
            Left            =   4080
            TabIndex        =   27
            Top             =   420
            Width           =   405
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Account No.:"
            Height          =   195
            Left            =   60
            TabIndex        =   23
            Top             =   390
            Width           =   945
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Due Date:"
            Height          =   195
            Left            =   60
            TabIndex        =   19
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Date Issued:"
            Height          =   195
            Left            =   60
            TabIndex        =   18
            Top             =   810
            Width           =   930
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Account Name"
            Height          =   195
            Left            =   4080
            TabIndex        =   17
            Top             =   30
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Check No.:"
            Height          =   195
            Left            =   60
            TabIndex        =   16
            Top             =   30
            Width           =   795
         End
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   28
         Top             =   3870
         Width           =   2055
      End
      Begin VB.ComboBox cmbFP 
         BackColor       =   &H00F4FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   2475
      End
      Begin VB.TextBox txtOtherFP 
         Height          =   315
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1470
         Width           =   2445
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   7260
         TabIndex        =   9
         Top             =   4530
         Width           =   1395
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   5700
         TabIndex        =   8
         Top             =   4530
         Width           =   1395
      End
      Begin VB.TextBox txtCustPayID 
         BackColor       =   &H00F5F5F5&
         Height          =   315
         Left            =   4860
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   90
         Width           =   1635
      End
      Begin b8Controls4.b8Line b8Line1 
         Height          =   30
         Left            =   -30
         TabIndex        =   1
         Top             =   450
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin MSComCtl2.DTPicker dtpCustPayDate 
         Height          =   315
         Left            =   7140
         TabIndex        =   2
         Top             =   90
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM - dd - yyyy"
         Format          =   181010435
         CurrentDate     =   38961
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   4
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
         TabIndex        =   29
         Top             =   4410
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8DataPicker b8DPCust 
         Height          =   360
         Left            =   1620
         TabIndex        =   34
         Top             =   600
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropWinWidth    =   6210
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Customer:"
         Height          =   195
         Left            =   270
         TabIndex        =   35
         Top             =   660
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Amount:"
         Height          =   195
         Left            =   270
         TabIndex        =   25
         Top             =   3900
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Form of Payment:"
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   1110
         Width           =   1290
      End
      Begin VB.Label lblOtherFP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Other"
         Height          =   195
         Left            =   420
         TabIndex        =   12
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         Height          =   195
         Left            =   4560
         TabIndex        =   6
         Top             =   90
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Left            =   6660
         TabIndex        =   5
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   7
      Top             =   0
      Width           =   10305
      Begin VB.Label Label14 
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
         TabIndex        =   37
         Top             =   360
         Width           =   3900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Payment Entry"
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
         TabIndex        =   36
         Top             =   30
         Width           =   3570
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   60
         Picture         =   "frmCustPayEntry.frx":000C
         Top             =   30
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmCustPayEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mFormState As String

Dim ProdPackList() As tProdPack

Dim curCustPay As tCustPay
Dim newCustPay As tCustPay

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isOn As Boolean

Dim mAmountDue As Double

Public Function ShowAdd(Optional ByVal dCustPayDate As Date = 0, _
                        Optional ByVal lCustID As Long = 0, _
                        Optional ByVal sFP As String = "", _
                        Optional ByVal dAmount As Double = 0, _
                        Optional ByVal dAmountDue As Double = 0, _
                        Optional ByVal sRemarks As String = "", _
                        Optional ByRef lNewCustPayID As Long) As Boolean
    
    'set form state
    mFormState = "add"
    
    'evaluate param
    If dCustPayDate = 0 Then
        newCustPay.CustPayDate = Now
    Else
        newCustPay.CustPayDate = dCustPayDate
        dtpCustPayDate.Enabled = False
    End If
    
    newCustPay.FP = sFP
    newCustPay.Amount = dAmount
    newCustPay.Remarks = sRemarks
    newCustPay.FK_CustID = lCustID
    mAmountDue = dAmountDue
 
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    On Error Resume Next
    lNewCustPayID = newCustPay.CustPayID
    Err.Clear
    
End Function


Public Function ShowEdit(ByVal lCustPayID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curCustPay.CustPayID = lCustPayID
    
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


Private Sub cmbFP_Change()
    
    Select Case cmbFP.ListIndex
        Case 0 'cash
            Call Form_ShowCheck(False)
            txtOtherFP.Visible = False
            lblOtherFP.Visible = False
        Case 1 'check
            Call Form_ShowCheck(True)
            txtOtherFP.Visible = False
            lblOtherFP.Visible = False
        Case 2 'other
            Call Form_ShowCheck(False)
            txtOtherFP.Visible = True
            lblOtherFP.Visible = True
    End Select
    
End Sub

Private Sub cmbFP_Click()
    Call cmbFP_Change
End Sub








Private Sub Form_Activate()
    
    If isOn = True Then
        Exit Sub
    End If
    isOn = True
    
    
    'make mouse pointer bussy
    Me.MousePointer = vbHourglass
   
    Select Case mFormState
        Case "add"
            
            Me.Caption = "Add Payment Entry"
            
            Form_RefreshFP newCustPay.FP
            
            txtCustPayID.Text = modFunction.ComNumZ(modRSCustPay.GetNewCustPayID, 10)
            
            dtpCustPayDate.Value = newCustPay.CustPayDate
            dtpDateIssued.Value = newCustPay.CustPayDate
            dtpDateDue.Value = newCustPay.CustPayDate
            
            'supplier
            
            If Form_UseThisCust(newCustPay.FK_CustID) = True Then
                b8DPCust.Enabled = False
            End If
            
            txtAmount.Text = FormatNumber(newCustPay.Amount, 2)
            txtRemarks.Text = newCustPay.Remarks
            
        Case "edit"
        
            Me.Caption = "Add Payment Entry"
            
            If GetCustPayByID(curCustPay.CustPayID, curCustPay) = False Then
                WriteErrorLog Me.Name, "", ""
                Unload Me
                GoTo RAE
            End If
            
            With curCustPay
                txtCustPayID.Text = modFunction.ComNumZ(.CustPayID, 10)
                dtpCustPayDate.Value = .CustPayDate
                Form_UseThisCust curCustPay.FK_CustID
                Form_RefreshFP .FP
                
                If LCase(.FP) = "check" Then
                    txtCheckNo.Text = .CheckNo
                    txtAccountNo.Text = .AccountNo
                    dtpDateIssued.Value = .DateIssued
                    dtpDateDue.Value = .DateDue
                    txtAccountName.Text = .AccountName
                    b8dpBankName.DisplayData = .BankName
                    chkCleared.Value = IIf(.Cleared, vbChecked, vbUnchecked)
                    
                ElseIf LCase(.FP) <> "cash" Then
                    txtOtherFP.Text = .FP
                End If
                
                txtAmount.Text = FormatNumber(.Amount, 2)
                txtRemarks.Text = .Remarks
                
            End With
            
            'disable some controls
            b8DPCust.Enabled = False
            
    End Select
    
    
RAE:
    'restoremouse pointer tonormal
    Me.MousePointer = vbNormal
End Sub


Private Sub Form_Load()
    
    isOn = False
    PaintGrad bgHeader, &HEDEBE9, &HFFFFFF, 0

    'set bank list
    With b8dpBankName
        Set .DropDBCon = PrimeDB
        .SQLFields = "tblBank.BankName, tblBank.Address"
        .SQLTable = "tblBank"
        .SQLWhereFields = "tblBank.BankName, tblBank.Address"
        .SQLOrderBy = "tblBank.BankName"
        
        .BoundFieldIndex = 0 'Bound Index
        .DisplayFieldIndex = 0
        .AddColumn "Bank", 120
        .AddColumn "Address", 160

    End With
    
    
    'set Customer list
    With b8DPCust
        Set .DropDBCon = PrimeDB
        .SQLFields = "String(10-Len(Trim([CustID])),'0') & [CustID] AS CCustID, tblCust.CustName"
        .SQLTable = "tblCust"
        .SQLWhereFields = "tblCust.CustID, tblCust.CustName"
        .SQLWhere = " tblCust.Active = True "
        .SQLOrderBy = "tblCust.CustName"
        
        .BoundFieldIndex = 0 'Bound Index
        .DisplayFieldIndex = 1
        .AddColumn "Customer ID", 100
        .AddColumn "Customer", 240
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    isOn = False
End Sub



Private Sub SaveAdd()
    
    'validate
    'amount
    If Not (GetTxtVal(txtAmount.Text) > 0) Then
        MsgBox "Please enter valid amount.", vbExclamation
        HLTxt txtAmount
        Exit Sub
    End If
    If mAmountDue > 0 Then
        If GetTxtVal(txtAmount.Text) > mAmountDue Then
            MsgBox "Please enter valid amount. It mus be less than or equal to " & FormatNumber(mAmountDue, 2) & ".", vbExclamation
            HLTxt txtAmount
            Exit Sub
        End If
    End If
            
    'supplier
    If IsEmpty(b8DPCust.BoundData) Then
        MsgBox "Please select Customer.", vbExclamation
        b8DPCust.FocusedDropButton
        Exit Sub
    End If
    
    
            
    Select Case cmbFP.ListIndex
        Case 0 'cash
        
            With newCustPay
                .FP = "cash"
                .Cleared = True
                
                '.AccountName
                'CheckNo
                'DateDue
                'AccountNo
                'BankName
            End With
                        
        Case 1 'check
            
            If IsEmpty(txtCheckNo.Text) Then
                MsgBox "Please enter 'Check No.'", vbExclamation
                HLTxt txtCheckNo
                Exit Sub
            End If
            
            If IsEmpty(txtAccountNo.Text) Then
                MsgBox "Please enter 'Account No.'", vbExclamation
                HLTxt txtAccountNo
                Exit Sub
            End If
            
            If DateValue(dtpDateIssued) > DateValue(dtpDateDue) Then
                MsgBox "'Due Date' must be less than or equal to 'Date Issued'", vbExclamation
                dtpDateDue.SetFocus
                Exit Sub
            End If
            
            If IsEmpty(b8dpBankName.DisplayData) Then
                MsgBox "Please enter 'Bank Name'", vbExclamation
                b8dpBankName.SetFocus
                Exit Sub
            End If
            
            'check check no duplication
            Dim tmpCustPay As tCustPay
            If GetCustPayByCheckNo(Trim(txtCheckNo.Text), Trim(b8dpBankName.DisplayData), tmpCustPay) = True Then
                MsgBox "The Check No. that you have entered is already existed.", vbExclamation
                HLTxt txtCheckNo
                Exit Sub
            End If
            
            With newCustPay
                .FP = "check"
                .AccountName = txtAccountName.Text
                .CheckNo = txtCheckNo.Text
                .DateDue = DateValue(dtpDateDue.Value)
                .DateIssued = DateValue(dtpDateIssued.Value)
                .AccountNo = Trim(txtAccountNo.Text)
                .BankName = Trim(b8dpBankName.DisplayData)
                .Cleared = IIf(chkCleared.Value = vbChecked, True, False)
            End With
            
        Case 2 'other
        
            If Len(Trim(txtOtherFP.Text)) < 1 Then
                MsgBox "Please enter other Form of Payment.", vbExclamation
                HLTxt txtOtherFP
                Exit Sub
            End If

            If LCase(Trim(txtOtherFP.Text)) = "check" Or LCase(Trim(txtOtherFP.Text)) = "cash" Then
                MsgBox "Other Form of Payment must not be be 'Check' or 'Cash'", vbExclamation
                HLTxt txtOtherFP
                Exit Sub
            End If
            
            With newCustPay

                .FP = Trim(txtOtherFP.Text)
                '.AccountName
                'CheckNo
                'DateDue
                'AccountNo
                'BankName
                .Cleared = True

            End With
                        
    End Select
    
    'setremaining CustPay Info
    With newCustPay
        .CustPayID = CLng(GetTxtVal(txtCustPayID.Text))
        .FK_CustID = CLng(b8DPCust.BoundData)
        .CustPayDate = dtpCustPayDate.Value
        .Amount = GetTxtVal(txtAmount.Text)
        .Remarks = txtRemarks.Text
        .RC = Now
        'RM
        .RCU = CurrentUser.UserID
        'RMU
    End With
    
    If AddCustPay(newCustPay) = True Then
        'set flag
        mShowAdd = True
        'close this form
        Unload Me
    Else
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'AddCustPay(newCustPay) = True'"
    End If
    
End Sub


Private Sub SaveEdit()
    
    'validate
    'amount
    If Not (GetTxtVal(txtAmount.Text) > 0) Then
        MsgBox "Please enter valid amount.", vbExclamation
        HLTxt txtAmount
        Exit Sub
    End If
    If mAmountDue > 0 Then
        If GetTxtVal(txtAmount.Text) > mAmountDue Then
            MsgBox "Please enter valid amount. It mus be less than or equal to " & FormatNumber(mAmountDue, 2) & ".", vbExclamation
            HLTxt txtAmount
            Exit Sub
        End If
    End If
            
    'supplier
    If IsEmpty(b8DPCust.BoundData) Then
        MsgBox "Please select Customer.", vbExclamation
        b8DPCust.FocusedDropButton
        Exit Sub
    End If
    
    
            
    Select Case cmbFP.ListIndex
        Case 0 'cash
        
            With curCustPay
                .FP = "cash"
                .Cleared = True
                
                '.AccountName
                'CheckNo
                'DateDue
                'AccountNo
                'BankName
            End With
                        
        Case 1 'check
            
            If IsEmpty(txtCheckNo.Text) Then
                MsgBox "Please enter 'Check No.'", vbExclamation
                HLTxt txtCheckNo
                Exit Sub
            End If
            
            If IsEmpty(txtAccountNo.Text) Then
                MsgBox "Please enter 'Account No.'", vbExclamation
                HLTxt txtAccountNo
                Exit Sub
            End If
            
            If DateValue(dtpDateIssued) > DateValue(dtpDateDue) Then
                MsgBox "'Due Date' must be less than or equal to 'Date Issued'", vbExclamation
                dtpDateDue.SetFocus
                Exit Sub
            End If
            
            If IsEmpty(b8dpBankName.DisplayData) Then
                MsgBox "Please enter 'Bank Name'", vbExclamation
                b8dpBankName.SetFocus
                Exit Sub
            End If
            
            'check check no duplication
            If Trim(curCustPay.CheckNo) <> Trim(txtCheckNo.Text) Then
                Dim tmpCustPay As tCustPay
                If GetCustPayByCheckNo(Trim(txtCheckNo.Text), Trim(b8dpBankName.DisplayData), tmpCustPay) = True Then
                    MsgBox "The Check No. that you have entered is already existed.", vbExclamation
                    HLTxt txtCheckNo
                    Exit Sub
                End If
            End If
            
            With curCustPay
                .FP = "check"
                .AccountName = txtAccountName.Text
                .CheckNo = txtCheckNo.Text
                .DateDue = DateValue(dtpDateDue.Value)
                .DateIssued = DateValue(dtpDateIssued.Value)
                .AccountNo = Trim(txtAccountNo.Text)
                .BankName = Trim(b8dpBankName.DisplayData)
                .Cleared = IIf(chkCleared.Value = vbChecked, True, False)
            End With
            
        Case 2 'other
        
            If Len(Trim(txtOtherFP.Text)) < 1 Then
                MsgBox "Please enter other Form of Payment.", vbExclamation
                HLTxt txtOtherFP
                Exit Sub
            End If

            If LCase(Trim(txtOtherFP.Text)) = "check" Or LCase(Trim(txtOtherFP.Text)) = "cash" Then
                MsgBox "Other Form of Payment must not be be 'Check' or 'Cash'", vbExclamation
                HLTxt txtOtherFP
                Exit Sub
            End If
            
            With curCustPay

                .FP = Trim(txtOtherFP.Text)
                '.AccountName
                'CheckNo
                'DateDue
                'AccountNo
                'BankName
                .Cleared = True

            End With
                        
    End Select
    
    'setremaining CustPay Info
    With curCustPay
        '.CustPayID = CLng(GetTxtVal(txtCustPayID.Text))
        .FK_CustID = CLng(b8DPCust.BoundData)
        .CustPayDate = dtpCustPayDate.Value
        .Amount = GetTxtVal(txtAmount.Text)
        .Remarks = txtRemarks.Text
        '.RC
        .RM = Now
        '.RCU
        .RMU = CurrentUser.UserID
    End With
    
    If EditCustPay(curCustPay) = True Then
        'set flag
        mShowEdit = True
        'close this form
        Unload Me
    Else
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'editCustPay(curpts) = True'"
    End If
    
End Sub


Private Sub Form_RefreshFP(Optional sFP As String = "Cash")
    
    cmbFP.AddItem "Cash"
    cmbFP.AddItem "Check"
    cmbFP.AddItem "Other"

    
    Dim i As Integer
    For i = 0 To cmbFP.ListCount - 1
        If LCase(Trim(cmbFP.List(i))) = LCase(Trim(sFP)) Then
            cmbFP.ListIndex = i
            Exit Sub
        End If
    Next
    
    cmbFP.ListIndex = 2
    
End Sub

'--------------------------------------------------------------
'CHECK Info Procedures
'--------------------------------------------------------------

Private Sub Form_ShowCheck(ByVal NewValue As Boolean)

    Dim lBackColor As Long
    
    txtCheckNo.Enabled = NewValue
    txtAccountNo.Enabled = NewValue
    dtpDateIssued.Enabled = NewValue
    dtpDateDue.Enabled = NewValue
    txtAccountName.Enabled = NewValue
    b8dpBankName.ClearEnabled = NewValue
    b8dpBankName.DropEnabled = NewValue
    chkCleared.Enabled = NewValue
    
    If NewValue = True Then
        lBackColor = vbWhite
    Else
        lBackColor = &H8000000F
    End If
    
    txtCheckNo.BackColor = lBackColor
    txtAccountNo.BackColor = lBackColor
    txtAccountName.BackColor = lBackColor
    b8dpBankName.BackColor = lBackColor


    bgCheck.Enabled = NewValue
    
End Sub

Private Sub Form_SetCleared()
    
    If DateValue(dtpDateIssued) = DateValue(dtpDateDue) Then
        chkCleared.Value = vbChecked
    Else
        chkCleared.Value = vbUnchecked
    End If
    
End Sub

Private Sub dtpDateDue_Change()

    Form_SetCleared
End Sub

Private Sub dtpDateIssued_Change()

    Form_SetCleared
End Sub

'--------------------------------------------------------
'end Check Info procedures
'--------------------------------------------------------


Private Sub txtAmount_Validate(Cancel As Boolean)
    
    Cancel = True
    
    If Not (GetTxtVal(txtAmount.Text) > 0) Then
        MsgBox "Please enter valid amount.", vbExclamation
        HLTxt txtAmount
        Exit Sub
    End If
    
    If mAmountDue > 0 Then
        If GetTxtVal(txtAmount.Text) > mAmountDue Then
            MsgBox "Please enter valid amount. It mus be less than or equal to " & FormatNumber(mAmountDue, 2) & ".", vbExclamation
            HLTxt txtAmount
            Exit Sub
        End If
    End If
        
    Cancel = False
End Sub


Private Function Form_UseThisCust(ByVal lCustID As Long) As Boolean
    
    Dim vCust As tCust
    
    If modRSCust.GetCustByID(lCustID, vCust) = True Then
    
        b8DPCust.BoundData = vCust.CustID
        b8DPCust.DisplayData = vCust.CustName
        
        Form_UseThisCust = True
    Else
        Form_UseThisCust = False
    End If
    
End Function
