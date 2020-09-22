VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl b8DatePicker 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2580
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   172
   Begin VB.OptionButton optDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View by Date"
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Value           =   -1  'True
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpByDate 
      Height          =   345
      Left            =   330
      TabIndex        =   5
      Top             =   330
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20578305
      CurrentDate     =   38968
   End
   Begin b8Controls4.b8MonthView b8MV 
      Height          =   345
      Left            =   330
      TabIndex        =   1
      Top             =   1170
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   609
      MonthVal        =   9
      YearVal         =   2006
      Enabled         =   0   'False
   End
   Begin b8Controls4.b8YearView b8YV 
      Height          =   315
      Left            =   330
      TabIndex        =   0
      Top             =   2010
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      YearVal         =   2006
      YearToString    =   "2006"
   End
   Begin VB.OptionButton optDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View by Month"
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   870
      Width           =   1455
   End
   Begin VB.OptionButton optDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View by Year"
      Height          =   315
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1740
      Width           =   1455
   End
End
Attribute VB_Name = "b8DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_ViewIndex = 0
Const m_def_MinDate = 0
'Property Variables:
Dim m_ViewIndex As Integer
Dim m_MinDate As Date
Dim m_MaxDate As Date
'Event Declarations:
Event Change()


Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function

Private Function GetViewIndex() As Integer
    
    Dim i As Integer
    
    For i = 0 To optDate.UBound
        If optDate(i).Value = True Then
            GetViewIndex = i
            Exit For
        End If
    Next

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Dim i As Integer
        
    For i = 0 To optDate.UBound
        optDate(i).BackColor = New_BackColor
    Next
    UserControl.BackColor() = New_BackColor
    
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=3,0,0,0
Public Property Get MinDate() As Date

    Select Case GetViewIndex
        Case 0 'by date
            m_MinDate = DateValue(dtpByDate.Value)
        Case 1 ' by month
            m_MinDate = DateSerial(b8MV.YearVal, b8MV.MonthVal, 1)
        Case 2 'by year
            m_MinDate = DateSerial(b8YV.YearVal, 1, 1)
    End Select
    
    MinDate = m_MinDate
End Property

Public Property Let MinDate(ByVal New_MinDate As Date)
    
    Select Case GetViewIndex
        Case 0 'by date
            dtpByDate.Value = New_MinDate
        Case 1 ' by month
            b8MV.YearVal = Year(New_MinDate)
            b8MV.MonthVal = Month(New_MinDate)
        Case 2 'by year
            b8YV.YearVal = Year(New_MinDate)
    End Select
    
    m_MinDate = New_MinDate
    PropertyChanged "MinDate"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=3,0,0,0
Public Property Get MaxDate() As Date
        
    Select Case GetViewIndex
        Case 0 'by date
            m_MaxDate = DateValue(dtpByDate.Value)
        Case 1 ' by month
            m_MaxDate = DateSerial(b8MV.YearVal, b8MV.MonthVal + 1, 0)
        Case 2 'by year
            m_MaxDate = DateSerial(b8YV.YearVal + 1, 1, 0)
    End Select
    
    MaxDate = m_MaxDate
    
End Property

Public Property Let MaxDate(ByVal New_MaxDate As Date)
    
    Select Case GetViewIndex
        Case 0 'by date
            dtpByDate.Value = New_MaxDate
        Case 1 ' by month
            b8MV.YearVal = Year(New_MaxDate)
            b8MV.MonthVal = Month(New_MaxDate)
        Case 2 'by year
            b8YV.YearVal = Year(New_MaxDate)
    End Select
        
    m_MaxDate = New_MaxDate
    
    PropertyChanged "MaxDate"
End Property



Private Sub b8MV_Change()
    RaiseEvent Change
End Sub

Private Sub b8YV_Change()
    RaiseEvent Change
End Sub

Private Sub b8YV_Click()
    RaiseEvent Change
End Sub

Private Sub dtpByDate_Change()
    RaiseEvent Change
End Sub

Private Sub dtpByDate_Click()
    RaiseEvent Change
End Sub

Private Sub optDate_Click(Index As Integer)

    dtpByDate.Enabled = False
    b8MV.Enabled = False
    b8YV.Enabled = False
    
    If optDate(Index).Value = True Then
        
        Select Case Index
            Case 0 'by date
                dtpByDate.Enabled = True
            Case 1 'by month
                b8MV.Enabled = True
            Case 2 'by year
                b8YV.Enabled = True
        End Select
    End If

    RaiseEvent Change
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MinDate = m_def_MinDate
    m_MaxDate = Now
    m_ViewIndex = m_def_ViewIndex

    
End Sub

Private Sub UserControl_Paint()
    Dim i As Integer
        
    For i = 0 To optDate.UBound
        If optDate(i).BackColor <> UserControl.BackColor Then
            optDate(i).BackColor = UserControl.BackColor
        End If
    Next
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_MinDate = PropBag.ReadProperty("MinDate", m_def_MinDate)
    m_MaxDate = PropBag.ReadProperty("MaxDate", Now)
    m_ViewIndex = PropBag.ReadProperty("ViewIndex", m_def_ViewIndex)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    dtpByDate.Width = GetWidth - dtpByDate.Left
    b8MV.Width = GetWidth - b8MV.Left
    b8YV.Width = GetWidth - b8YV.Left
    Err.Clear
End Sub

Private Sub UserControl_Show()
    Dim i As Integer
        
    For i = 0 To optDate.UBound
        optDate(i).BackColor = UserControl.BackColor
    Next

    dtpByDate.Value = Now
    b8MV.MonthVal = Month(Now) - 1
    b8MV.YearVal = Year(Now)
    b8YV.YearVal = Year(Now)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MinDate", m_MinDate, m_def_MinDate)
    Call PropBag.WriteProperty("MaxDate", m_MaxDate, Now)
    Call PropBag.WriteProperty("ViewIndex", m_ViewIndex, m_def_ViewIndex)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ViewIndex() As Integer
    ViewIndex = m_ViewIndex
End Property

Public Property Let ViewIndex(ByVal New_ViewIndex As Integer)
    
    If New_ViewIndex <= optDate.UBound Then
        m_ViewIndex = New_ViewIndex
        optDate(New_ViewIndex).Value = True
        PropertyChanged "ViewIndex"
    End If
End Property

