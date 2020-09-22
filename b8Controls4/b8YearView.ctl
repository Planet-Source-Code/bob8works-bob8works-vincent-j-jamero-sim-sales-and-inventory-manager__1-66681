VERSION 5.00
Begin VB.UserControl b8YearView 
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   97
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "b8YearView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long


'Default Property Values:
Const m_def_Min = 1900
Const m_def_Max = 2100
Const m_def_YearVal = 0
'Property Variables:
Dim m_Min As Integer
Dim m_Max As Integer
Dim m_YearVal As Integer
'Event Declarations:
Event Click() 'MappingInfo=cmbYear,cmbYear,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=cmbYear,cmbYear,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=cmbYear,cmbYear,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=cmbYear,cmbYear,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event Change()



Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function


Private Sub cmbYear_Change()
    If IsNumeric(cmbYear.Text) = True Then
        YearVal = Val(cmbYear.Text)
    End If
End Sub

Private Sub cmbYear_Validate(Cancel As Boolean)
    If cmbYear.ListIndex < 0 Then
        Cancel = True
    End If
End Sub

Private Sub UserControl_Initialize()
    InitCommonControls

    YearVal = Year(Now)
    Refresh
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    cmbYear.Width = GetWidth
    If GetHeight < cmbYear.Height Then
        UserControl.Height = cmbYear.Height * Screen.TwipsPerPixelY
    End If
    
    cmbYear.SelStart = 0
    cmbYear.SelLength = 0
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbYear,cmbYear,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = cmbYear.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    cmbYear.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    cmbYear.Enabled = New_Enabled
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbYear,cmbYear,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = cmbYear.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set cmbYear.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbYear,cmbYear,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Dim i As Integer
    
    cmbYear.Clear

    For i = Min To Max
        cmbYear.AddItem CStr(i)
    Next
    
    
End Sub

Private Sub cmbYear_Click()
    If IsNumeric(cmbYear.Text) = True Then
        YearVal = Val(cmbYear.Text)
    End If
    RaiseEvent Click
End Sub

Private Sub cmbYear_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cmbYear_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub cmbYear_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get Min() As Integer
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    If New_Min < 1 Or New_Min > Max Then
        Exit Property
    End If
    m_Min = New_Min
    If YearVal < New_Min Then
        YearVal = New_Min
    End If
    Refresh
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,9999
Public Property Get Max() As Integer
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    If New_Max < Min Or New_Max > 9999 Then
        Exit Property
    End If
    If YearVal > New_Max Then
        YearVal = New_Max
    End If
    m_Max = New_Max
    Refresh
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get YearVal() As Integer
    YearVal = Val(cmbYear.Text)
End Property

Public Property Let YearVal(ByVal New_YearVal As Integer)
    Dim i As Integer
    
    If New_YearVal < Min Or New_YearVal > Max Then
        Exit Property
    End If
    
    
    For i = 0 To Max - Min
        If Val(cmbYear.List(i)) = New_YearVal Then
            cmbYear.ListIndex = i
            Exit For
        End If
    Next


    m_YearVal = New_YearVal
    PropertyChanged "YearVal"
    If Me.Enabled = True Then
        RaiseEvent Change
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbYear,cmbYear,-1,Text
Public Property Get YearToString() As String
Attribute YearToString.VB_Description = "Returns/sets the text contained in the control."
    YearToString = cmbYear.Text
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_YearVal = m_def_YearVal
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    cmbYear.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set cmbYear.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_YearVal = PropBag.ReadProperty("YearVal", m_def_YearVal)
    cmbYear.Text = PropBag.ReadProperty("YearToString", "")
End Sub

Private Sub UserControl_Show()
    Refresh
    If YearVal = 0 Then
        YearVal = Year(Now)
    End If
    cmbYear.SelStart = 1
    cmbYear.SelLength = 0
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", cmbYear.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", cmbYear.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("YearVal", m_YearVal, m_def_YearVal)
    Call PropBag.WriteProperty("YearToString", cmbYear.Text, "")
End Sub

