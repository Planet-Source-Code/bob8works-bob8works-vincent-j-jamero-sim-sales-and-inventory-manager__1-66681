VERSION 5.00
Begin VB.UserControl b8SBCenter 
   Alignable       =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Image imgRight 
      Height          =   1605
      Left            =   3900
      MousePointer    =   9  'Size W E
      Picture         =   "b8SBCenter.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   60
   End
   Begin VB.Image imgLeft 
      Height          =   1605
      Left            =   0
      Picture         =   "b8SBCenter.ctx":004E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   60
   End
End
Attribute VB_Name = "b8SBCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim bMouseDown As Boolean
Dim iMX As Integer

'events
Public Event Resize()
Public Event BeforeResize(ByVal NewWidth As Integer)
'Default Property Values:
Const m_def_MinWidth = 160
'Property Variables:
Dim m_MinWidth As Integer




Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function


Private Sub imgRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMouseDown = True
    iMX = X
End Sub

Private Sub imgRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim iNewWidth As Integer
    
    If bMouseDown = True Then
            
        iNewWidth = UserControl.Width + (X - iMX)
        
        If iNewWidth >= m_MinWidth * Screen.TwipsPerPixelX Then
            RaiseEvent BeforeResize(iNewWidth)
            UserControl.Width = iNewWidth
        Else
            'bMouseDown = False
        End If
        'RaiseEvent Resize
        
    End If
    
    
End Sub

Private Sub imgRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMouseDown = False
End Sub

Private Sub UserControl_Resize()
    
    If GetWidth < m_MinWidth Then
        UserControl.Width = m_MinWidth * Screen.TwipsPerPixelX
    End If

    imgLeft.Move 0, 0, imgLeft.Width, GetHeight
    imgRight.Move GetWidth - imgRight.Width, 0, imgRight.Width, GetHeight

    RaiseEvent Resize
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,160
Public Property Get MinWidth() As Integer
    MinWidth = m_MinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Integer)
    m_MinWidth = New_MinWidth
    PropertyChanged "MinWidth"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MinWidth = m_def_MinWidth
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_MinWidth = PropBag.ReadProperty("MinWidth", m_def_MinWidth)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("MinWidth", m_MinWidth, m_def_MinWidth)
End Sub

