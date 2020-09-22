VERSION 5.00
Begin VB.UserControl b83DRect 
   BackColor       =   &H00F5F5F5&
   CanGetFocus     =   0   'False
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   ClipBehavior    =   0  'None
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   Windowless      =   -1  'True
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   68
      X2              =   68
      Y1              =   66
      Y2              =   34
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   20
      X2              =   60
      Y1              =   68
      Y2              =   68
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   6
      X2              =   6
      Y1              =   34
      Y2              =   56
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   20
      X2              =   60
      Y1              =   32
      Y2              =   32
   End
End
Attribute VB_Name = "b83DRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function


Private Sub UserControl_Resize()
        
    Dim iH As Integer
    Dim iW As Integer
    
    iW = GetWidth
    iH = GetHeight
    
    On Error Resume Next
    
    With Line1
        .X1 = 0
        .X2 = 0
        .Y1 = 0
        .Y2 = iH
    End With
    
    With line2
        .X1 = 0
        .X2 = iW
        .Y1 = 0
        .Y2 = 0
    End With
    
    With Line3
        .X1 = 0
        .X2 = iW
        .Y1 = iH - 1
        .Y2 = iH - 1
    End With
    
    With Line4
        .X1 = iW - 1
        .X2 = iW - 1
        .Y1 = 0
        .Y2 = iH
    End With

    Err.Clear
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line1,Line1,-1,BorderColor
Public Property Get Color1() As OLE_COLOR
Attribute Color1.VB_Description = "Returns/sets the color of an object's border."
    Color1 = Line1.BorderColor
End Property

Public Property Let Color1(ByVal New_Color1 As OLE_COLOR)
    Line1.BorderColor() = New_Color1
    PropertyChanged "Color1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line2,Line2,-1,BorderColor
Public Property Get Color2() As OLE_COLOR
Attribute Color2.VB_Description = "Returns/sets the color of an object's border."
    Color2 = line2.BorderColor
End Property

Public Property Let Color2(ByVal New_Color2 As OLE_COLOR)
    line2.BorderColor() = New_Color2
    PropertyChanged "Color2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line3,Line3,-1,BorderColor
Public Property Get Color3() As OLE_COLOR
Attribute Color3.VB_Description = "Returns/sets the color of an object's border."
    Color3 = Line3.BorderColor
End Property

Public Property Let Color3(ByVal New_Color3 As OLE_COLOR)
    Line3.BorderColor() = New_Color3
    PropertyChanged "Color3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line4,Line4,-1,BorderColor
Public Property Get Color4() As OLE_COLOR
Attribute Color4.VB_Description = "Returns/sets the color of an object's border."
    Color4 = Line4.BorderColor
End Property

Public Property Let Color4(ByVal New_Color4 As OLE_COLOR)
    Line4.BorderColor() = New_Color4
    PropertyChanged "Color4"
End Property

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

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Line1.BorderColor = PropBag.ReadProperty("Color1", -2147483640)
    line2.BorderColor = PropBag.ReadProperty("Color2", -2147483640)
    Line3.BorderColor = PropBag.ReadProperty("Color3", -2147483640)
    Line4.BorderColor = PropBag.ReadProperty("Color4", -2147483640)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Color1", Line1.BorderColor, -2147483640)
    Call PropBag.WriteProperty("Color2", line2.BorderColor, -2147483640)
    Call PropBag.WriteProperty("Color3", Line3.BorderColor, -2147483640)
    Call PropBag.WriteProperty("Color4", Line4.BorderColor, -2147483640)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub

