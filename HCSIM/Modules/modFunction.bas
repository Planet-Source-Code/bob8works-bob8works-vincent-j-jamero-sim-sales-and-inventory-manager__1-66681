Attribute VB_Name = "modFunction"
'' Module Function
'' Code By: Vincent J. Jamero

Option Explicit

Public Enum FindOptions
    PartOfWord = 0
    MatchCase = 1
    WholeWordOnly = 3
End Enum

'API for opening a browser
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hWnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long



Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long


Public Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Sub FormDrag(frmName As Form) 'procedure to drag a no-titlebar form
    ReleaseCapture
    Call SendMessage(frmName.hWnd, &HA1, 2, 0&)
End Sub



Public Function MakeGradient(ByRef frm As Object, Scheme As Integer)
    Dim cr(255) As Integer
    Dim cG(255) As Integer
    Dim cB(255) As Integer
    Dim d As Double
    Dim i As Integer
    
    
    Select Case Scheme
        Case 1
            For i = 0 To 255
                cr(i) = 255 - (i * 0.2)
                cG(i) = 255 - (i * 0.2)
                cB(i) = 255 - (i * 0.2)
            Next
    End Select
    

    frm.ScaleMode = vbPixels
    d = frm.ScaleHeight / 255
    frm.DrawWidth = d + 1
    For i = 0 To 255
        frm.ForeColor = RGB(cr(i), cG(i), cB(i))
        frm.Line (0, i * d)-(frm.ScaleWidth, i * d)
    Next
    'Frm.AutoRedraw = True
End Function











Public Function CheckTextBox(ByRef txt As Object, Optional sMSG As String = "TextBox", Optional sHowMSG As Boolean = True, Optional MinimumChar As Integer = 1) As Boolean
On Error Resume Next
    If Len(Trim(txt.Text)) < MinimumChar Then
        
        If sHowMSG Then
            MsgBox sMSG, vbExclamation
        End If
        
        txt.Text = ""
        txt.SetFocus
        
        CheckTextBox = False
    Else
        CheckTextBox = True
    End If
End Function

Public Function HLTxt(ByRef txt As Object)
On Error Resume Next
    txt.SelStart = 0
    txt.SelLength = Len(txt)
    txt.SetFocus
End Function


Public Function AddListItem(ByRef vListItem As ListView, sText As Variant, Optional imgIndex As Integer = 1, Optional sSubItem1 As Variant = "", Optional sSubItem2 As Variant = "", Optional sSubItem3 As Variant = "")
Dim lastIndex As Integer
    
On Error Resume Next
    
    lastIndex = vListItem.ListItems.Count + 1
    vListItem.ListItems.Add lastIndex, , sText, imgIndex, imgIndex
    If sSubItem1 <> "" Then _
        vListItem.ListItems(lastIndex).SubItems(1) = sSubItem1
        If sSubItem2 <> "" Then _
        vListItem.ListItems(lastIndex).SubItems(2) = sSubItem2
        If sSubItem3 <> "" Then _
        vListItem.ListItems(lastIndex).SubItems(3) = sSubItem3
End Function






Public Function cSentenceCase(sText As String) As String
    
    Dim splitText() As String
    Dim newWord As String
    Dim i As Integer
    
    'check if null---------------
    If Len(sText) < 1 Then
        cSentenceCase = ""
        Exit Function
    End If
    'end Null --------------------
    
    'convert
    sText = Trim(sText)
    
    splitText = Split(sText, " ")
    
    For i = 0 To UBound(splitText)
        If Len(Trim(splitText(i))) > 0 Then
            newWord = UCase(Left(Trim(splitText(i)), 1)) & LCase(Right(Trim(splitText(i)), Len(Trim(splitText(i))) - 1))
            cSentenceCase = cSentenceCase & " " & newWord
        End If
    Next
    
    cSentenceCase = Trim(cSentenceCase)
End Function



Public Function FillRecordToList(ByRef vRS As ADODB.Recordset, ByRef lv As ListView, sTableKey As String, Optional RecStartPos As Long = 0, Optional LimitCount As Long = 100, Optional WithID As Boolean = True, Optional WithIcon As Boolean = False)

    Dim i As Long
    Dim newColumnWidth As Integer
    Dim LimitCounter As Long
    Dim sCell As String
    Dim oldScaleMode As ScaleModeConstants
    
On Error Resume Next
    
    'minum fields must be 2
    If vRS.Fields.Count < 2 Then Exit Function
    
    'get old scale mode
    oldScaleMode = lv.Container.ScaleMode

    lv.Container.ScaleMode = vbTwips
    lv.ListItems.Clear
    
    
    If AnyRecordExisted(vRS) Then
        
        'create column headers
        For i = lv.ColumnHeaders.Count To vRS.Fields.Count - 2
            lv.ColumnHeaders.Add
        Next
        
        'set items
        vRS.Requery
        vRS.Move RecStartPos
        
        LimitCounter = 0
        
        
        While Not vRS.EOF
        
            
        
            'add
            If WithIcon = True Then
                If WithID = True Then
                    lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), vRS.Fields(1).Value, 1, 1
                Else
                    lv.ListItems.Add , , vRS.Fields(0).Value, 1, 1
                End If
            Else
                If WithID = True Then
                    lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), vRS.Fields(1).Value
                Else
                    lv.ListItems.Add , , vRS.Fields(0).Value
                End If
            End If
            'add sub items
            For i = 2 To vRS.Fields.Count - 1
           
                lv.ListItems(lv.ListItems.Count).SubItems(i - 1) = ReadField(vRS.Fields(i))
            Next
            
            vRS.MoveNext
            
            LimitCounter = LimitCounter + 1
            
            If LimitCounter >= LimitCount Then
                GoTo tagExitSub
            End If
            
        Wend
       
    End If
    
tagExitSub:
    'restore scale mode
    lv.Container.ScaleMode = oldScaleMode

End Function


Public Function FillRecordToListWN(ByRef vRS As ADODB.Recordset, ByRef lv As ListView, sTableKey As String, Optional RecStartPos As Long = 0, Optional LimitCount As Long = 100, Optional WithID As Boolean = True, Optional WithIcon As Boolean = False, Optional ColIndex As Integer = 0)

    Dim i As Long
    Dim newColumnWidth As Integer
    Dim LimitCounter As Long
    Dim sCell As String
    Dim oldScaleMode As ScaleModeConstants
    
    Dim CPOS As Long
    
On Error Resume Next
    
    'minum fields must be 2
    If vRS.Fields.Count < 2 Then Exit Function
    
    'get old scale mode
    oldScaleMode = lv.Container.ScaleMode

    lv.Container.ScaleMode = vbTwips
    lv.ListItems.Clear
    
    
    If AnyRecordExisted(vRS) Then
        
        'create column headers
        For i = lv.ColumnHeaders.Count To vRS.Fields.Count - 1
            lv.ColumnHeaders.Add
        Next
        
        'set items
        vRS.Requery
        vRS.Move RecStartPos
        
        LimitCounter = 0
        CPOS = RecStartPos
        
        While Not vRS.EOF
        
            CPOS = CPOS + 1
            
        
            'add
            If ColIndex = 0 Then
                
                If WithIcon = True Then
                    If WithID = True Then
                        lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), CPOS
                    Else
                        lv.ListItems.Add , , CPOS, 1, 1
                    End If
                Else
                    If WithID = True Then
                        lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), CPOS
                    Else
                        lv.ListItems.Add , , CPOS
                    End If
                End If
                
            
            Else
                
                If WithIcon = True Then
                    If WithID = True Then
                        lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), vRS.Fields(1).Value, 1, 1
                    Else
                        lv.ListItems.Add , , vRS.Fields(0).Value, 1, 1
                    End If
                Else
                    If WithID = True Then
                        lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), vRS.Fields(1).Value
                    Else
                        lv.ListItems.Add , , vRS.Fields(0).Value
                    End If
                End If
            
            End If
            
            
            
            
            'add sub items
            
            If ColIndex = 0 Then
            
                 For i = 1 To vRS.Fields.Count '- 1
                
                     lv.ListItems(lv.ListItems.Count).SubItems(i - 1) = ReadField(vRS.Fields(i - 1))
                 Next
                 
                 vRS.MoveNext
                 
                 LimitCounter = LimitCounter + 1
                 
                 If LimitCounter >= LimitCount Then
                     GoTo tagExitSub
                 End If
            

            Else
                Dim sbi As Integer
                sbi = 1
                For i = 2 To vRS.Fields.Count '- 1
                
                    If ColIndex = i Then
                        lv.ListItems(lv.ListItems.Count).SubItems(sbi) = CPOS
                        sbi = sbi + 1
                        lv.ListItems(lv.ListItems.Count).SubItems(sbi) = ReadField(vRS.Fields(i))
                    Else
                        lv.ListItems(lv.ListItems.Count).SubItems(sbi) = ReadField(vRS.Fields(i))
                    End If
                    
                    sbi = sbi + 1
                Next
                 
                vRS.MoveNext
                 
                LimitCounter = LimitCounter + 1
                 
                If LimitCounter >= LimitCount Then
                    GoTo tagExitSub
                End If
            
            End If
             
             
             
            
        Wend 'individual record
       
    End If
    
tagExitSub:
    'restore scale mode
    lv.Container.ScaleMode = oldScaleMode

End Function

Public Function SortLV(ByRef lv As ListView, Optional HeaderIndex As Integer = 0, Optional newSortOrder As ListSortOrderConstants = lvwAscending, Optional AutoOrder As Boolean = True)
    
    Dim lvHeader As ColumnHeader
    
    If AutoOrder = True Then
        If lv.SortOrder = lvwAscending Then
           lv.SortOrder = lvwDescending
        Else
           lv.SortOrder = lvwAscending
        End If
    Else
        lv.SortOrder = newSortOrder
    End If
    
    If HeaderIndex > lv.ColumnHeaders.Count - 1 Then
        HeaderIndex = 0
    End If
    
    lv.SortKey = HeaderIndex
    lv.Sorted = True
    lv.Refresh
    
    For Each lvHeader In lv.ColumnHeaders
        lvHeader.Icon = 0
    Next
    
    On Error Resume Next
    lv.ColumnHeaders(HeaderIndex + 1).Icon = lv.SortOrder + 1
    Err.Clear
End Function

Public Function UnSortLV(ByRef lv As ListView)
    
    Dim lvHeader As ColumnHeader
    
    lv.Sorted = False
    
    For Each lvHeader In lv.ColumnHeaders
        lvHeader.Icon = 0
    Next
End Function
    
    
    
    
    

Public Function GetLVKey(lvListItem As ListItem) As String
On Error GoTo errh:
    GetLVKey = Right(lvListItem.Key, Len(lvListItem.Key) - 4)
    Exit Function
errh:
    GetLVKey = ""
End Function

Public Function GetKey(sKey As String) As String
On Error GoTo errh:
    GetKey = Right(sKey, Len(sKey) - 4)
    Exit Function
errh:
    GetKey = ""
End Function
Public Function SetLVKey(sID As String, sTableKey As String) As String
    SetLVKey = Left(sTableKey, 4) & sID
End Function

Public Function FindLVItem(ByRef vLV As ListView, sCriteria As String, Optional iOption As FindOptions = 0, Optional MultiSelect As Boolean = False, Optional InverseSelection As Boolean = False, Optional FindNext As Boolean = False)

    Dim i As Integer
    Dim isFound As Boolean
    Dim li As Integer
    Dim StartPos As Integer
    
'On Error GoTo eh
    
    If vLV.ListItems.Count < 1 Then Exit Function

    If FindNext = True And vLV.SelectedItem.Index < vLV.ListItems.Count Then
        For li = 1 To vLV.SelectedItem.Index
            vLV.ListItems(li).Selected = False
        Next
        StartPos = vLV.SelectedItem.Index + 1
    Else
        For li = 1 To vLV.ListItems.Count
            vLV.ListItems(li).Selected = False
        Next
        StartPos = 1
    End If
    
    'set flag to default
    isFound = False
    
    For li = StartPos To vLV.ListItems.Count
        
        Select Case iOption
            
            Case FindOptions.PartOfWord  'normal

                If InStr(1, LCase(vLV.ListItems(li).Text), LCase(sCriteria)) > 0 Then
                                        
                    isFound = True

                Else

                    'check subitems
                    For i = 1 To vLV.ListItems(li).ListSubItems.Count
                        If InStr(1, LCase(vLV.ListItems(li).ListSubItems(i)), LCase(sCriteria)) > 0 Then
                            
                            isFound = True
                            Exit For
                        
                        End If
                    Next
                                        
                End If
                
            Case FindOptions.MatchCase  'match case
            
            Case FindOptions.WholeWordOnly  ' whole word only
                
            
        End Select
        
        
        
        
        If isFound Then
            
            vLV.ListItems(li).Selected = CBool(True - InverseSelection)
            vLV.ListItems(li).EnsureVisible
            
            If Not MultiSelect Then Exit For
        
        Else
            vLV.ListItems(li).Selected = CBool(False - InverseSelection)
        End If
        
    Next
    
    If FindNext = True And isFound = False And StartPos > 1 Then
        
        For li = 1 To StartPos
            
            Select Case iOption
                
                Case FindOptions.PartOfWord  'normal
    
                    If InStr(1, LCase(vLV.ListItems(li).Text), LCase(sCriteria)) > 0 Then
                                            
                        isFound = True
    
                    Else
    
                        'check subitems
                        For i = 1 To vLV.ListItems(li).ListSubItems.Count
                            If InStr(1, LCase(vLV.ListItems(li).ListSubItems(i)), LCase(sCriteria)) > 0 Then
                                
                                isFound = True
                                Exit For
                            
                            End If
                        Next
                                            
                    End If
                    
                Case FindOptions.MatchCase  'match case
                
                Case FindOptions.WholeWordOnly  ' whole word only
                    
                
            End Select
            
            
            
            
            If isFound Then
                
                vLV.ListItems(li).Selected = CBool(True - InverseSelection)
                vLV.ListItems(li).EnsureVisible
                
                If Not MultiSelect Then Exit For
            
            Else
                vLV.ListItems(li).Selected = CBool(False - InverseSelection)
            End If
            
        Next
    End If
'On Error Resume Next
Exit Function
eh:
    MsgBox Err.Description
    Resume Next
End Function


Public Function GetLVSelectedCount(ByRef lv As ListView) As Integer
    Dim i As Integer
    Dim iSelectedCount As Integer
    
    'default
    GetLVSelectedCount = 0
    
    'check if there is a record in the list
    If lv.ListItems.Count < 1 Then Exit Function
    
    
    iSelectedCount = 0
    For i = 1 To lv.ListItems.Count
        If lv.ListItems(i).Selected = True And Len(GetLVKey(lv.ListItems(i))) > 0 Then
            iSelectedCount = iSelectedCount + 1
        End If
    Next
    
    'return
    GetLVSelectedCount = iSelectedCount
End Function



Public Function CenterForm(ByRef frm As Form)
    frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
End Function

Public Sub OpenURL(urlADD As String, sourceHWND As Long)
     Call ShellExecute(sourceHWND, vbNullString, urlADD, "", vbNullString, 1)
End Sub


Public Function IsEmpty(s As String) As Boolean
    If Len(Trim(s)) < 1 Then
        IsEmpty = True
    Else
        IsEmpty = False
    End If
End Function



Public Function FindNoCharMatch(Str1 As String, Str2 As String) As Boolean
    
    Dim i As Integer
    Dim sC As String
    
    'default
    FindNoCharMatch = False

    'check the first stirng
    For i = 1 To Len(Str1)
        sC = Mid(Str1, i, 1)
    
        If InStr(1, Str2, sC) > 0 Then
            'found
            Exit Function
        End If
    Next
    
    'check the second stirng
    For i = 1 To Len(Str2)
        sC = Mid(Str2, i, 1)
    
        If InStr(1, Str1, sC) > 0 Then
            'found
            Exit Function
        End If
    Next
    
    
    'return success
    FindNoCharMatch = True
End Function



Public Function FillRecToIC(ByRef vRS As ADODB.Recordset, ByRef ICL As ImageCombo)
    
    On Error Resume Next
    
    ICL.ComboItems.Clear
    
    vRS.MoveFirst
    
    While vRS.EOF = False
        
        ICL.ComboItems.Add , SetLVKey(vRS.Fields(0).Value, "iclc"), vRS.Fields(1).Value
    
        vRS.MoveNext
    Wend
    
    ICL.ComboItems(1).Selected = True
    
    Err.Clear
    
End Function

Public Function GetKeyOnSplit(sKey As String, sDelimeter As String, Index As Integer) As String
    On Error GoTo errf:
    Dim sp() As String
    
    sp = Split(sKey, sDelimeter)
    
    If Index <= UBound(sp) Then
        GetKeyOnSplit = sp(Index)
    Else
        GetKeyOnSplit = ""
    End If
    Exit Function

errf:
    GetKeyOnSplit = ""
End Function

Public Function GetTxtVal(ByVal sTxt As String) As Double

    Dim sNew As String
    Dim sC As String
    Dim i As Integer
    
    'default
    GetTxtVal = 0
        
    sTxt = Trim(sTxt)
    
    If Len(sTxt) > 0 Then
        For i = 1 To Len(sTxt)
            sC = Mid(sTxt, i, 1)
            If sC = "-" Or sC = "." Or sC = "1" Or sC = "2" Or sC = "3" Or sC = "4" Or sC = "5" Or sC = "6" Or sC = "7" Or sC = "8" Or sC = "9" Or sC = "0" Then
                sNew = sNew & sC
            End If
        Next
    
        If Len(sNew) > 0 Then
            GetTxtVal = Val(sNew)
        End If
    End If
    
    
End Function


Public Function SetTextBoxValue(ByRef TxtBox As Variant)
    On Error Resume Next
    SaveSetting App.EXEName, "TextBoxValue", TxtBox.Name, TxtBox.Text
    Err.Clear
End Function

Public Function GetTextBoxValue(ByRef TxtBox As Variant, Optional sDefault As String = "")
    Dim sValue As String
    On Error Resume Next
    sValue = GetSetting(App.EXEName, "TextBoxValue", TxtBox.Name, "")
    
    If Len(Trim(sValue)) < 1 Then
        sValue = sDefault
    End If
    
    TxtBox.Text = sValue
    Err.Clear
End Function

Public Function GetDateDays(dDate As Date) As Long
    GetDateDays = 0 + CDate(FormatDateTime(dDate, vbShortDate))
End Function


'FillEmptyHoles
'Created: 1:26 AM   June 13, 2006
'Description: fills up empty cells
'
'Coder's situation:
'Im thinking about Ruby on how she fills up the missing part of my life
'and my stomach says "I need to be fill up too." hehe!!!
Public Function FillEmptyHoles(ByRef lv As ListView, Optional sSTR As String = " ")
    
    Dim i As Integer
    Dim x As Integer
    
    For i = 1 To lv.ListItems.Count
        If IsEmpty(lv.ListItems(i).Text) = True Then
            lv.ListItems(i).Text = sSTR
        End If
        For x = 1 To lv.ColumnHeaders.Count - 1
            If IsEmpty(lv.ListItems(i).SubItems(x)) = True Then
                lv.ListItems(i).SubItems(x) = sSTR
            End If
        Next
    Next
    
End Function


'Get Header left
'Created: 10:28 PM   June 14, 2006
'Description: get the listview column header left
'
'Coder's Situation:
'I'm thinking about vb6, I guess it's dead. I have to learn c#, hehe...
Public Function GetHeaderLeft(ByRef lv As ListView, ByVal ColumnHeaderIndex As Integer)

    Dim i As Integer
    Dim x As Integer
    
    If ColumnHeaderIndex > lv.ColumnHeaders.Count Or ColumnHeaderIndex < 2 Then
        GetHeaderLeft = 0
        Exit Function
    End If
    
    x = 0
    For i = 2 To ColumnHeaderIndex
        x = x + lv.ColumnHeaders(i - 1).Width
    Next
    
    GetHeaderLeft = x
    
End Function

'Not In List
'Created: 10:16 AM   June 23, 2006
'Description: Check if the selected item is valid
'
'Coder's Situation:
'???
Public Function ItemNotInList(ByRef ic As ImageCombo) As Boolean
    
    Dim tmpT As String
    
    'default
    ItemNotInList = True
    
    On Error GoTo RAE
    
    tmpT = ic.SelectedItem.Text
    ItemNotInList = False
RAE:
End Function




'Requires Class: clsGrad
Public Sub PaintGrad(ByRef Obj As Object, lColor1 As Long, lColor2 As Long, iAngle As Integer)
    Dim cGrad As New clsGrad
    On Error Resume Next
    Obj.AutoRedraw = True
    cGrad.Color1 = lColor1
    cGrad.Color2 = lColor2
    cGrad.Angle = iAngle
    cGrad.Draw Obj
    Obj.Refresh
    Set cGrad = Nothing
    Err.Clear
End Sub

Public Function GetRSec(ByVal dDate As Date) As Date

   GetRSec = DateValue(dDate) + (Hour(Now) / 24) + (Minute(Now)) / 3600 + (Second(Now) / 216000)
End Function

Public Function ComNumZ(ByVal vVal As Variant, ByVal iWidth As Integer) As String
    If Len(Trim(vVal)) > iWidth Then
        ComNumZ = CStr(vVal)
    Else
        ComNumZ = String$(iWidth - Len(Trim(vVal)), "0") & Trim(vVal)
    End If
End Function

