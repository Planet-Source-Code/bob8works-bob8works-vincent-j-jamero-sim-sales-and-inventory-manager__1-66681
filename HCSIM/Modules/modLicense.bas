Attribute VB_Name = "modLicense"
Option Explicit


Private Const sLK As String = "bob8workssimferrermarketing"

Public Function CheckLicense() As Boolean

    Dim sKey As String
    Dim sNewKey As String
        
    
    'read registry
    sKey = GetSetting("bob8workssim", "license", "key", "")
    
    
    If sKey <> sLK Then
        
        'invalid
        
        'prompt for key
        sNewKey = frmLicense.ShowForm
        
        'save setting
        SaveSetting "bob8workssim", "license", "key", sNewKey
        MsgBox "Please restart Sales and Inventory Manager application.", vbInformation
        'set flag
        CheckLicense = False
    Else
        CheckLicense = True
    End If
    
End Function
