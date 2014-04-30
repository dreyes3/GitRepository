Attribute VB_Name = "AutoExpandDropDown"
Option Explicit

Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const DT_CALCRECT = &H400

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function SendMessageLong Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function DrawText Lib "user32" Alias _
    "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
    ByVal nCount As Long, lpRect As RECT, ByVal wFormat _
    As Long) As Long

Public Function AutoSizeDropDownWidth(Combo As Object) As Boolean
    '**************************************************************
    'PURPOSE: Automatically size the combo box drop down width
    '         based on the width of the longest item in the combo box
    
    'PARAMETERS: Combo - ComboBox to size
    
    'RETURNS: True if successful, false otherwise
    
    'ASSUMPTIONS: 1. Form's Scale Mode is vbTwips, which is why
    '                conversion from twips to pixels are made.
    '                API functions require units in pixels
    '
    '             2. Combo Box's parent is a form or other
    '                container that support the hDC property
    
    'EXAMPLE: AutoSizeDropDownWidth Combo1
    '****************************************************************
    Dim lRet As Long
    Dim bAns As Boolean
    Dim lCurrentWidth As Single
    Dim rectCboText As RECT
    Dim lParentHDC As Long
    Dim lListCount As Long
    Dim lCtr As Long
    Dim lTempWidth As Long
    Dim lWidth As Long
    Dim sSavedFont As String
    Dim sngSavedSize As Single
    Dim bSavedBold As Boolean
    Dim bSavedItalic As Boolean
    Dim bSavedUnderline As Boolean
    Dim bFontSaved As Boolean
    
    On Error GoTo ErrorHandler
    
    If Not TypeOf Combo Is ComboBox Then Exit Function
    lParentHDC = Combo.Parent.hdc
    If lParentHDC = 0 Then Exit Function
    lListCount = Combo.ListCount
    If lListCount = 0 Then Exit Function
    
    
    'Change font of parent to combo box's font
    'Save first so it can be reverted when finished
    'this is necessary for drawtext API Function
    'which is used to determine longest string in combo box
    With Combo.Parent
    
        sSavedFont = .FontName
        sngSavedSize = .FontSize
        bSavedBold = .FontBold
        bSavedItalic = .FontItalic
        bSavedUnderline = .FontUnderline
        
        .FontName = Combo.FontName
        .FontSize = Combo.FontSize
        .FontBold = Combo.FontBold
        .FontItalic = Combo.FontItalic
        .FontUnderline = Combo.FontItalic
    
    End With
    
    bFontSaved = True
    
    'Get the width of the largest item
    For lCtr = 0 To lListCount
       DrawText lParentHDC, Combo.List(lCtr), -1, rectCboText, _
            DT_CALCRECT
       'adjust the number added (20 in this case to
       'achieve desired right margin
       lTempWidth = rectCboText.Right - rectCboText.Left + 20
    
       If (lTempWidth > lWidth) Then
          lWidth = lTempWidth
       End If
    Next
    lWidth = lWidth + 20
    lCurrentWidth = SendMessageLong(Combo.hwnd, CB_GETDROPPEDWIDTH, _
        0, 0)
    
    If lCurrentWidth > lWidth Then 'current drop-down width is
    '                               sufficient
    
        AutoSizeDropDownWidth = True
        GoTo ErrorHandler
        Exit Function
    End If
     
    'don't allow drop-down width to
    'exceed screen.width
     
    If lWidth > Screen.Width \ Screen.TwipsPerPixelX - 20 Then _
       lWidth = Screen.Width \ Screen.TwipsPerPixelX - 20
    
    lRet = SendMessageLong(Combo.hwnd, CB_SETDROPPEDWIDTH, lWidth, 0)
    
    AutoSizeDropDownWidth = lRet > 0
ErrorHandler:
    On Error Resume Next
    If bFontSaved Then
    'restore parent's font settings
      With Combo.Parent
        .FontName = sSavedFont
        .FontSize = sngSavedSize
        .FontUnderline = bSavedUnderline
        .FontBold = bSavedBold
        .FontItalic = bSavedItalic
     End With
    End If
End Function



