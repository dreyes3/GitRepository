Attribute VB_Name = "TreeView"
'===========Bas module code=============
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1

Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

  Const NM_CUSTOMDRAW = (-12&)
  Const WM_NOTIFY As Long = &H4E&
  Const WM_SETREDRAW = &HB

  Const ODS_SELECTED = &H1
  Const COLOR_WINDOWBACKGROUND = 9
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_HIGHLIGHTTEXT = 14

  Const CDDS_PREPAINT As Long = &H1&
  Const CDDS_POSTPAINT As Long = &H2&
  Const CDDS_PREERASE As Long = &H3&
  Const CDDS_POSTERASE As Long = &H4&
  Const CDDS_SUBITEM As Long = &H20000
  
  Const CDRF_DODEFAULT = &H0&
  Const CDRF_NEWFONT As Long = &H2&
  Const CDRF_NOTIFYPOSTPAINT As Long = &H10&
  Const CDRF_NOTIFYITEMDRAW As Long = &H20&
  Const CDRF_NOTIFYSUBITEMDRAW As Long = &H20&
  Const CDRF_NOTIFYPOSTERASE As Long = &H40&
  Const CDRF_NOTIFYITEMERASE As Long = &H80&
  Const CDDS_ITEM As Long = &H10000
  Const CDDS_ITEMPREPAINT As Long = CDDS_ITEM Or CDDS_PREPAINT
  Const CDDS_ITEMPOSTPAINT As Long = CDDS_ITEM Or CDDS_POSTPAINT
  Const CDDS_ITEMPREERASE As Long = CDDS_ITEM Or CDDS_PREERASE
  Const CDDS_ITEMPOSTERASE As Long = CDDS_ITEM Or CDDS_POSTERASE

  Type NMHDR
    hWndFrom As Long      ' Window handle of control sending message
    idFrom As Long        ' Identifier of control sending message
    code  As Long         ' Specifies the notification code
  End Type
  
  ' sub struct of the NMCUSTOMDRAW struct
  Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type
  
  ' generic customdraw struct
  Public Type NMCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hdc As Long
    rc As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemlParam As Long
  End Type
  
  ' treeview specific customdraw struct
  Public Type NMTVCUSTOMDRAW
    nmcd As NMCUSTOMDRAW
    clrText As Long
    clrTextBk As Long
    ' if IE >= 4.0 this member of the struct can be used
    iLevel As Integer
  End Type
  
  Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
  End Type
  
  Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
  End Type
  



Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (pDest As Any, ByVal dwLength As Long, ByVal bFill As Byte)

Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hwnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const TV_FIRST = &H1100&
Const TVM_SETITEMA = (TV_FIRST + 13)
Private Const TVM_GETITEMRECT = (TV_FIRST + 4)
Public Const TVM_SETTOOLTIPS = (TV_FIRST + 24)
Const TVM_ENSUREVISIBLE = (TV_FIRST + 20)
Const TVM_SETITEMHEIGHT = (TV_FIRST + 27)
Public Const TVM_GETITEMHEIGHT = (TV_FIRST + 28)

Const TVIF_INTEGRAL = &H80
Const TVIF_HANDLE = &H10

Public Const TVS_NONEVENHEIGHT = &H4000
Public Const TVS_NOSCROLL = &H2000
Public Const TVS_NOHSCROLL = &H8000
Public Const TVS_NOTOOLTIPS = &H80
Public Const TVS_SINGLEEXPAND = &H400
Public Const TVS_HASLINES As Long = 2

Public Type TVITEM   ' was TV_ITEM
  mask As Long
  hItem As Long
  state As Long
  stateMask As Long
  pszText As Long   ' pointer
  cchTextMax As Long
  iImage As Long
  iSelectedImage As Long
  cChildren As Long
  lParam As Long
End Type

Public Type TVITEMEX
  mask As Long
  hItem As Long
  state As Long
  stateMask As Long
  pszText As Long   ' pointer
  cchTextMax As Long
  iImage As Long
  iSelectedImage As Long
  cChildren As Long
  lParam As Long
  iIntegral As Long
End Type

Private Declare Function InvalidateRectByNum Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function ValidateRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Declare Function ValidateRectBynum& Lib "user32" Alias "ValidateRect" (ByVal hwnd As Long, ByVal lpRect As Long)

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_WNDPROC As Long = (-4&)
Private Const WS_HSCROLL = &H100000

Public OldProc As Long
Public LineHeight As Long
Dim bFromCode As Boolean
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByValhBrush As Long) As Long

Public Function WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim style As Long, nd As node, hgt_old As Long, hgt_new As Long
  Dim nLines As Long
  Dim hItem As Long
  Dim lBrushColor As Long
  Dim rc As RECT
  Dim tvex As TVITEMEX
  Dim lTemp As Long
  Dim TextSize As POINTAPI
  Dim intTop As Integer
  Dim intBottom As Integer
  Dim intLeft As Integer
  Dim intRight As Integer
  Dim strKeyword As String
  Dim strTemp() As String
  Static i As Integer
  Dim j As Integer
  Dim k As Integer
  
  Select Case iMsg
    Case WM_NOTIFY
      Dim udtNMHDR As NMHDR
      CopyMemory udtNMHDR, ByVal lParam, 12&
      With udtNMHDR
        If .code = NM_CUSTOMDRAW Then
          Dim udtNMTVCUSTOMDRAW As NMTVCUSTOMDRAW
          CopyMemory udtNMTVCUSTOMDRAW, ByVal lParam, Len(udtNMTVCUSTOMDRAW)
          With udtNMTVCUSTOMDRAW.nmcd
            Select Case .dwDrawStage
              Case CDDS_PREPAINT
                  If LineHeight = 0 Then
                      LineHeight = SendMessage(.hdr.hWndFrom, TVM_GETITEMHEIGHT, 0, ByVal 0&)
                  End If
                   WindowProc = CDRF_NOTIFYITEMDRAW Or CDRF_NOTIFYPOSTPAINT
                   Exit Function
              Case CDDS_ITEMPREPAINT
                   WindowProc = CDRF_NOTIFYPOSTPAINT Or CDRF_NEWFONT
                   Exit Function
              Case CDDS_ITEMPOSTPAINT
                   Set nd = GetNodeFromlParam(.lItemlParam)
                   hItem = .dwItemSpec
                   rc.Left = hItem
                   If SendMessage(.hdr.hWndFrom, TVM_GETITEMRECT, 1&, rc) = 0 Then
                      WindowProc = CDRF_DODEFAULT
                      Exit Function
                   End If
                   hgt_old = .rc.Bottom - .rc.Top
                   rc.Right = .rc.Right
                   SetTextColor .hdc, 255
                   Call DrawText(.hdc, nd.Text, Len(nd.Text), rc, DT_WORDBREAK Or DT_CALCRECT)
                   hgt_new = rc.Bottom - rc.Top
                   If nd.Parent Is Nothing Then
                      Exit Function
                   End If
                   'If hgt_new <= LineHeight Then
                   '   WindowProc = CDRF_DODEFAULT
                   '   Exit Function
                   'End If
                   If hgt_new > hgt_old Then
                      nLines = Int(hgt_new / LineHeight) + 1
                      tvex.hItem = hItem
                      tvex.iIntegral = nLines
                      tvex.mask = TVIF_INTEGRAL Or TVIF_HANDLE
                      SendMessage .hdr.hWndFrom, TVM_SETITEMA, 0, tvex
                      On Error Resume Next
                      If nd.Parent.Expanded = True Then
                         SendMessage .hdr.hWndFrom, WM_SETREDRAW, True, ByVal 0&
                         nd.Parent.Expanded = False
                         nd.Parent.Expanded = True
                      End If
                      On Error GoTo 0
                   End If
                   With udtNMTVCUSTOMDRAW
                        SetBkColor .nmcd.hdc, .clrTextBk
                        SetTextColor .nmcd.hdc, .clrText
                        If .nmcd.uItemState And ODS_SELECTED Then
                           lBrushColor = COLOR_HIGHLIGHT + 1
                        Else
                           lBrushColor = COLOR_WINDOWBACKGROUND + 1
                        End If
                   End With
                   i = i + 1
                   rc.Right = .rc.Right
                   FillRect .hdc, rc, lBrushColor
                   Call DrawText(.hdc, nd.Text, -1, rc, DT_WORDBREAK)
                   intTop = rc.Top
                   intBottom = rc.Bottom
                   intLeft = rc.Left
                   intRight = rc.Right
                   'Highlight Keyword
                   For j = 0 To UBound(gstrKeyword)
                        strKeyword = gstrKeyword(j, 0)
                        If InStr(1, nd.Text, strKeyword) > 0 Then
                           strTemp = Split(strKeyword, " ")
                           For intCount = 0 To UBound(strTemp)
                               highlightKeyword .hdc, nd.Text, rc, strTemp(intCount), intTop, intBottom, intLeft, intRight, gstrKeyword(j, 1), gstrKeyword(j, 2), gstrKeyword(j, 3)
                           Next
                        End If
                   Next
                   WindowProc = CDRF_NOTIFYPOSTPAINT
                   Exit Function
                Case CDDS_POSTPAINT
                   If bFromCode Then
                      WindowProc = 4&
                      bFromCode = False
                      Exit Function
                   End If
                   bFromCode = True
                   SetWindowLong .hdr.hWndFrom, GWL_STYLE, GetWindowLong(.hdr.hWndFrom, GWL_STYLE) And Not WS_HSCROLL
                   SetWindowPos .hdr.hWndFrom, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
                   SendMessage .hdr.hWndFrom, WM_SETREDRAW, True, ByVal 0&
             End Select
          End With
        End If
      End With
  End Select
  WindowProc = CallWindowProc(OldProc, hwnd, iMsg, wParam, lParam)
End Function

Function GetNodeFromlParam(ByVal lParam As Long) As node
  Dim pNode As Long
  Dim nod As node
  If lParam Then
    CopyMemory pNode, ByVal lParam + 8, 4
    If pNode Then
      CopyMemory nod, pNode, 4
      Set GetNodeFromlParam = nod
      FillMemory nod, 4, 0
    End If
  End If
End Function

Private Function hItemFromNode(ByVal nod As node) As Long
   CopyMemory hItemFromNode, ByVal (ObjPtr(nod) + 68), 4&
End Function

Public Function TrimNulls(sTemp As String) As String
   Dim l As Long
   l = InStr(1, sTemp, Chr(0))
   If l = 1 Then
      TrimNulls = ""
   ElseIf l > 0 Then
      TrimNulls = Left$(sTemp, l - 1)
   Else
      TrimNulls = sTemp
   End If
End Function

Private Sub highlightKeyword(ByVal lngHDC As Long, ByVal strNode As String, ByRef rc As RECT, ByVal strKeyword As String, ByVal intTop As Integer, ByVal intBottom As Integer, ByVal intLeft As Integer, ByVal intRight As Integer, ByVal intRed As Integer, ByVal intGreen As Integer, ByVal intBlue As Integer)
    
    Dim i  As Integer
    Dim j As Double
    Dim k As Integer
    Dim intRow As Integer
    Dim strTemp As String
    
    k = InStr(1, strNode, strKeyword)
    j = 0
    If k > 0 Then
       If Len(strNode) <> Len(strKeyword) Then
          'Find Whole Word only
          If k - 1 > 0 Then
             If Mid(strNode, k - 1, 1) <> " " Then
                Exit Sub
             End If
          End If
          If k + Len(strKeyword) <= Len(strNode) Then
             If Mid(strNode, k + Len(strKeyword), 1) <> " " Then
                Exit Sub
             End If
          End If
       End If
       
       strTemp = Mid(strNode, 1, k + Len(strKeyword) - 1)
       rc.Top = intTop
       rc.Bottom = intBottom
       rc.Left = intLeft
       rc.Right = intRight
       Call DrawText(lngHDC, strTemp, Len(strTemp), rc, DT_WORDBREAK Or DT_CALCRECT)
       i = (rc.Bottom - rc.Top) / 13
       If i > 1 Then
          'Set top and bottom if is a break line case
          intRow = i
          strTemp = Mid(strNode, 1, k - 1)
          strTemp = findPrefix(lngHDC, strTemp, rc, intRow)
          
       Else
          intRow = 1
          strTemp = Mid(strNode, 1, k - 1)
       End If
       If strTemp <> "" Then
          Call DrawText(lngHDC, strTemp, Len(strTemp), rc, DT_CALCRECT)
          i = rc.Right - rc.Left
       Else
          i = 0
       End If
       Call DrawText(lngHDC, strKeyword, Len(strKeyword), rc, DT_CALCRECT)
       rc.Left = rc.Left + i
       rc.Right = rc.Left + rc.Right
       SetBkColor lngHDC, RGB(intRed, intGreen, intBlue)
       SetTextColor lngHDC, RGB(255, 255, 255)
       If intRow > 1 Then
          rc.Top = rc.Top + ((intRow - 1) * 13)
          rc.Bottom = rc.Top + 13
       End If
       Call DrawText(lngHDC, strKeyword, Len(strKeyword), rc, DT_WORDBREAK)
    End If
End Sub

Private Function findPrefix(ByVal lngHDC As Long, ByVal strNode As String, ByRef rc As RECT, ByVal intRow As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    Dim strSplit() As String
    
    strTemp = Trim(strNode)
    Call DrawText(lngHDC, strTemp, Len(strTemp), rc, DT_WORDBREAK Or DT_CALCRECT)
    Do While (rc.Bottom - rc.Top) / 13 = intRow
        'Loop to get the prefix string
        strSplit = Split(strTemp, " ")
        strTemp = ""
        For i = 0 To UBound(strSplit) - 1
            strTemp = strTemp & IIf(strTemp = "", "", " ") & strSplit(i)
        Next
        Call DrawText(lngHDC, strTemp, Len(strTemp), rc, DT_WORDBREAK Or DT_CALCRECT)
    Loop
    findPrefix = Mid(strNode, Len(strTemp) + 1)
    If Trim(findPrefix) = "" Then
       findPrefix = ""
    Else
       findPrefix = Trim(findPrefix) & " "
    End If
End Function

