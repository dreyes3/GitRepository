Attribute VB_Name = "modMsgBox"


Option Explicit
'--------------------API--------------------
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'??API?MessageBox??VB???MsgBox
Private Declare Function MessageBox Lib "user32" _
   Alias "MessageBoxA" _
  (ByVal hWnd As Long, _
   ByVal lpText As String, _
   ByVal lpCaption As String, _
   ByVal wType As Long) As Long
   
Private Declare Function SetWindowsHookEx Lib "user32" _
   Alias "SetWindowsHookExA" _
  (ByVal idHook As Long, _
   ByVal lpfn As Long, _
   ByVal hmod As Long, _
   ByVal dwThreadId As Long) As Long
   
Private Declare Function UnhookWindowsHookEx Lib "user32" _
   (ByVal hHook As Long) As Long

Private Declare Function MoveWindow Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long
   
Private Declare Function GetWindowRect Lib "user32" _
  (ByVal hWnd As Long, _
   lpRect As RECT) As Long
   
Public Declare Function GetDlgItem Lib "user32" _
  (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
  
Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" _
  (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long

Private Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" _
  (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long


Private hHook As Long


'Windows-defined Return values. The return
'values and control IDs are identical.
Private Const IDOK = 1
Private Const IDCANCEL = 2
Private Const IDABORT = 3
Private Const IDRETRY = 4
Private Const IDIGNORE = 5
Private Const IDYES = 6
Private Const IDNO = 7
Private Const IDPROMPT = &HFFFF&

'----------------------form handle----------------------'
Private hFormhWnd As Long
'----------------------Class variable-------------------'
Private mvarOKMsg As String 'local copy
Private mvarCANCELMsg As String 'local copy
Private mvarABORTMsg As String 'local copy
Private mvarRETRYMsg  As String 'local copy
Private mvarIGNOREMsg As String 'local copy
Private mvarYESMsg As String 'local copy
Private mvarNOMsg As String 'local copy

Public Property Let OKMsg(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.

    mvarOKMsg = vData
End Property

Public Property Let CANCELMsg(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.

    mvarCANCELMsg = vData
End Property

Public Property Let ABORTMsg(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.

    mvarABORTMsg = vData
End Property
Public Property Let RETRYMsg(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.

    mvarRETRYMsg = vData
End Property
Public Property Let IGNOREMsg(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.

    mvarIGNOREMsg = vData
End Property
Public Property Let YESMsg(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.

    mvarYESMsg = vData
End Property
Public Property Let NOMsg(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.

    mvarNOMsg = vData
End Property

'Wrapper function for the MessageBox API
Public Function sMsgBox(hWnd As Long, sPrompt As String, _
                       Optional dwStyle As Long, _
                       Optional sTitle As String) As Long
  
    Dim hInstance As Long
    Dim hThreadId As Long
    

    hInstance = App.hInstance
    hThreadId = App.ThreadID
    hFormhWnd = hWnd
    
    'set hook
    hHook = SetWindowsHookEx(WH_CBT, _
                            AddressOf CBTProc, _
                            hInstance, hThreadId)
    'call the MessageBox API and return the
    'value as the result of the function.
    sMsgBox = MessageBox(hWnd, sPrompt, sTitle, dwStyle)

End Function


Public Function CBTProc(ByVal nCode As Long, _
                               ByVal wParam As Long, _
                               ByVal lParam As Long) As Long
      
 
    Dim rc As RECT
    Dim rcFrm As RECT
    Dim newLeft As Long
    Dim newTop As Long
    Dim dlgWidth As Long
    Dim dlgHeight As Long
    Dim scrWidth As Long
    Dim scrHeight As Long
    Dim frmLeft As Long
    Dim frmTop As Long
    Dim frmWidth As Long
    Dim frmHeight As Long
    Dim hwndMsgBox As Long
    
  'When the message box is about to be shown,
  'we'll change the titlebar text, prompt message
  'and button captions
    If nCode = HCBT_ACTIVATE Then

        hwndMsgBox = wParam
          'in a HCBT_ACTIVATE message, wParam holds
          'the handle to the messagebox
          'SetWindowText wParam, "VBnet MessageBox Hook Demo"
        
        'Call GetWindowRect(hwndMsgBox, rc)
        'Call GetWindowRect(hFormhWnd, rcFrm)
     
        'frmLeft = rcFrm.Left
        'frmTop = rcFrm.Top
        'frmWidth = rcFrm.Right - rcFrm.Left
        'frmHeight = rcFrm.Bottom - rcFrm.Top

        'dlgWidth = rc.Right - rc.Left
        'dlgHeight = rc.Bottom - rc.Top
      
        'scrWidth = Screen.Width \ Screen.TwipsPerPixelX
        'scrHeight = Screen.Height \ Screen.TwipsPerPixelY
      
        'newLeft = frmLeft + ((frmWidth - dlgWidth) \ 2)
        'newTop = frmTop + ((frmHeight - dlgHeight) \ 2)
        
     'the ID's of the buttons on the message box
     'correspond exactly to the values they return,
     'so the same values can be used to identify
     'specific buttons in a SetDlgItemText call.
        Call SetDlgItemText(hwndMsgBox, IDOK, mvarOKMsg)
        Call SetDlgItemText(hwndMsgBox, IDCANCEL, mvarCANCELMsg)
        Call SetDlgItemText(hwndMsgBox, IDABORT, mvarABORTMsg)
        Call SetDlgItemText(hwndMsgBox, IDRETRY, mvarRETRYMsg)
        Call SetDlgItemText(hwndMsgBox, IDIGNORE, mvarIGNOREMsg)
        Call SetDlgItemText(hwndMsgBox, IDYES, mvarYESMsg)
        Call SetDlgItemText(hwndMsgBox, IDNO, mvarNOMsg)
        
        'Change the dialog prompt text ...
        'SetDlgItemText wParam, IDPROMPT, "MyApp will now locate the application." & _
        '                                  "Please select the drive to search."

     

        'Call MoveWindow(hwndMsgBox, newLeft, newTop, dlgWidth, dlgHeight, True)
      
        'release the hook
        UnhookWindowsHookEx hHook
    End If
    CBTProc = False

End Function

