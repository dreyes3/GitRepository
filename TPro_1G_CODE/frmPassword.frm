VERSION 5.00
Begin VB.Form frmPassword 
   Caption         =   "CWT TravelPro"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   1125
   ScaleWidth      =   4530
   Begin VB.TextBox txtPswd 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim oldParent As Long
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
   txtPswd.PasswordChar = "*"
   
End Sub
Private Function pValidatePassword() As Boolean
   Dim rs As New ADODB.Recordset
   Dim strPwd As String
   Dim strPassword As String
   
  
   Set rs = gdbConn.Execute("Select * from tblPassword")
   strPwd = rs!RaiseChequePwd
   rs.Close
   Set rs = Nothing
   
   If Trim(txtPswd) = strPwd Then
      pValidatePassword = True
   Else
      pValidatePassword = False
   End If
   
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmPassword = Nothing
    Unload Me
End Sub

Private Sub txtPswd_KeyPress(KeyAscii As Integer)
    Dim strMsg As String
    
    If KeyAscii = 13 Then
        If pValidatePassword = True Then
            Unload Me
            frmRaiseCheque.Show
        Else
            'MsgBox "Invalid Password"
            strMsg = "Invalid Password"
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
            txtPswd = ""
            txtPswd.SetFocus
        End If
    End If
End Sub
