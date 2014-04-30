VERSION 5.00
Begin VB.Form frmFareRmkFill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CWT TravelPro - Fare Remarks Fill In"
   ClientHeight    =   1290
   ClientLeft      =   3420
   ClientTop       =   2565
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6060
      TabIndex        =   4
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   315
      Left            =   6060
      TabIndex        =   3
      Top             =   540
      Width           =   1095
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   0
      Left            =   3060
      TabIndex        =   2
      Top             =   660
      Width           =   2715
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "Field"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   660
      Width           =   2475
   End
   Begin VB.Label lblRmkText 
      Caption         =   "Remark"
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   6555
   End
End
Attribute VB_Name = "frmFareRmkFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strRmk As String
Dim lngCount As Long

Private Sub cmdCancel_Click()
    lblRmkText.Caption = ""
    Me.Hide

End Sub

Private Sub cmdDone_Click()
Dim strMsg As String

For lngCount = 0 To txtField.Count - 1
    If txtField(lngCount).Text = "" Then
        strMsg = "Nothing entered for " & lblField(lngCount).Caption
        'MsgBox strMsg
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        Exit Sub
    End If
    
Next

For lngCount = 0 To lngCount - 1
    strRmk = Left(strRmk, InStr(strRmk, "%" & lngCount) - 1) _
        & txtField(lngCount).Text & Mid(strRmk, InStr(strRmk, "%" & lngCount) + Len("%" & lngCount))
Next
lblRmkText.Caption = strRmk
Me.Hide

End Sub

Public Sub FormatRemark()
Dim lngBeg As Long
Dim lngEnd As Long
Dim lngLen As Long
Dim strPrompt As String

strRmk = lblRmkText.Caption


Do While InStr(1, strRmk, "[")
    lngEnd = 1
    lngBeg = InStr(lngEnd, strRmk, "[") - 1
    lngEnd = InStr(lngEnd, strRmk, "]") + 1
    lngLen = (lngEnd - lngBeg)
    strPrompt = Mid(strRmk, lngBeg + 2, lngLen - 3)
    strRmk = Left(strRmk, lngBeg) & "%" & lngCount & Mid(strRmk, lngEnd)
    Debug.Print strRmk
    If lngCount > 0 Then
        Load lblField(lngCount)
        Load txtField(lngCount)
        Me.Height = Me.Height + 300
    End If
    With txtField(lngCount)
        .Text = ""
        .Tag = lngBeg
        .Top = 660 + (300 * lngCount)
        .Visible = True
    End With
    With lblField(lngCount)
        .Caption = strPrompt
        .Top = 660 + (300 * lngCount)
        .Visible = True
    End With
    lngCount = lngCount + 1
Loop

For lngCount = 0 To txtField.Count - 1
    txtField(lngCount).TabIndex = lngCount
Next

End Sub

Private Sub Form_Load()
  Dim oldParent As Long
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
  Me.Move 0, 0
  Me.Move frmSideBar.Width, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFareRmkFill = Nothing
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
Dim lngBeg As Long
Dim lngLen As Long

        Select Case KeyAscii
            Case 8, 32, 46, 48 To 57, 65 To 90, 97 To 122
               KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change valid characters to uppercase
            Case Else
               KeyAscii = 0
        End Select
  
End Sub
