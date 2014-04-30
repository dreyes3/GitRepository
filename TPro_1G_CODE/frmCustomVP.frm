VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient1.ocx"
Begin VB.Form frmCustomVP 
   BorderStyle     =   0  'None
   Caption         =   "CWT Desktop - PNR"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser wbPNR 
      Height          =   2040
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   10935
      ExtentX         =   19288
      ExtentY         =   3598
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MyCommandButton.MyButton cmdEnd 
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Tag             =   "ER"
      Top             =   150
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BackColor       =   15323324
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   15323324
      BackColorDisabled=   15323324
      BorderColor     =   14990496
      TransparentColor=   14215660
      Caption         =   "ER"
      Depth           =   1
      DepthEvent      =   1
   End
   Begin MyCommandButton.MyButton cmdEnd 
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Tag             =   "E"
      Top             =   150
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BackColor       =   15323324
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   15323324
      BackColorDisabled=   15323324
      BorderColor     =   14990496
      TransparentColor=   14215660
      Caption         =   "E"
      Depth           =   1
      DepthEvent      =   1
   End
   Begin MyCommandButton.MyButton cmdIgnore 
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Tag             =   "IR"
      Top             =   150
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BackColor       =   15323324
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   15323324
      BackColorDisabled=   15323324
      BorderColor     =   14990496
      TransparentColor=   14215660
      Caption         =   "IR"
      Depth           =   1
      DepthEvent      =   1
   End
   Begin MyCommandButton.MyButton cmdIgnore 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Tag             =   "I"
      Top             =   150
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BackColor       =   15323324
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   15323324
      BackColorDisabled=   15323324
      BorderColor     =   14990496
      TransparentColor=   14215660
      Caption         =   "I"
      Depth           =   1
      DepthEvent      =   1
   End
   Begin ARGradientControl.ARGradient ARGradient1 
      Height          =   2550
      Left            =   -120
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4498
      Color           =   16447215
      FinColor        =   16048334
      Caption         =   "ARGradient1"
      ForeColor       =   -2147483630
      GradientSteps   =   65
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MyCommandButton.MyButton cmdContract 
         Height          =   135
         Left            =   5520
         TabIndex        =   8
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   238
         BackColor       =   15523541
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmCustomVP.frx":0000
         BackColorDown   =   15523541
         BackColorOver   =   15523541
         BackColorFocus  =   15523541
         BackColorDisabled=   15523541
         BorderColor     =   8540205
         TransparentColor=   14215660
         Caption         =   ""
         DepthMode       =   2
         DepthEvent      =   1
         PictureDisabled =   "frmCustomVP.frx":0082
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdSignon 
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   150
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BackColor       =   11658967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         AppearanceThemes=   2
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   11658967
         BackColorDisabled=   11658967
         BorderColor     =   5342839
         TransparentColor=   14215660
         Caption         =   "Sign On"
         Depth           =   1
         DepthEvent      =   1
      End
      Begin MyCommandButton.MyButton cmdSignOff 
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   150
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BackColor       =   11658967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         AppearanceThemes=   2
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   11658967
         BackColorDisabled=   11658967
         BorderColor     =   5342839
         TransparentColor=   14215660
         Caption         =   "Sign Off"
         Depth           =   1
         DepthEvent      =   1
      End
      Begin MyCommandButton.MyButton cmdEmulate 
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   150
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BackColor       =   11658967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         AppearanceThemes=   2
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   11658967
         BackColorDisabled=   11658967
         BorderColor     =   5342839
         TransparentColor=   14215660
         Caption         =   "Emulate"
         Depth           =   1
         DepthEvent      =   1
      End
      Begin MyCommandButton.MyButton cmdExpand 
         Height          =   135
         Left            =   5520
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   238
         BackColor       =   15523541
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmCustomVP.frx":0134
         BackColorDown   =   15523541
         BackColorOver   =   15523541
         BackColorFocus  =   15523541
         BackColorDisabled=   15523541
         BorderColor     =   8540205
         TransparentColor=   14215660
         Caption         =   ""
         DepthMode       =   2
         DepthEvent      =   1
         PictureDisabled =   "frmCustomVP.frx":01B6
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdFP 
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   150
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   11658967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         AppearanceThemes=   2
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   11658967
         BackColorDisabled=   11658967
         BorderColor     =   5342839
         TransparentColor=   14215660
         Caption         =   "FocalPoint"
         Depth           =   1
         DepthEvent      =   1
      End
      Begin MyCommandButton.MyButton cmdPC 
         Height          =   255
         Left            =   5040
         TabIndex        =   11
         Top             =   150
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   11658967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         AppearanceThemes=   2
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   11658967
         BackColorDisabled=   11658967
         BorderColor     =   5342839
         TransparentColor=   14215660
         Caption         =   "P and C"
         Depth           =   1
         DepthEvent      =   1
      End
   End
End
Attribute VB_Name = "frmCustomVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents objVPlistener As VIEWPOINTLISTENERLib.ListenerObj
Attribute objVPlistener.VB_VarHelpID = -1

Private Sub cmdEmulate_Click()

Dim popwindow
Set popwindow = CreateObject("GIUtils.MenuExec")
popwindow.PostMenuMessage ("ID_EMULATE")

End Sub

Private Sub cmdEnd_Click(Index As Integer)

Dim strResponse As String
Dim strReceive As String
cmdEnd(Index).Enabled = False
If frmSideBar.txtRequestor.Text = "" Then
    strReceive = gobjHost.AgentSine
Else
    strReceive = frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine
End If


If cmdEnd(Index).Tag = "E" Then
    strResponse = gobjHost.ENDPNR(strReceive)
Else
    strResponse = gobjHost.ENDPNR(strReceive, True)
End If

If strResponse <> "True" Then
     modMsgBox.OKMsg = "OK"
     modMsgBox.sMsgBox gVPMDIHwnd, "Unable to End PNR. Response from GDS: " & vbCrLf & strResponse, vbOKOnly + vbDefaultButton1, "CWT Desktop - End PNR"
Else
    Set gobjPNR = New CWT_GalileoPNR3.PNR
    gobjPNR.loadPNR
    displayPNRinBar
End If
cmdEnd(Index).Enabled = True

End Sub
'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
'Change to Public sub
Public Sub cmdExpand_Click()
   
   Me.Move 0, gFPHeight - frmCustomVP.Height
   cmdExpand.Visible = False
   cmdContract.Visible = True
    
   'JY - 20100604 - Reset back the height of frmSideBar
   frmSideBar.Height = gFPHeight - frmCustomVP.Height
   frmSideBar.ARGradient1.Height = gFPHeight - frmCustomVP.Height
   frmSideBar.fraInfo.Height = gFPHeight - frmCustomVP.Height - frmSideBar.fraInfo.Top - gPadding
   frmSideBar.treeViewTraveller.Height = frmSideBar.fraInfo.Height - (frmSideBar.treeViewTraveller.Top - frmSideBar.fraInfo.Top) - gPadding
   frmSideBar.Move 0, 0
       resizePCWindow frmSideBar.cmdReverse.Visible, frmCustomVP.cmdContract.Visible
   
End Sub
'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
'Change to Public sub
Public Sub cmdContract_Click()

    Me.Move 0, gFPHeight - gPadding
    cmdExpand.Visible = True
    cmdContract.Visible = False
    
    'JY - 20100604 - Reset back the height of frmSideBar
    frmSideBar.Height = gFPHeight
    frmSideBar.ARGradient1.Height = gFPHeight
    frmSideBar.fraInfo.Height = gFPHeight - frmSideBar.fraInfo.Top - gPadding
    frmSideBar.treeViewTraveller.Height = frmSideBar.fraInfo.Height - (frmSideBar.treeViewTraveller.Top - frmSideBar.fraInfo.Top) - gPadding
    frmSideBar.Move 0, 0
    resizePCWindow frmSideBar.cmdReverse.Visible, frmCustomVP.cmdContract.Visible
            
End Sub

Private Sub cmdIgnore_Click(Index As Integer)

cmdIgnore(Index).Enabled = False
If cmdIgnore(Index).Tag = "I" Then
    gobjHost.IgnorePNR
Else
    gobjHost.IgnorePNR True, gobjPNR.RecLoc
End If

Set gobjPNR = New CWT_GalileoPNR3.PNR
gobjPNR.loadPNR
displayPNRinBar
cmdIgnore(Index).Enabled = True

End Sub

Private Sub cmdSignOff_Click()

Dim strResponse As String
Dim strTemp As String

strResponse = gobjHost.SignOff

If strResponse <> "True" Then
     modMsgBox.OKMsg = "OK"
     modMsgBox.sMsgBox gVPMDIHwnd, "Unable to Sign Off, response from GDS:" & vbCrLf & strResponse, vbOKOnly + vbDefaultButton1, "CWT Desktop - SignOff"
     populateDefault
Else
     disableControls
     'JY - 20100608 - Reset back the values after signed off
     strTemp = gobjHost.WorkAreas
End If

End Sub
Private Sub cmdSignon_Click()

    populateDefault
    enableControls
    Dim popwindow
    Set popwindow = CreateObject("GIUtils.MenuExec")
    popwindow.PostMenuMessage ("ID_SIGN_ON")
    
End Sub

Private Sub Form_Load()

Dim oldParent As Long

Set objVPlistener = New VIEWPOINTLISTENERLib.ListenerObj
oldParent = SetParent(Me.hwnd, gVPMDIHwnd)
Me.Width = gFPWidth
Me.Height = gCustomVPHeight
ARGradient1.Width = gFPWidth
ARGradient1.Height = frmCustomVP.Height

wbPNR.Width = ARGradient1.Width - wbPNR.Left - gPadding
wbPNR.Height = frmCustomVP.Height - wbPNR.Top - gPadding

cmdExpand.Left = ARGradient1.Width / 2
cmdContract.Left = ARGradient1.Width / 2

Me.Move 0, 0
Me.Move 0, gFPHeight - frmCustomVP.Height
RemoveMenus Me, False, False, False, False, False, True, True
       
   ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    If gIntModuleType <> gModuleType.PC Then
         'Do not need to perform enableControls for Smart Point
    Else
        enableControls
    End If

If gobjHost.AgentSine <> "" Then populateDefault

wbPNR.Navigate ("C:\fp\swdir\CustomViewpoint\CustomViewpoint.html")

End Sub
Private Sub populateDefault()

If Not IsNumeric(gobjHost.AgentSine) Then
    cmdEmulate.Visible = False
Else
    cmdEmulate.Visible = True
End If

End Sub

Private Sub cmdFP_Click()
    Dim oDU As Galileo.DesktopUtils
    Set oDU = New Galileo.DesktopUtils
    
    Dim pm As Viewpoint.PartitionManager
    Set pm = New Viewpoint.PartitionManager

    Dim childWin As String
    childWin = pm.partitions(0).title
    
    Dim pcHdl As Long
    pcHdl = oDU.GetGalileoDesktopChildWindowHandle("Point-and-Click")
    If (pcHdl <> 0) Then
        'Minimize Point & Click Screen
        oDU.MinimizeWindow (pcHdl)
        'Set Focus on Focal Point window
        oDU.SetHighZOrder (oDU.GetGalileoDesktopChildWindowHandle(childWin))
    End If
    
End Sub

Private Sub cmdPC_Click()
    
    'Send command to display Point & Click Screen
    pDisplayToFP ":*R"
    resizePCWindow frmSideBar.cmdReverse.Visible, frmCustomVP.cmdContract.Visible
    
End Sub

Private Sub objVPlistener_OnViewpointEvent(ByVal EventStr As String)
   
   Dim mxmldomEvent As MSXML2.DOMDocument
   Dim strEventName As String
   Set mxmldomEvent = New MSXML2.DOMDocument
   Dim strTemp As String
   
   mxmldomEvent.async = False
   If mxmldomEvent.loadXML(EventStr) = False Then Exit Sub
   If IsNull(mxmldomEvent.documentElement) = False And IsNull(mxmldomEvent.firstChild) = False Then
        strEventName = mxmldomEvent.documentElement.firstChild.nodeName
   End If

   If strEventName = "SwitchedToFocalpoint" Then
        Call EnumChildWindows(glngTargetHwnd, AddressOf EnumChildWindowProc, 0&)
   End If
        
   If strEventName = "SignOn" Then
        populateDefault
        enableControls
        'JY - 20100608 - Reset back the values after signed off
        strTemp = gobjHost.WorkAreas
    End If
    
    If strEventName = "PseudoCityCodeModified" Then
        populateDefault
        'JY - 20100608 - Reset back the values after signed off
        strTemp = gobjHost.WorkAreas
    End If
    
End Sub

Private Sub enableControls()

Dim intI As Integer

If gobjHost.AgentSine = "" Then
    cmdSignOff.Enabled = False
    cmdSignon.Enabled = True
    cmdEmulate.Enabled = False
    frmBars.SftTabs.Enabled = False
    frmSideBar.fraSearch.Enabled = False
    For intI = 0 To cmdIgnore.Count - 1
        cmdIgnore.item(intI).Enabled = False
    Next
    For intI = 0 To cmdEnd.Count - 1
        cmdEnd.item(intI).Enabled = False
    Next
Else
    cmdSignOff.Enabled = True
    cmdSignon.Enabled = False
    cmdEmulate.Enabled = True
    frmBars.SftTabs.Enabled = True
    frmSideBar.fraSearch.Enabled = True
    For intI = 0 To cmdIgnore.Count - 1
        cmdIgnore.item(intI).Enabled = True
    Next
    For intI = 0 To cmdEnd.Count - 1
        cmdEnd.item(intI).Enabled = True
    Next
End If

End Sub
Private Sub disableControls()

For intI = 0 To cmdIgnore.Count - 1
    cmdIgnore.item(intI).Enabled = False
Next
For intI = 0 To cmdEnd.Count - 1
    cmdEnd.item(intI).Enabled = False
Next

cmdSignOff.Enabled = False
cmdSignon.Enabled = True
cmdEmulate.Enabled = False
frmBars.SftTabs.Enabled = False
frmSideBar.fraSearch.Enabled = False

End Sub

