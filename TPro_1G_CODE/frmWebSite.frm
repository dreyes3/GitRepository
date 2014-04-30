VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Begin VB.Form frmWebSite 
   Caption         =   "WebSite"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
      ExtentX         =   1296
      ExtentY         =   1296
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
      Location        =   "http:///"
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButton     =   0
      MaxButton       =   0
      MinButton       =   0
      OldForeColor    =   0
      ChangeSkinButton=   0   'False
      SysDisableSkinCaption=   "&Disable Skin"
      LcK1            =   "..02*-0..*/305*.-2-/"
      LcK2            =   $"frmWebSite.frx":0000
      AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
   End
End
Attribute VB_Name = "frmWebSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
'This form will be served as a browser to open a website

Private gbolHBTWebsite As Boolean
Private classHBTListener As HBTListener

Private Sub Form_Load()

     Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")

    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    If gIntModuleType = gModuleType.SYEX Then
    
        MakeWinTopMost (Me.hwnd)
        Me.Top = 131 * Screen.TwipsPerPixelY
        Me.Width = frmCustomVP.Width - frmSideBar.Width - gPadding
        'Me.Height = frmSideBar.Height - gPadding
        ' ZhiSam - V1.2.18 20120311 - CR-203 - Desktop to Create Retention Line and Update TAW to TAU (SyEx with Tpro)
        ' Reset the HBU screen's height to same as focalpoint window size
        Me.Height = gFPHeight
        
    Else
        If gIntModuleType = gModuleType.PC Then
            oldParent = SetParent(Me.hwnd, gVPMDIHwnd)
        End If
           
        'Close down the sideBar & Custom ViewPoint
        frmSideBar.cmdReverse_Click
        frmCustomVP.cmdContract_Click
   
        Me.Width = frmCustomVP.Width - frmSideBar.Width - gPadding
        Me.Height = frmSideBar.Height - gPadding
        Me.Move frmSideBar.Width, 0

    End If

   
   WebBrowser1.Width = Me.Width
   WebBrowser1.Height = Me.Height - 100

   
              
End Sub

Public Function openWebsite(strType As String)
   If strType = "HBT" Then
      gbolHBTWebsite = True
      'Open HBT Website
      openHBTSite
   End If
End Function

Private Function openHBTSite()
   WebBrowser1.Navigate (gstrHBTURL)
End Function

Private Sub Form_Unload(Cancel As Integer)

    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' do not trigger frmSideBar and frmCustomVP for SyEx module
    If gIntModuleType <> gModuleType.SYEX Then
        'Open back the sideBar and Custom ViewPoint
        frmSideBar.cmdExpand_Click
        frmCustomVP.cmdExpand_Click
    End If
   

   
   If gbolHBTWebsite = True Then
      gbolExitHBT = True
      If gbolCreatPNR = True Then
         'Switch back to Air tab
         frmBars.SftTabs.Tabs.Current = 0
      End If
      If gobjPNR.CheckPNRStatus = 3 Then gobjHost.terminalEntry "IR"
   End If
   
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
   Dim strJS As String
          
   If gbolHBTWebsite = True Then
        If URL = gstrHBTURL Then
        
           'Set the HBT Listener
           Set classHBTListener = New HBTListener
    
           'populate values into the hidden textbox
           strJS = "document.getElementById(""PNRLocator"").value = '" & gobjPNR.RecLoc & "';"
           strJS = strJS & "document.getElementById(""ProfileNames"").value = '" & gobjPNR.CompInfo.ProfileName & "';"
           strJS = strJS & "document.getElementById(""ProfilePCC"").value = '" & gobjPNR.CompInfo.ProfilePCC & "';"
           strJS = strJS & "document.getElementById(""AgentPCC"").value = '" & gobjHost.AgentPCC & "';"
           strJS = strJS & "document.getElementById(""AgentSignIn"").value = '" & gobjHost.AgentSine & "';"
           strJS = strJS & "init();"
           
           'Set the g_VBOBject in the HBT Website to the HBTListener created
           Set WebBrowser1.Document.Script.g_VBObject = classHBTListener
           WebBrowser1.Document.parentWindow.execScript strJS, "JScript"
                                                                                                                                                                                                       
        End If
   End If
End Sub

