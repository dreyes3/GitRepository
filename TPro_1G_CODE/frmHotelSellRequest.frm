VERSION 5.00
Begin VB.Form frmHotelSellRequest 
   Caption         =   "CWT TravelPro - Hotel Rental"
   ClientHeight    =   2655
   ClientLeft      =   5610
   ClientTop       =   7905
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   5070
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Co&ntinue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame fraAddHotel 
      Caption         =   "Select a Hotel Sell Method: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4815
      Begin VB.ComboBox cmbAirSeg 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
      Begin VB.ComboBox cmbSellMtd 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblSegment 
         Caption         =   "After Segment No:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Hotel Rental"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmHotelSellRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objVP As Viewpoint.ViewpointSrv
Dim objVP2 As Viewpoint.ViewpointSrv2
Dim WithEvents objVPlistener As VIEWPOINTLISTENERLib.ListenerObj
Attribute objVPlistener.VB_VarHelpID = -1
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date
Dim datViewPtStart As Date
Dim strSellType As String

Private Sub cmbSellMtd_click()
    If cmbSellMtd.listindex = 2 Then
        lblSegment.Visible = False
        cmbAirSeg.Visible = False
    Else
        lblSegment.Visible = True
        cmbAirSeg.Visible = True
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdContinue_Click()
Dim intTemp As Integer
Dim intresponse As Integer

datTouchEnd = Now

  pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
  gconModHtl, frmSideBar.cmbSelectType.Text, gconSModHtlSell, _
  Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
  pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
  gconModHtl, frmSideBar.cmbSelectType.Text, gconSModHtlSell, _
  Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
  pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
  gconModHtl, frmSideBar.cmbSelectType.Text, gconSModHtlSell, _
  Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
 
objVP.RetrieveCurrentPNR
If cmbAirSeg.listindex > -1 Then
    objVP.SetCurrentSegment (Left(cmbAirSeg.Text, InStr(cmbAirSeg.Text, ".") - 1))
End If

Select Case cmbSellMtd.listindex
Case 0:
    'Me.Hide
    datViewPtStart = Now
    objVP.HotelAvail
    strSellType = "VP - HOTEL REFERENCE SELL"
    'Me.Show
Case 1:
    'Me.Hide
    datViewPtStart = Now
    objVP.HotelSell
    strSellType = "VP - HOTEL DIRECT SELL"
    'Me.Show
Case 2:
    Unload Me
    'Load frmCarPassiveSell
    'frmCarPassiveSell.Show
    'Call ViewPoint Hotel Passive
    ShellExecute 0, "open", "C:/fp/swdir/CustomViewpoint/Scripts/passive.html", vbNullString, vbNullString, 1
    'If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModHtl, frmSideBar.cmbSelectType.Text, gconSModHtlSell, _
    Me.Name, gconFormLoad, gstrProcessGrpID, , datTouchEnd

End Select

End Sub

Private Sub Form_Load()
Dim lngC As Long
Dim oldParent As Long

datFormLoadStart = Now
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
Me.Move 0, 0
Me.Move frmSideBar.Width, 0
'gStartCarTime = Now

Set objVP = New Viewpoint.ViewpointSrv
Set objVPlistener = New VIEWPOINTLISTENERLib.ListenerObj

cmbSellMtd.AddItem "Book a Hotel through Reference Sell in Galileo"
cmbSellMtd.AddItem "Book a Hotel through Direct Sell in Galileo"
cmbSellMtd.AddItem "Add a Hotel Passive Segment to the PNR"

cmbSellMtd.listindex = 0

'300707:  Tpro Hotel rental - To display AIR segments
Set gobjPNR = New CWT_GalileoPNR3.PNR
gobjPNR.loadPNR
If gobjPNR.AirSegCount > 0 Then
    For lngC = 1 To gobjPNR.AirSegCount
        cmbAirSeg.AddItem gobjPNR.AirSeg(lngC).TextAirSeg
    Next
    cmbAirSeg.listindex = 0
End If

datFormLoadEnd = Now
If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmHotelSellRequest = Nothing
End Sub

Private Sub objVPlistener_OnViewpointEvent(ByVal EventStr As String)
Dim mxmldomEvent As MSXML2.DOMDocument
Dim strEventName As String
Set mxmldomEvent = New MSXML2.DOMDocument

Debug.Print EventStr
mxmldomEvent.async = False
If mxmldomEvent.loadXML(EventStr) = False Then Exit Sub
If IsNull(mxmldomEvent.documentElement) = False And IsNull(mxmldomEvent.firstChild) = False Then
    strEventName = mxmldomEvent.documentElement.firstChild.nodeName
End If
    
    If strEventName = "HotelSell" Then
        gobjHost.terminalEntry "R.TPRO HOTELSELL+ER"
        gobjHost.terminalEntry "ER"
        gobjHost.terminalEntry "ER"
          
        pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
        gconModHtl, frmSideBar.cmbSelectType.Text, gconSModHtlSell, _
        Me.Name, strSellType, gstrProcessGrpID, , datViewPtStart
        
        Unload Me
        pDisplayToFP "*R"
        gstrProductType = "HL"
        Load frmCarRmkMI
        frmCarRmkMI.Show
    End If
    
End Sub



