VERSION 5.00
Begin VB.Form frmCarSellRequest 
   Caption         =   "CWT TravelPro - Car Rental"
   ClientHeight    =   2655
   ClientLeft      =   5475
   ClientTop       =   4425
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   5070
   Begin VB.Frame fraAddCar 
      Caption         =   "Select a Car Sell Method: "
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
      TabIndex        =   3
      Top             =   600
      Width           =   4815
      Begin VB.ComboBox cmbSellMtd 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
      Begin VB.ComboBox cmbAirSeg 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3135
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
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
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
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
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
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Car Rental"
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
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmCarSellRequest"
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
  gconModCar, frmSideBar.cmbSelectType.Text, gconSModBkCar, _
  Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
  pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
  gconModCar, frmSideBar.cmbSelectType.Text, gconSModBkCar, _
  Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
  pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
  gconModCar, frmSideBar.cmbSelectType.Text, gconSModBkCar, _
  Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
 



objVP.RetrieveCurrentPNR
If cmbAirSeg.listindex > -1 Then
    objVP.SetCurrentSegment (Left(cmbAirSeg.Text, InStr(cmbAirSeg.Text, ".") - 1))
End If

Select Case cmbSellMtd.listindex
Case 0:
    'Me.Hide
    datViewPtStart = Now
    objVP.CarAvail
    strSellType = "VP-CAR REFERENCE SELL"
    'Me.Show
Case 1:
    'Me.Hide
    datViewPtStart = Now
    objVP.CarSell
    strSellType = "VP - CAR DIRECT SELL"
    'Me.Show
Case 2:
    Unload Me
    Load frmCarPassiveSell
    frmCarPassiveSell.Show
    
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
gStartCarTime = Now

Set objVP = New Viewpoint.ViewpointSrv
Set objVPlistener = New VIEWPOINTLISTENERLib.ListenerObj

cmbSellMtd.AddItem "Book a car through Reference Sell in Galileo"
cmbSellMtd.AddItem "Book a car through Direct Sell in Galileo"
cmbSellMtd.AddItem "Add a car Passive Segment to the PNR"

cmbSellMtd.listindex = 0

'300707:  Tpro Car rental - To display AIR segments
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
    Set frmCarSellRequest = Nothing
End Sub

Private Sub objVPlistener_OnViewpointEvent(ByVal EventStr As String)
Dim mxmldomEvent As MSXML2.DOMDocument
Dim strEventName As String
Set mxmldomEvent = New MSXML2.DOMDocument

'Debug.Print EventStr
mxmldomEvent.async = False
If mxmldomEvent.loadXML(EventStr) = False Then Exit Sub
If IsNull(mxmldomEvent.documentElement) = False And IsNull(mxmldomEvent.firstChild) = False Then
    strEventName = mxmldomEvent.documentElement.firstChild.nodeName
End If

    If strEventName = "CarSell" Then
        gobjHost.terminalEntry "R.TPRO CARSELL+ER"
        gobjHost.terminalEntry "ER"
        gobjHost.terminalEntry "ER"
          
        pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
        gconModCar, frmSideBar.cmbSelectType.Text, gconSModBkCar, _
        Me.Name, strSellType, gstrProcessGrpID, , datViewPtStart
        
        Unload Me
        pDisplayToFP "*R"
        gstrProductType = "CX"
        Load frmCarRmkMI
        frmCarRmkMI.Show
    End If
    
End Sub

