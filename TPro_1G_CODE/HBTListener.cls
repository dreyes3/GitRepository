VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HBTListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub HotelBookWebsiteEvent(eventId As Integer, paramXML As String)
    'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
    'This function will be called when HBT returns an event to CWT Desktop
    If eventId = 0 Then
       'Exit button is clicked in HBT
       Unload frmWebSite
    End If
End Sub
