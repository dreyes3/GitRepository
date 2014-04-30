VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeatMaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mcolSeatMap As Collection
Public MapExist As Boolean  'local copy



Public Sub getSeatMap(Vendor As String, FltNum As String, Class As String, StartDate As String, From As String, sTo As String)
Dim xmldomSeatMap As MSXML2.DOMDocument
Dim xmlnlTemp1 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp2 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp3 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp4 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp5 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp6 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp8 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp9 As MSXML2.IXMLDOMNodeList
Dim lngC1 As Long
Dim lngC2 As Long
Dim lngC3 As Long
Dim lngC4 As Long
Dim lngC5 As Long
Dim lngC6 As Long
Dim lngC7 As Long
Dim lngC8 As Long
Dim lngC9 As Long
Dim objSMap As SeatMap
Dim objRow As SeatMapRow
Dim objRowAtt As SeatRowAttribute
Dim objSeat As SeatMapSeat
Dim objSeatAtt As SeatMapSeatAttribute
Dim strTempDisplayType As String
Dim strRequest As String

strRequest = strRequest & "<SeatMap_5_0>"
    strRequest = strRequest & "<SeatMapMods>"
        strRequest = strRequest & "<AirV>" & Vendor & "</AirV>"
        strRequest = strRequest & "<FltNum>" & FltNum & "</FltNum>"
        strRequest = strRequest & "<BIC>" & Class & "</BIC>"
        strRequest = strRequest & "<StartDt>" & StartDate & "</StartDt>"
        strRequest = strRequest & "<StartCity>" & From & "</StartCity>"
        strRequest = strRequest & "<EndCity>" & sTo & "</EndCity>"
        strRequest = strRequest & "<NumSeats>1</NumSeats>"
    strRequest = strRequest & "</SeatMapMods>"
strRequest = strRequest & "</SeatMap_5_0>"


'strRequest = strRequest & "<SeatMap_5_0>"

'strRequest = strRequest & "   <SeatMapMods>"

 'strRequest = strRequest & "     <AirV>UA</AirV>"

 'strRequest = strRequest & "     <FltNum>804</FltNum>"

' strRequest = strRequest & "     <BIC>C</BIC>"

' strRequest = strRequest & "     <StartDt>20080928</StartDt>"

'strRequest = strRequest & "      <StartCity>SIN</StartCity>"

' strRequest = strRequest & "     <EndCity>NRT</EndCity>"

' strRequest = strRequest & "  </SeatMapMods>"

'strRequest = strRequest & "</SeatMap_5_0>"



'strRequest = ""

'strRequest = strRequest & "<PNRBFManagement_11>"
'strRequest = strRequest & " <PNRBFRetrieveMods>"


'    strRequest = strRequest & "  <CurrentPNR />"

'strRequest = strRequest & "  </PNRBFRetrieveMods>"

'strRequest = strRequest & "  </PNRBFManagement_11>"



Set xmldomSeatMap = CreateObject("microsoft.xmldom")
xmldomSeatMap.async = False
'Set xmldomSeatMap = New MSXML2.DOMDocument

      'Set gobjHost = New GalileoHost
    
      xmldomSeatMap.loadXML gobjHost.SendQuery(strRequest, "SeatMap_5_0", "SeatMap", "GetSeatMap")
   
      'xmldomSeatMap.save "C:\AAA.xml"
      
      
      If InStr(xmldomSeatMap.xml, "TransactionErrorCode") <> 0 Then
      If InStr(xmldomSeatMap.xml, "ErrText") > 0 Then
           If Trim(xmldomSeatMap.selectNodes("//ErrText/Text").Item(0).Text) <> "" Then
              'MsgBox xmldomSeatMap.selectNodes("//xmldomSeatMap.xmlErrText/Text").Item(0).Text, vbApplicationModal + vbExclamation + vbOKOnly, "CWT Agent Desktop - Seat"
           
           Else
              'MsgBox "Error loading seat map", vbApplicationModal + vbExclamation + vbOKOnly, "CWT Agent Desktop - Seat"
           End If
           MapExist = False
           Exit Sub
       Else
            'MsgBox xmldomSeatMap.xml, vbApplicationModal + vbExclamation + vbOKOnly, "CWT Agent Desktop - Seat"
            MapExist = False
            Exit Sub
       End If
      End If
      Set xmlnlTemp1 = xmldomSeatMap.selectNodes("//SeatMap/SeatMapQual").Item(0).childNodes
      
      MapExist = True
      
      For lngC1 = 0 To xmlnlTemp1.length - 1
        With xmlnlTemp1.Item(lngC1)
        Select Case .nodeName
        Case "TypeDisp"
               strTempDisplayType = .Text
        Case "SectionsAry"
        
        Set xmlnlTemp8 = .selectNodes("//SectionsAry/Sections")
        For lngC8 = 0 To xmlnlTemp8.length - 1
                
        Set xmlnlTemp9 = xmlnlTemp8(lngC8).childNodes
        
               Set objSMap = New SeatMap
               If strTempDisplayType <> "" Then objSMap.DisplayType = strTempDisplayType
               objSMap.AClass = Class
               objSMap.Flight = FltNum
               objSMap.Vendor = Vendor
               objSMap.StartCity = From
               objSMap.EndCity = sTo
               objSMap.StartDate = StartDate
               
               'Set xmlnlTemp2 = .selectNodes("//SectionsAry/Sections/RowAry/Row")
               'objSMap.ColLabel = .selectSingleNode("//SectionsAry/Sections/ColLabel").Text
For lngC9 = 0 To xmlnlTemp9.length - 1
    Select Case xmlnlTemp9.Item(lngC9).nodeName
    Case "ColLabel"
        objSMap.ColLabel = xmlnlTemp9.Item(lngC9).Text
    
    Case "RowAry"
    
        Set xmlnlTemp2 = xmlnlTemp9.Item(lngC9).childNodes
              
    
    For lngC3 = 0 To xmlnlTemp2.length - 1
        Set objRow = New SeatMapRow
        
            For lngC6 = 0 To xmlnlTemp2.Item(lngC3).childNodes.length - 1
            With xmlnlTemp2.Item(lngC3).childNodes.Item(lngC6)
            Select Case .nodeName
                Case "Num"
                    objRow.RowNumber = .Text
                Case "AttribAry"
                    If .hasChildNodes Then
                        Set xmlnlTemp3 = .childNodes
                        For lngC4 = 0 To xmlnlTemp3.length - 1
                            Set objRowAtt = New SeatRowAttribute
                            Select Case xmlnlTemp3(lngC4).nodeName
                                Case "Attrib"
                                    objRowAtt.RAttribute = xmlnlTemp3(lngC4).Text
                                    objRow.RowAttributeAdd objRowAtt
                            End Select
                        Next
                    
                    End If
                 Case "SeatsAry"
                    If .hasChildNodes Then
                        Set xmlnlTemp4 = .childNodes
                        For lngC5 = 0 To xmlnlTemp4.length - 1
                            Set objSeat = New SeatMapSeat
                            For lngC7 = 0 To xmlnlTemp4.Item(lngC5).childNodes.length - 1
                               Select Case xmlnlTemp4(lngC5).childNodes.Item(lngC7).nodeName
                               
                                   Case "ColID"
                                       objSeat.ColumnID = xmlnlTemp4(lngC5).childNodes.Item(lngC7).Text
                                   Case "Status"
                                       objSeat.Status = xmlnlTemp4(lngC5).childNodes.Item(lngC7).Text
                                   Case "AttribAry"
                                       If xmlnlTemp4(lngC5).childNodes.Item(lngC7).hasChildNodes Then
                                           Set xmlnlTemp5 = xmlnlTemp4(lngC5).childNodes.Item(lngC7).childNodes
                                               For lngC2 = 0 To xmlnlTemp5.length - 1
                                               
                                                   Set objSeatAtt = New SeatMapSeatAttribute
                                                   
                                                   Select Case xmlnlTemp5(lngC2).nodeName
                                                       
                                                       Case "Attrib"
                                                           objSeatAtt.SeatAttribute = xmlnlTemp5(lngC2).Text
                                                           objSeat.SeatAttributeAdd objSeatAtt
                                                   End Select
                                               Next
                                       End If
                    
                               End Select
                            Next lngC7
                            objRow.SeatAdd objSeat
                        Next
                    
                    End If
            End Select
            End With
            Next lngC6
        
        
        objSMap.SeatMapRowAdd objRow
    Next lngC3
    
    
    
    End Select
    Next lngC9
               '---------------------------------------------------------------------------
        mcolSeatMap.Add objSMap
        
        Next lngC8
        
        End Select
        
        
        
        End With
        
        
        
        
    Next lngC1
    
    
    
End Sub
Public Function SeatMap(ByVal Index As Long) As SeatMap
    Set SeatMap = mcolSeatMap.Item(Index)
End Function

Public Property Get SeatMapCount() As Long
    SeatMapCount = mcolSeatMap.Count
End Property

Private Sub Class_Initialize()
    Set mcolSeatMap = New Collection
End Sub
