VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeatMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SeatMapSeat"
Attribute VB_Ext_KEY = "Member1" ,"SeatRowAttribute"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarDisplayType As String 'local copy
Private mvarColLabel As String 'local copy
Private mvarSeatRow As SeatMapRow
Private mcolSeatRow As Collection
Private mvarFlight As String 'local copy
Private mvarStartCity As String 'local copy
Private mvarEndCity As String 'local copy
Private mvarAClass As String 'local copy
Private mvarVendor As String 'local copy
Private mvarStartDate As String 'local copy



Public Property Get SeatMapSeat() As SeatMapRow
    If mvarSeatMapRow Is Nothing Then
        Set mvarSeatMapRow = New SeatMapRow
    End If


    Set SeatMapRow = mvarSeatMapRow
End Property


Public Property Set SeatMapRow(vData As SeatMapRow)
    Set mvarSeatMapRow = vData
End Property




Private Sub Class_Initialize()
Set mcolSeatRow = New Collection
End Sub

Private Sub Class_Terminate()
  Set mvarSeatMapRow = Nothing
  
End Sub



Public Property Let ColLabel(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ColLabel = 5
    mvarColLabel = vData
End Property


Public Property Get ColLabel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ColLabel
    ColLabel = mvarColLabel
End Property



Public Property Let DisplayType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DisplayType = 5
    mvarDisplayType = vData
End Property


Public Property Get DisplayType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DisplayType
    DisplayType = mvarDisplayType
End Property
Public Sub SeatMapRowAdd(ByVal SeatRow As SeatMapRow)
    mcolSeatRow.Add SeatRow
    Set SeatRow = Nothing
End Sub

Public Function SeatRow(ByVal Index As Long) As SeatMapRow
    Set SeatRow = mcolSeatRow.Item(Index)
End Function

Public Property Get SeatCountRow() As Long
    SeatCountRow = mcolSeatRow.Count
End Property

Public Property Let Vendor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RowNumber = 5
    mvarVendor = vData
End Property


Public Property Get Vendor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RowNumber
    Vendor = mvarVendor
End Property
Public Property Let Flight(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RowNumber = 5
    mvarFlight = vData
End Property


Public Property Get Flight() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RowNumber
    Flight = mvarFlight
End Property
Public Property Let AClass(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RowNumber = 5
    mvarAClass = vData
End Property


Public Property Get AClass() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RowNumber
    AClass = mvarAClass
End Property

Public Property Let StartDate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RowNumber = 5
    mvarStartDate = vData
End Property


Public Property Get StartDate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RowNumber
    StartDate = mvarStartDate
End Property
Public Property Let StartCity(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RowNumber = 5
    mvarStartCity = vData
End Property


Public Property Get StartCity() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RowNumber
    StartCity = mvarStartCity
End Property

Public Property Let EndCity(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RowNumber = 5
    mvarEndCity = vData
End Property


Public Property Get EndCity() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RowNumber
    EndCity = mvarEndCity
End Property

