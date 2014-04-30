VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeatData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarSegNum As Integer 'local copy
Private mvarPaxNo As Integer 'local copy
Private mvarStatus As String 'local copy
Private mvarSeatLocation As String 'local copy
'Private mvarSeatAttribute As Collection 'local copy
Private mvarSeatAttribute1 As String
Private mvarSeatAttribute2 As String


'Public Function SeatAttribute(ByVal Index As Variant) As String
'    Set SeatAttribute = mvarSeatAttribute.Item(Index)
'End Function

'Public Sub AddSeatAttribute(ByVal SeatA As String)
'        mvarSeatAttribute.Add SeatA

'End Sub



'Public Property Get SeatAttributeCount() As Long
'    SeatAttributeCount = mvarSeatAttribute.Count
'End Property
Public Property Let SeatAttribute1(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SeatLocation = 5
    mvarSeatAttribute1 = vData
End Property


Public Property Get SeatAttribute1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SeatLocation
    SeatAttribute1 = mvarSeatAttribute1
End Property
Public Property Let SeatAttribute2(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SeatLocation = 5
    mvarSeatAttribute2 = vData
End Property


Public Property Get SeatAttribute2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SeatLocation
    SeatAttribute2 = mvarSeatAttribute2
End Property
Public Property Let SeatLocation(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SeatLocation = 5
    mvarSeatLocation = vData
End Property


Public Property Get SeatLocation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SeatLocation
    SeatLocation = mvarSeatLocation
End Property



Public Property Let Status(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Status = 5
    mvarStatus = vData
End Property


Public Property Get Status() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
End Property



Public Property Let PaxNo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PaxNo = 5
    mvarPaxNo = vData
End Property


Public Property Get PaxNo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PaxNo
    PaxNo = mvarPaxNo
End Property



Public Property Let SegNum(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SegNum = 5
    mvarSegNum = vData
End Property


Public Property Get SegNum() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SegNum
    SegNum = mvarSegNum
End Property



Private Sub Class_Initialize()
Set mvarSeatAttribute = New Collection
End Sub

Private Sub Class_Terminate()
Set mvarSeatAttribute = Nothing
End Sub
