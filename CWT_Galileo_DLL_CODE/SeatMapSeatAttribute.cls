VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeatMapSeatAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarSeatAttribute As String 'local copy
Public Property Let SeatAttribute(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SeatAttribute = 5
    mvarSeatAttribute = vData
End Property


Public Property Get SeatAttribute() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SeatAttribute
    SeatAttribute = mvarSeatAttribute
End Property



