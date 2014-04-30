VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeatRowAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Private mvarRAttribute As String 'local copy
Public Property Let RAttribute(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAttribute = 5
    mvarRAttribute = vData
End Property


Public Property Get RAttribute() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RAttribute
    RAttribute = mvarRAttribute
End Property



