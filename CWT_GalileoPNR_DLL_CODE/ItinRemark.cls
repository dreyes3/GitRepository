VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItinRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarItemNum As Integer 'local copy
Private mvarSegNum As Integer 'local copy
Private mvarRemarkText As String 'local copy
Public Property Let RemarkText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemarkText = 5
    mvarRemarkText = vData
End Property


Public Property Get RemarkText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemarkText
    RemarkText = mvarRemarkText
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



Public Property Let ItemNum(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemNum = 5
    mvarItemNum = vData
End Property


Public Property Get ItemNum() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemNum
    ItemNum = mvarItemNum
End Property



