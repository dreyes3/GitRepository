VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QMinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarItemNum As Integer 'local copy
Private mvarDateTime As Date 'local copy
Private mvarQueueNum As String 'local copy
Private mvarQueueCat As String 'local copy
Private mvarPCC As String 'local copy
Private mvarFreeText As String 'local copy
Public Property Let FreeText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FreeText = 5
    mvarFreeText = vData
End Property


Public Property Get FreeText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FreeText
    FreeText = mvarFreeText
End Property



Public Property Let PCC(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PCC = 5
    mvarPCC = vData
End Property


Public Property Get PCC() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PCC
    PCC = mvarPCC
End Property



Public Property Let QueueCat(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.QueueCat = 5
    mvarQueueCat = vData
End Property


Public Property Get QueueCat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.QueueCat
    QueueCat = mvarQueueCat
End Property



Public Property Let QueueNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.QueueNum = 5
    mvarQueueNum = vData
End Property


Public Property Get QueueNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.QueueNum
    QueueNum = mvarQueueNum
End Property



Public Property Let DateTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTime = 5
    mvarDateTime = vData
End Property


Public Property Get DateTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTime
    DateTime = mvarDateTime
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



