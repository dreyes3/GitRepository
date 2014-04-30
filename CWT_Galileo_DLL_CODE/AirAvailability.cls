VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AirAvailability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mbytNumSeats As Byte 'local copy
Private mvarItem As Long 'local copy
Private mstrAvailType As String 'local copy
Private mstrOrig As String 'local copy
Private mstrDest As String 'local copy
Private mdtmDeptDate As Date 'local copy

Private mcolJourneys As Collection
Public Function AddJourney(TotJrnyTime As Long, Optional FQNumber As Byte = 0) As Journey
Dim objJny As Journey
Set objJny = New Journey

With objJny
    .FareIndex = FQNumber
    .JourneyTime = TotJrnyTime
End With
mcolJourneys.Add objJny
Set objJny = Nothing
End Function


Public Property Let DeptDate(ByVal DepartDate As Date)
    mdtmDeptDate = DepartDate
End Property

Public Property Get DeptDate() As Date
    DeptDate = mdtmDeptDate
End Property

Public Property Get JourneyCount() As Integer
    JourneyCount = mcolJourneys.Count
End Property

Public Property Let Dest(ByVal DestAptCode As String)
    mstrDest = DestAptCode
End Property

Public Property Get Dest() As String
    Dest = mstrDest
End Property

Public Property Let Orig(ByVal OrigAptCode As String)
    mstrOrig = OrigAptCode
End Property

Public Property Get Orig() As String
    Orig = mstrOrig
End Property

Public Property Let AvailType(ByVal GDSAvailType As String)
    mstrAvailType = GDSAvailType
End Property

Public Property Get AvailType() As String
    AvailType = mstrAvailType
End Property

Public Function Item(ByVal Index As Variant) As Journey
    Set Item = mcolJourneys(Index)
End Function

Public Property Let NumSeats(ByVal NumberSeats As Byte)
    mbytNumSeats = NumberSeats
End Property

Public Property Get NumSeats() As Byte
    NumSeats = mbytNumSeats
End Property

Private Sub Class_Initialize()
    Set mcolJourneys = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolJourneys = Nothing
End Sub
