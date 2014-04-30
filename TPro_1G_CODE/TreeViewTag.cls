VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeViewTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' Class Name:   cTags
' Author:       Graeme Grant
' Date:         24/01/2002
' Version:      01.00.00
' Description:  A Variant collection object - Designed to replace Type'd
'               variables/arrays.
' Edit History: 01.00.00 24/01/2002 Initial Release
'
'===========================================================================

Option Explicit

Private moTagCol As Collection

Public Sub Add(ByVal lPtr As Long, ParamArray Elements() As Variant)
    '
    '## lPtr = Object Pointer 'ObjPtr([..Object..)' or
    '          ID (Set ID higher than number of items to
    '              exist in the collection to avoid clashing).
    '
    If UBound(Elements) > 0 Then                '## Did we get anything?
        If Not Exist(lPtr) Then                 '## Does it already exist??
            moTagCol.Add Elements, CStr(lPtr)   '## New - Add to collection
        End If
    End If
End Sub

Public Sub Remove(ByVal lPtr As Long)
    If Exist(lPtr) Then                         '## Does it exist?
        moTagCol.Remove CStr(lPtr)              '## Yep - Remove it
    End If
End Sub

Public Function Exist(ByVal lPtr As Long) As Boolean
    On Error GoTo ErrorHandler
    Dim Tmp As Variant
    Tmp = moTagCol(CStr(lPtr))(0)               '## Can we get a value returned?
    Exist = True                                '## Then it must exist!
ErrorHandler:
                                                '## ERROR: Doesn't exist...
End Function

Public Property Get Count() As Long
    Count = moTagCol.Count
End Property

Public Property Get Element(ByVal lPtr As Long, ByVal ElementID As Long) As Variant
    Element = moTagCol(CStr(lPtr))(ElementID)
End Property

Public Property Let Element(ByVal lPtr As Long, ByVal ElementID As Long, ByVal NewData As Variant)
    moTagCol(CStr(lPtr))(ElementID) = NewData
End Property

Public Property Set Element(ByVal lPtr As Long, ByVal ElementID As Long, ByVal NewObj As Variant)
    Set moTagCol(CStr(lPtr))(ElementID) = NewObj
End Property

Public Function NewEnum() As IUnknown
    Set NewEnum = moTagCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set moTagCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set moTagCol = Nothing
End Sub

