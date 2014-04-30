VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrModuleName As String
Private mstrProcedureName As String
Private mstrObjectName As String
Private mobjFileSys As Scripting.FileSystemObject
Private mobjTextStream As Scripting.TextStream
Private mstrFilePath As String
Private mstrFileName As String
Private mblnFileOpen As Boolean
Private mintFile As Integer

Public Property Get ModuleName() As String
    ModuleName = mstrModuleName
End Property

Public Property Let ModuleName(NameText As String)
    mstrModuleName = NameText
End Property

Public Property Get ProcedureName() As String
    ProcedureName = mstrProcedureName
End Property

Public Property Let ProcedureName(NameText As String)
    mstrProcedureName = NameText
End Property

Public Property Get ObjectName() As String
    ObjectName = mstrObjectName
End Property

Public Property Let ObjectName(NameText As String)
    mstrObjectName = NameText
End Property

Public Sub OpenLog(ByVal AppPath As String, ByVal appEXEName As String, ByVal AppTitle As String, _
    ByVal appMajor As Integer, appMinor As Integer, appRevision As Integer)
    
    
    Dim objFiles As Scripting.Files
    Dim objFile As Scripting.File
    Dim objFolder As Scripting.Folder
    Dim strAbsPathName As String
    Dim strErrDelete As String
    Dim strtemp
    strErrDelete = ""
    mstrFilePath = AppPath & "\Log"
    strAbsPathName = mobjFileSys.GetAbsolutePathName(mstrFilePath)
    If Not mobjFileSys.FolderExists(strAbsPathName) Then
        mobjFileSys.CreateFolder (mstrFilePath)
    Else
        Set objFolder = mobjFileSys.GetFolder(strAbsPathName)
        Set objFiles = objFolder.Files
        For Each objFile In objFiles

            If (objFile.DateLastModified < Now - 3) Then
                On Error Resume Next
                objFile.Delete (True)
                If Err.Number > 0 Then strErrDelete = strErrDelete & Err.Number & " - " & Err.Description & " (" & objFile.Name & ")" & vbCr
            End If
        Next
        Set objFile = Nothing
        Set objFiles = Nothing
        Set objFolder = Nothing
        
    End If
    
    mstrFileName = mstrFilePath & "\Tpro_" & Format(Now, "ddmmm_hhnnss") & ".log"
    
    'added on 4/2 to replace FSO OpenTextFile function
    mintFile = FreeFile()
    Open mstrFileName For Append As #mintFile
    Print #mintFile, AppTitle & " log | " & Format(Now, "dd-mmm-yyyy Hh:Nn:ss")
    Print #mintFile, AppPath & "\" & appEXEName & " (ver " & appMajor & "." & appMinor & "." & appRevision & ")"
    If strErrDelete <> "" Then
        Print #mintFile, "Init Error: " & strErrDelete
    End If
    'Set mobjTextStream = mobjFileSys.OpenTextFile(mstrFileName, ForWriting, True)
    'mobjTextStream.WriteLine AppTitle & " log | " & Format(Now, "dd-mmm-yyyy Hh:Nn:ss")
    'mobjTextStream.WriteLine AppPath & "\" & appEXEName & " (ver " & appMajor & "." & appMinor & "." & appRevision & ")"
    mblnFileOpen = True

Exit Sub
    
End Sub

Public Sub ErrorToLog(ErrSource As String, ErrNumber As String, ErrDescription As String)
Dim blnNewFile As Boolean

If Not mblnFileOpen Then
    Me.OpenLog App.Path, App.EXEName, App.Title, App.Major, App.Minor, App.Revision
    blnNewFile = True
End If
    
    Print #mintFile,
    Print #mintFile, ">> Error encountered at " & Format(Now, "dd-mmm-yyyy Hh:Nn:Ss")
    Print #mintFile, "Error Source: " & ErrSource
    Print #mintFile, "Module Name: " & Me.ModuleName
    Print #mintFile, "Procedure Name: " & Me.ProcedureName
    Print #mintFile, "Object Name: " & Me.ObjectName
    Print #mintFile,
    Print #mintFile, "Error " & ErrNumber & " (" & ErrDescription & ")"
    Print #mintFile, ">> End of Error."
    
    'added on 4/2 to replace FSO TextStream function
    'With mobjTextStream
    '    .WriteBlankLines (1)
    '    .WriteLine ">>>Beg Error>   " & Format(Now, "dd-mmm-yyyy Hh:Nn:Ss")
    '    .WriteLine Chr(9) & "Error " & ErrNumber & " (" & ErrDescription & ")"
    '    .WriteLine Chr(9) & ErrSource & " | " & Me.ModuleName & " | " & Me.ProcedureName & " | " & Me.ObjectName
    '    .WriteLine ">>>End Error>   "
    'End With

If blnNewFile Then Me.CloseLog

End Sub

Public Sub EventToLog(ByVal EventName As String, Optional ByVal BlankLineBefore As Integer = 0, _
    Optional ByVal BlankLineAfter As Integer = 0)
    
    If Not mblnFileOpen Then Err.Raise -100069, "AppLog", "No open log file"
    
    'Print #mintFile,
    'Print #mintFile, "[" & Format(Now, "dd-mmm-yyyy Hh:Nn:Ss") & "]   " & "Event Logging: " & EventName
    
    'added on 4/2 to replace FSO TextStream function
    'With mobjTextStream
    '    .WriteBlankLines (BlankLineBefore)
    '    .WriteLine ">>>" & EventName & ">   " & Format(Now, "ddmmmyy Hh:Nn:Ss")
    '    .WriteBlankLines (BlankLineAfter)
    'End With
End Sub

Public Sub LineTextToLog(LineText As String, Optional ByVal BlankLineBefore As Integer = 0, _
    Optional ByVal BlankLineAfter As Integer = 0)

If Not mblnFileOpen Then Err.Raise -100069, "AppLog", "No open log file"

'Print #mintFile,
Print #mintFile, "[" & Format(Now, "dd-mmm-yyyy Hh:Nn:Ss") & "]   " & LineText
'added on 4/2 to replace FSO TextStream function
'With mobjTextStream
'        .WriteBlankLines (BlankLineBefore)
'        .WriteLine LineText
'        .WriteBlankLines (BlankLineAfter)
'    End With
End Sub

Public Property Get LogOpen() As Boolean
    LogOpen = mblnFileOpen
End Property

Public Sub CloseLog()
On Error Resume Next
    Close #mintFile
    mobjTextStream.Close
    mblnFileOpen = False
End Sub

Private Sub Class_Initialize()
    Set mobjFileSys = New Scripting.FileSystemObject
End Sub

Private Sub Class_Terminate()
    Set mobjFileSys = Nothing
End Sub

Public Property Get FilePath() As String
    FilePath = mstrFilePath
End Property

Public Property Get FileName() As String
    FileName = mstrFileName
End Property

'Public Sub XMLFileToLog(ByVal FileName As String, ByVal XMLStream As String)
'    Dim objXMLStream As Scripting.TextStream
'
'    Set objXMLStream = mobjFileSys.OpenTextFile(FileName, ForWriting, True)
'    objXMLStream.Write XMLStream
'    objXMLStream.Close
'    Set objXMLStream = Nothing
'
'Call EventToLog(">>>XML File " & FileName & " has been written", 1, 1)
'End Sub

