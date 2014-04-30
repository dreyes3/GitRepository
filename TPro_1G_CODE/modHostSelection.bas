Attribute VB_Name = "modHostSelection"
' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
'******************For IPAddress******************
Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type

Public Declare Function WSAGetLastError Lib "wsock32" () As Long

Public Declare Function WSAStartup Lib "wsock32" _
  (ByVal wVersionRequired As Long, _
   lpWSADATA As WSADATA) As Long
   
Public Declare Function WSACleanup Lib "wsock32" () As Long

Public Declare Function GetHostname Lib "wsock32" _
  Alias "gethostname" (ByVal szHost As _
   String, ByVal dwHostLen As Long) As Long
   
Public Declare Function gethostbyname Lib "wsock32" _
  (ByVal szHost As String) As Long
   
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (hpvDest As Any, _
   ByVal hpvSource As Long, _
   ByVal cbCopy As Long)
'***********************************************
Public Enum gModuleType
    [PC] = 0 ' Point and Click
    [SP] = 1 ' Smart Point with Full module
    [SYEX] = 2 ' Symphonie Express
End Enum



Public gstrIPAddress As String
Public gstrHostName As String
'Public gcnCWTApplication As ADODB.Connection


'***********************************************
' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
' To determine appropriate GUI that user require to use
Public Function GUISelection() As Integer
Dim iModuleValue As Integer

'set default as 0
iModuleValue = 0

gstrHostName = GetServerHostname
iModuleValue = getModuleValue(gstrHostName)
GUISelection = iModuleValue

End Function

Public Function GetServerHostname() As String

   Dim sTempHostName    As String * 256
   Dim sHostname As String
   Dim lpHost    As Long
   Dim HOST      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   
   If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
   End If
    
  'gethostname returns the name of the local host into
  'the buffer specified by the name parameter. The host
  'name is returned as a null-terminated string. The
  'form of the host name is dependent on the Windows
  'Sockets provider - it can be a simple host name, or
  'it can be a fully qualified domain name. However, it
  'is guaranteed that the name returned will be successfully
  'parsed by gethostbyname and WSAAsyncGetHostByName.

  'In actual application, if no local host name has been
  'configured, gethostname must succeed and return a token
  'host name that gethostbyname or WSAAsyncGetHostByName
  'can resolve.
   If GetHostname(sTempHostName, 256) = SOCKET_ERROR Then
      GetIPAddress = ""
      MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
              " has occurred. Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   
  'gethostbyname returns a pointer to a HOSTENT structure
  '- a structure allocated by Windows Sockets. The HOSTENT
  'structure contains the results of a successful search
  'for the host specified in the name parameter.

  'The application must never attempt to modify this
  'structure or to free any of its components. Furthermore,
  'only one copy of this structure is allocated per thread,
  'so the application should copy any information it needs
  'before issuing any other Windows Sockets function calls.

  'gethostbyname function cannot resolve IP address strings
  'passed to it. Such a request is treated exactly as if an
  'unknown host name were passed. Use inet_addr to convert
  'an IP address string the string to an actual IP address,
  'then use another function, gethostbyaddr, to obtain the
  'contents of the HOSTENT structure.
   sTempHostName = Trim$(sTempHostName)
   lpHost = gethostbyname(sTempHostName)
    
   If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Windows Sockets are not responding. " & _
              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   
   sHostname = UCase(Mid(sTempHostName, 1, InStr(1, sTempHostName, Chr(0)) - 1))
   GetServerHostname = sHostname
   
  'to extract the returned IP address, we have to copy
  'the HOST structure and its members
   'CopyMemory HOST, lpHost, Len(HOST)
   'CopyMemory dwIPAddr, HOST.hAddrList, 4
   
  'create an array to hold the result
   'ReDim tmpIPAddr(1 To HOST.hLen)
   'CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
   
  'and with the array, build the actual address,
  'appending a period between members
   'For i = 1 To HOST.hLen
   '   sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   'Next
  
  'the routine adds a period to the end of the
  'string, so remove it here
   'GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   SocketsCleanup
    
End Function

Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function

Public Sub SocketsCleanup()

    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub
Public Function HiByte(ByVal wParam As Integer) As Byte
  
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function


Public Function LoByte(ByVal wParam As Integer) As Byte

  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function


Public Function getModuleValue(ByVal strHostName) As Integer
Dim rsModuleValue As New ADODB.Recordset
Dim strSQL As String
Dim iModuleValue As Integer

iModuleValue = 0

strSQL = "SELECT optionValue FROM tblHostSetting where SessionID = '" & Trim(strHostName) & "'"
RunSQLCommand SQLType.Select_, strSQL, gdbAPPConn, rsModuleValue


    If rsModuleValue.EOF = False Then
        iModuleValue = rsModuleValue!optionvalue
    End If
    
    getModuleValue = iModuleValue

End Function

 ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
Public Function SwitchWinSetting(lFormhwnd As Long)
   
    If gIntModuleType <> gModuleType.PC Then
         MakeWinTopMost (lFormhwnd)
    Else
        oldParent = SetParent(lFormhwnd, gVPMDIHwnd)
    End If
    
End Function
