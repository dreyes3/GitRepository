Attribute VB_Name = "modListView"
      Option Explicit

      'Structures

      Public Type POINT
        x As Long
        y As Long
      End Type

      Public Type LV_FINDINFO
        flags As Long
        psz As String
        lParam As Long
        pt As POINT
        vkDirection As Long
      End Type

      Public Type LV_ITEM
        mask As Long
        iItem As Long
        iSubItem As Long
        State As Long
        stateMask As Long
        pszText As Long
        cchTextMax As Long
        iImage As Long
        lParam As Long
        iIndent As Long
      End Type

      'Constants
      Private Const LVFI_PARAM = 1
      Private Const LVIF_TEXT = &H1

      Private Const LVM_FIRST = &H1000
      Private Const LVM_FINDITEM = LVM_FIRST + 13
      Private Const LVM_GETITEMTEXT = LVM_FIRST + 45
      Public Const LVM_SORTITEMS = LVM_FIRST + 48
      
      'API declarations

      Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long


      'Module Functions and Procedures

      'CompareDates: This is the sorting routine that gets passed to the
      'ListView control to provide the comparison test for date values.

      Public Function CompareDates(ByVal lngParam1 As Long, _
                                   ByVal lngParam2 As Long, _
                                   ByVal hWnd As Long) As Long

        Dim strName1 As String
        Dim strName2 As String
        Dim dDate1 As Date
        Dim dDate2 As Date

        'Obtain the item names and dates corresponding to the
        'input parameters

        ListView_GetItemData lngParam1, hWnd, strName1, dDate1
        ListView_GetItemData lngParam2, hWnd, strName2, dDate2

        'Compare the dates
        'Return 0 ==> Less Than
        '       1 ==> Equal
        '       2 ==> Greater Than

        If dDate1 < dDate2 Then
          CompareDates = 0
        ElseIf dDate1 = dDate2 Then
          CompareDates = 1
        Else
          CompareDates = 2
        End If

      End Function

      'GetItemData - Given Retrieves

      Public Sub ListView_GetItemData(lngParam As Long, _
                                      hWnd As Long, _
                                      strName As String, _
                                      dDate As Date)
        Dim objFind As LV_FINDINFO
        Dim lngIndex As Long
        Dim objItem As LV_ITEM
        Dim baBuffer(32) As Byte
        Dim lngLength As Long

        '
        ' Convert the input parameter to an index in the list view
        '
        objFind.flags = LVFI_PARAM
        objFind.lParam = lngParam
        lngIndex = SendMessage(hWnd, LVM_FINDITEM, -1, VarPtr(objFind))

        '
        ' Obtain the name of the specified list view item
        '
        objItem.mask = LVIF_TEXT
        objItem.iSubItem = 0
        objItem.pszText = VarPtr(baBuffer(0))
        objItem.cchTextMax = UBound(baBuffer)
        lngLength = SendMessage(hWnd, LVM_GETITEMTEXT, lngIndex, _
                                VarPtr(objItem))
        strName = Left$(StrConv(baBuffer, vbUnicode), lngLength)

        '
        ' Obtain the modification date of the specified list view item
        '
        objItem.mask = LVIF_TEXT
        objItem.iSubItem = 1
        objItem.pszText = VarPtr(baBuffer(0))
        objItem.cchTextMax = UBound(baBuffer)
        lngLength = SendMessage(hWnd, LVM_GETITEMTEXT, lngIndex, _
                                VarPtr(objItem))
        If lngLength > 0 Then
          dDate = CDate(Left$(StrConv(baBuffer, vbUnicode), lngLength))
        End If

      End Sub

      'GetListItem - This is a modified version of ListView_GetItemData
      ' It takes an index into the list as a parameter and returns
      ' the appropriate values in the strName and dDate parameters.

      Public Sub ListView_GetListItem(lngIndex As Long, _
                                      hWnd As Long, _
                                      strName As String, _
                                      dDate As Date)
        Dim objItem As LV_ITEM
        Dim baBuffer(32) As Byte
        Dim lngLength As Long

        '
        ' Obtain the name of the specified list view item
        '
        objItem.mask = LVIF_TEXT
        objItem.iSubItem = 0
        objItem.pszText = VarPtr(baBuffer(0))
        objItem.cchTextMax = UBound(baBuffer)
        lngLength = SendMessage(hWnd, LVM_GETITEMTEXT, lngIndex, _
                                VarPtr(objItem))
        strName = Left$(StrConv(baBuffer, vbUnicode), lngLength)

        '
        ' Obtain the modification date of the specified list view item
        '
        objItem.mask = LVIF_TEXT
        objItem.iSubItem = 1
        objItem.pszText = VarPtr(baBuffer(0))
        objItem.cchTextMax = UBound(baBuffer)
        lngLength = SendMessage(hWnd, LVM_GETITEMTEXT, lngIndex, _
                                VarPtr(objItem))
        If lngLength > 0 Then
          dDate = CDate(Left$(StrConv(baBuffer, vbUnicode), lngLength))
        End If

      End Sub

