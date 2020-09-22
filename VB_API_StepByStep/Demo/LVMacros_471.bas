Attribute VB_Name = "mListviewMacros_471"
Option Explicit
'
' Brad Martinez http://www.mvps.org
'
' change history:
'   ** note: modules dated 08/16/98 do *not* have the changes below applied **
'
'   06/18/98 - added "lvi.pszText = pszText" to ListView_GetItemText
'                    - uncommented conditional compilation code
'
'   11/17/98 - added "As Long" to "i" param in ListView_EnsureVisible
'                   - added #If (WIN32_IE >= &H300) to ListView_GetHeader
'
'   11/26/98 - added user-defined macros: ListView_GetSelectedItem, ListView_SetSelectedItem, ListView_SelectAll
'                   - added enum data types defined in mListviewDefs_471 to macro params
'                   - changed ListView_GetCallbackMask from Boolean to Long
'
'   06/23/99 - added private defs
'
' Set the following conditional compilation constant accordingly in the
' Make tab of the project Properties dialog box:
'
'  WIN32_IE < 768   (&H300, or not defined): don't include IE3 and IE4 listview macros
'  WIN32_IE = 768   (&H300): include only the IE3 listview macros
'  WIN32_IE = 1024 (&H400)  include both the IE3 and IE4 listview macros
'
' =============================================================================
' Listview control macros extracted and translated directly from Commctrl.h v1.2  (Comctl32.dll v4.71)
' =============================================================================
'
'     82 macros total: 52-v4.00.950 (IE2), 15-v4.70 (IE3), 14-v4.71 (IE4), 1-user-defined
'
'#Const WIN32_IE = &H400

'Private Type POINTAPI   ' pt
 ' x As Long
  'y As Long
'End Type

'Private Type RECT   ' rc
  'Left As Long
  'Top As Long
  'Right As Long
  'Bottom As Long
'End Type

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---


' Returns the low 16-bit integer from a 32-bit long integer
Public Function LOWORD(dwValue As Long) As Integer
  MoveMemory LOWORD, dwValue, 2
End Function

' Returns the low 16-bit integer from a 32-bit long integer

Public Function HIWORD(dwValue As Long) As Integer
  MoveMemory HIWORD, ByVal VarPtr(dwValue) + 2, 2
End Function

' Combines two integers into a long integer

Public Function MAKELONG(wLow As Long, wHigh As Long) As Long
  MAKELONG = wLow
  MoveMemory ByVal VarPtr(MAKELONG) + 2, wHigh, 2
End Function

' Combines two integers into a long integer

Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
  MAKELPARAM = MAKELONG(wLow, wHigh)
End Function

' =============================================================================

Public Function ListView_GetBkColor(hWnd As Long) As Long
  ListView_GetBkColor = SendMessage(hWnd, LVM_GETBKCOLOR, 0, 0)
End Function
 
Public Function ListView_SetBkColor(hWnd As Long, clrBk As Long) As Boolean
  ListView_SetBkColor = SendMessage(hWnd, LVM_SETBKCOLOR, 0, ByVal clrBk)
End Function
 
Public Function ListView_GetImageList(hWnd As Long, iImageList As LVSIL_Flags) As Long
  ListView_GetImageList = SendMessage(hWnd, LVM_GETIMAGELIST, ByVal iImageList, 0)
End Function

Public Function ListView_SetImageList(hWnd As Long, himl As Long, iImageList As Long) As Long
  ListView_SetImageList = SendMessage(hWnd, LVM_SETIMAGELIST, ByVal iImageList, ByVal himl)
End Function
 
Public Function ListView_GetItemCount(hWnd As Long) As Long
  ListView_GetItemCount = SendMessage(hWnd, LVM_GETITEMCOUNT, 0, 0)
End Function
 
Public Function ListView_GetItem(hWnd As Long, pitem As LVITEM) As Boolean
  ListView_GetItem = SendMessage(hWnd, LVM_GETITEM, 0, pitem)
End Function
 
Public Function ListView_SetItem(hWnd As Long, pitem As LVITEM) As Boolean
  ListView_SetItem = SendMessage(hWnd, LVM_SETITEM, 0, pitem)
End Function
 
Public Function ListView_InsertItem(hWnd As Long, pitem As LVITEM) As Long
  ListView_InsertItem = SendMessage(hWnd, LVM_INSERTITEM, 0, pitem)
End Function
 
Public Function ListView_DeleteItem(hWnd As Long, i As Long) As Boolean
  ListView_DeleteItem = SendMessage(hWnd, LVM_DELETEITEM, ByVal i, 0)
End Function
 
Public Function ListView_DeleteAllItems(hWnd As Long) As Boolean
  ListView_DeleteAllItems = SendMessage(hWnd, LVM_DELETEALLITEMS, 0, 0)
End Function
 
Public Function ListView_GetCallbackMask(hWnd As Long) As Long   ' LVStyles
  ListView_GetCallbackMask = SendMessage(hWnd, LVM_GETCALLBACKMASK, 0, 0)
End Function
 
Public Function ListView_SetCallbackMask(hWnd As Long, mask As LVStyles) As Boolean
  ListView_SetCallbackMask = SendMessage(hWnd, LVM_SETCALLBACKMASK, ByVal mask, 0)
End Function
 
Public Function ListView_GetNextItem(hWnd As Long, i As Long, flags As LVNI_Flags) As Long
  ListView_GetNextItem = SendMessage(hWnd, LVM_GETNEXTITEM, ByVal i, ByVal MAKELPARAM(flags, 0))
End Function

Public Function ListView_FindItem(hWnd As Long, iStart, plvfi As LVFINDINFO) As Long
  ListView_FindItem = SendMessage(hWnd, LVM_FINDITEM, ByVal iStart, plvfi)
End Function
 
Public Function ListView_GetItemRect(hWnd As Long, i As Long, prc As RECT, code As LVIR_Flags) As Boolean
  prc.Left = code
  ListView_GetItemRect = SendMessage(hWnd, LVM_GETITEMRECT, ByVal i, prc)
End Function

Public Function ListView_SetItemPosition(hwndLV As Long, i As Long, x As Long, y As Long) As Boolean
  ListView_SetItemPosition = SendMessage(hwndLV, LVM_SETITEMPOSITION, ByVal i, ByVal MAKELPARAM(x, y))
End Function
 
Public Function ListView_GetItemPosition(hwndLV As Long, i As Long, ppt As POINTAPI) As Boolean
  ListView_GetItemPosition = SendMessage(hwndLV, LVM_GETITEMPOSITION, ByVal i, ppt)
End Function
 
Public Function ListView_GetStringWidth(hwndLV As Long, psz As String) As Long
  ListView_GetStringWidth = SendMessage(hwndLV, LVM_GETSTRINGWIDTH, 0, ByVal psz)
End Function
 
Public Function ListView_HitTest(hwndLV As Long, pinfo As LVHITTESTINFO) As Long
  ListView_HitTest = SendMessage(hwndLV, LVM_HITTEST, 0, pinfo)
End Function
 
Public Function ListView_EnsureVisible(hwndLV As Long, i As Long, fPartialOK As Boolean) As Boolean
  ListView_EnsureVisible = SendMessage(hwndLV, LVM_ENSUREVISIBLE, ByVal i, ByVal MAKELPARAM(Abs(fPartialOK), 0))
End Function
 
Public Function ListView_Scroll(hwndLV As Long, dx As Long, dy As Long) As Boolean
  ListView_Scroll = SendMessage(hwndLV, LVM_SCROLL, ByVal dx, ByVal dy)
End Function
 
Public Function ListView_RedrawItems(hwndLV As Long, iFirst, iLast) As Boolean
  ListView_RedrawItems = SendMessage(hwndLV, LVM_REDRAWITEMS, ByVal iFirst, ByVal iLast)
End Function
 
Public Function ListView_Arrange(hwndLV As Long, code As LVA_Flags) As Boolean
  ListView_Arrange = SendMessage(hwndLV, LVM_ARRANGE, ByVal code, 0)
End Function
 
Public Function ListView_EditLabel(hwndLV As Long, i As Long) As Long
  ListView_EditLabel = SendMessage(hwndLV, LVM_EDITLABEL, ByVal i, 0)
End Function
 
Public Function ListView_GetEditControl(hwndLV As Long) As Long
  ListView_GetEditControl = SendMessage(hwndLV, LVM_GETEDITCONTROL, 0, 0)
End Function
 
Public Function ListView_GetColumn(hWnd As Long, iCol As Long, pcol As LVCOLUMN) As Boolean
  ListView_GetColumn = SendMessage(hWnd, LVM_GETCOLUMN, ByVal iCol, pcol)
End Function
 
Public Function ListView_SetColumn(hWnd As Long, iCol As Long, pcol As LVCOLUMN) As Boolean
  ListView_SetColumn = SendMessage(hWnd, LVM_SETCOLUMN, ByVal iCol, pcol)
End Function
 
Public Function ListView_InsertColumn(hWnd As Long, iCol As Long, pcol As LVCOLUMN) As Long
  ListView_InsertColumn = SendMessage(hWnd, LVM_INSERTCOLUMN, ByVal iCol, pcol)
End Function
 
Public Function ListView_DeleteColumn(hWnd As Long, iCol As Long) As Boolean
  ListView_DeleteColumn = SendMessage(hWnd, LVM_DELETECOLUMN, ByVal iCol, 0)
End Function
 
Public Function ListView_GetColumnWidth(hWnd As Long, iCol As Long) As Long
  ListView_GetColumnWidth = SendMessage(hWnd, LVM_GETCOLUMNWIDTH, ByVal iCol, 0)
End Function
 
Public Function ListView_SetColumnWidth(hWnd As Long, iCol As Long, cx As Long) As Boolean
  ListView_SetColumnWidth = SendMessage(hWnd, LVM_SETCOLUMNWIDTH, ByVal iCol, ByVal MAKELPARAM(cx, 0))
End Function
 
#If (WIN32_IE >= &H300) Then

Public Function ListView_GetHeader(hWnd As Long) As Long
  ListView_GetHeader = SendMessage(hWnd, LVM_GETHEADER, 0, 0)
End Function
'
#End If
 
Public Function ListView_CreateDragImage(hWnd As Long, i As Long, lpptUpLeft As POINTAPI) As Long
  ListView_CreateDragImage = SendMessage(hWnd, LVM_CREATEDRAGIMAGE, ByVal i, lpptUpLeft)
End Function
 
Public Function ListView_GetViewRect(hWnd As Long, prc As RECT) As Boolean
  ListView_GetViewRect = SendMessage(hWnd, LVM_GETVIEWRECT, 0, prc)
End Function
 
Public Function ListView_GetTextColor(hWnd As Long) As Long
  ListView_GetTextColor = SendMessage(hWnd, LVM_GETTEXTCOLOR, 0, 0)
End Function
 
Public Function ListView_SetTextColor(hWnd As Long, clrText As Long) As Boolean
  ListView_SetTextColor = SendMessage(hWnd, LVM_SETTEXTCOLOR, 0, ByVal clrText)
End Function
 
Public Function ListView_GetTextBkColor(hWnd As Long) As Long
  ListView_GetTextBkColor = SendMessage(hWnd, LVM_GETTEXTBKCOLOR, 0, 0)
End Function
 
Public Function ListView_SetTextBkColor(hWnd As Long, clrTextBk As Long) As Boolean
  ListView_SetTextBkColor = SendMessage(hWnd, LVM_SETTEXTBKCOLOR, 0, ByVal clrTextBk)
End Function
 
Public Function ListView_GetTopIndex(hwndLV As Long) As Long
  ListView_GetTopIndex = SendMessage(hwndLV, LVM_GETTOPINDEX, 0, 0)
End Function
 
Public Function ListView_GetCountPerPage(hwndLV As Long) As Long
  ListView_GetCountPerPage = SendMessage(hwndLV, LVM_GETCOUNTPERPAGE, 0, 0)
End Function
 
Public Function ListView_GetOrigin(hwndLV As Long, ppt As POINTAPI) As Boolean
  ListView_GetOrigin = SendMessage(hwndLV, LVM_GETORIGIN, 0, ppt)
End Function
 
Public Function ListView_Update(hwndLV As Long, i As Long) As Boolean
  ListView_Update = SendMessage(hwndLV, LVM_UPDATE, ByVal i, 0)
End Function
 
Public Function ListView_SetItemState(hwndLV As Long, i As Long, state As LVITEM_state, mask As LVITEM_state) As Boolean
  Dim lvi As LVITEM
  lvi.state = state
  lvi.stateMask = mask
  ListView_SetItemState = SendMessage(hwndLV, LVM_SETITEMSTATE, ByVal i, lvi)
End Function
 
Public Function ListView_GetItemState(hwndLV As Long, i As Long, mask As LVITEM_state) As Long   ' LVITEM_state
  ListView_GetItemState = SendMessage(hwndLV, LVM_GETITEMSTATE, ByVal i, ByVal mask)
End Function

#If (WIN32_IE >= &H300) Then

Public Function ListView_GetCheckState(hwndLV As Long, iIndex As Long) As Long   ' updated
  Dim dwState As Long
  dwState = SendMessage(hwndLV, LVM_GETITEMSTATE, ByVal iIndex, ByVal LVIS_STATEIMAGEMASK)
  ListView_GetCheckState = (dwState \ 2 ^ 12) - 1
  '((((UINT)(SendMessage(hwndLV, LVM_GETITEMSTATE, ByVal i, LVIS_STATEIMAGEMASK))) >> 12) -1)
End Function
'
#End If

Public Sub ListView_GetItemText(hwndLV As Long, i As Long, iSubItem As Long, _
                                                     pszText As Long, cchTextMax As Long)
  Dim lvi As LVITEM
  lvi.iSubItem = iSubItem
  lvi.cchTextMax = cchTextMax
  lvi.pszText = pszText
  SendMessage hwndLV, LVM_GETITEMTEXT, ByVal i, lvi
  pszText = lvi.pszText   ' fills pszText w/ pointer
End Sub
 
Public Sub ListView_SetItemText(hwndLV As Long, i As Long, iSubItem As Long, pszText As Long)
  Dim lvi As LVITEM
  lvi.iSubItem = iSubItem
  lvi.pszText = pszText
  SendMessage hwndLV, LVM_SETITEMTEXT, ByVal i, lvi
End Sub

Public Sub ListView_SetItemCount(hwndLV As Long, cItems As Long)
  SendMessage hwndLV, LVM_SETITEMCOUNT, ByVal cItems, 0
End Sub

#If (WIN32_IE >= &H300) Then

Public Sub ListView_SetItemCountEx(hwndLV As Long, cItems As Long, dwFlags As Long)
  SendMessage hwndLV, LVM_SETITEMCOUNT, ByVal cItems, ByVal dwFlags
End Sub
'
#End If

Public Function ListView_SortItems(hwndLV As Long, pfnCompare As Long, lParamSort As Long) As Boolean
  ListView_SortItems = SendMessage(hwndLV, LVM_SORTITEMS, ByVal lParamSort, ByVal pfnCompare)
End Function
 
Public Sub ListView_SetItemPosition32(hwndLV As Long, i As Long, x As Long, y As Long)
  Dim ptNewPos As POINTAPI
  ptNewPos.x = x
  ptNewPos.y = y
  SendMessage hwndLV, LVM_SETITEMPOSITION32, ByVal i, ptNewPos
End Sub
 
Public Function ListView_GetSelectedCount(hwndLV As Long) As Long
  ListView_GetSelectedCount = SendMessage(hwndLV, LVM_GETSELECTEDCOUNT, 0, 0)
End Function
 
Public Function ListView_GetItemSpacing(hwndLV As Long, fSmall As Boolean) As Long
  ListView_GetItemSpacing = SendMessage(hwndLV, LVM_GETITEMSPACING, ByVal fSmall, 0)
End Function

Public Function ListView_GetISearchString(hwndLV As Long, lpsz As String) As Boolean
  ListView_GetISearchString = SendMessage(hwndLV, LVM_GETISEARCHSTRING, 0, ByVal lpsz)
End Function

' =============================================================
' the next three macros are user-defined

' Returns the index of the item that is selected and has the focus rectangle

Public Function ListView_GetSelectedItem(hwndLV As Long) As Long
  ListView_GetSelectedItem = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED Or LVNI_SELECTED)
End Function
 
' Selects the specified item and gives it the focus rectangle.
' does not de-select any currently selected items

Public Function ListView_SetSelectedItem(hwndLV As Long, i As Long) As Boolean
  ListView_SetSelectedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, _
                                                                                                     LVIS_FOCUSED Or LVIS_SELECTED)
End Function
 
' Selects all listview items. The item with the focus rectangle maintains it.

Public Function ListView_SelectAll(hwndLV As Long) As Boolean
  ListView_SelectAll = ListView_SetItemState(hwndLV, -1, LVIS_SELECTED, LVIS_SELECTED)
End Function

' ==============================================================
#If (WIN32_IE >= &H300) Then
'
' // -1 for cx and cy means we'll use the default (system settings)
' // 0 for cx or cy means use the current setting (allows you to change just one param)
Public Function ListView_SetIconSpacing(hwndLV As Long, cx As Long, cy As Long) As Long
  ListView_SetIconSpacing = SendMessage(hwndLV, LVM_SETICONSPACING, 0, ByVal MAKELONG(cx, cy))
End Function
 
Public Function ListView_SetExtendedListViewStyle(hwndLV As Long, dw As Long) As Long
  ListView_SetExtendedListViewStyle = SendMessage(hwndLV, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal dw)
End Function

Public Function ListView_GetExtendedListViewStyle(hwndLV As Long) As Long
  ListView_GetExtendedListViewStyle = SendMessage(hwndLV, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
End Function

Public Function ListView_GetSubItemRect(hWnd As Long, iItem As Long, iSubItem As Long, _
                                                              code As Long, prc As RECT) As Boolean
  prc.Top = iSubItem
  prc.Left = code
  ListView_GetSubItemRect = SendMessage(hWnd, LVM_GETSUBITEMRECT, ByVal iItem, prc)
End Function
 
Public Function ListView_SubItemHitTest(hWnd As Long, plvhti As LVHITTESTINFO) As Long
  ListView_SubItemHitTest = SendMessage(hWnd, LVM_SUBITEMHITTEST, 0, plvhti)
End Function
 
Public Function ListView_SetColumnOrderArray(hWnd As Long, iCount As Long, lpiArray As Long) As Boolean
  ListView_SetColumnOrderArray = SendMessage(hWnd, LVM_SETCOLUMNORDERARRAY, ByVal iCount, lpiArray)
End Function

Public Function ListView_GetColumnOrderArray(hWnd As Long, iCount As Long, lpiArray As Long) As Boolean
  ListView_GetColumnOrderArray = SendMessage(hWnd, LVM_GETCOLUMNORDERARRAY, ByVal iCount, lpiArray)
End Function
 
Public Function ListView_SetHotItem(hWnd As Long, i As Long) As Long
  ListView_SetHotItem = SendMessage(hWnd, LVM_SETHOTITEM, ByVal i, 0)
End Function
 
Public Function ListView_GetHotItem(hWnd As Long) As Long
  ListView_GetHotItem = SendMessage(hWnd, LVM_GETHOTITEM, 0, 0)
End Function
 
Public Function ListView_SetHotCursor(hWnd As Long, hcur As Long) As Long
  ListView_SetHotCursor = SendMessage(hWnd, LVM_SETHOTCURSOR, 0, ByVal hcur)
End Function
 
Public Function ListView_GetHotCursor(hWnd As Long) As Long
  ListView_GetHotCursor = SendMessage(hWnd, LVM_GETHOTCURSOR, 0, 0)
End Function
 
Public Function ListView_ApproximateViewRect(hWnd As Long, iWidth As Long, _
                                                                      iHeight As Long, iCount As Long) As Long
  ListView_ApproximateViewRect = SendMessage(hWnd, _
                                                                          LVM_APPROXIMATEVIEWRECT, _
                                                                          ByVal iCount, _
                                                                          ByVal MAKELPARAM(iWidth, iHeight))
End Function
'
#End If  ' ' WIN32_IE >= &H300
'

' ==============================================================
#If (WIN32_IE >= &H400) Then

Public Function ListView_SetUnicodeFormat(hWnd As Long, fUnicode As Boolean) As Boolean
  ListView_SetUnicodeFormat = SendMessage(hWnd, LVM_SETUNICODEFORMAT, ByVal fUnicode, 0)
End Function

Public Function ListView_GetUnicodeFormat(hWnd As Long) As Boolean
  ListView_GetUnicodeFormat = SendMessage(hWnd, LVM_GETUNICODEFORMAT, 0, 0)
End Function

Public Function ListView_SetExtendedListViewStyleEx(hwndLV As Long, dwMask As Long, dw As Long) As Long
  ListView_SetExtendedListViewStyleEx = SendMessage(hwndLV, LVM_SETEXTENDEDLISTVIEWSTYLE, _
                                                                                    ByVal dwMask, ByVal dw)
End Function
 
Public Function ListView_SetWorkAreas(hWnd As Long, nWorkAreas As Long, prc() As RECT) As Boolean
  ListView_SetWorkAreas = SendMessage(hWnd, LVM_SETWORKAREAS, ByVal nWorkAreas, prc(0))
End Function

Public Function ListView_GetWorkAreas(hWnd As Long, nWorkAreas, prc() As RECT) As Boolean
  ListView_GetWorkAreas = SendMessage(hWnd, LVM_GETWORKAREAS, ByVal nWorkAreas, prc(0))
End Function

Public Function ListView_GetNumberOfWorkAreas(hWnd As Long, pnWorkAreas As Long) As Boolean
  ListView_GetNumberOfWorkAreas = SendMessage(hWnd, LVM_GETNUMBEROFWORKAREAS, 0, pnWorkAreas)
End Function

Public Function ListView_GetSelectionMark(hWnd As Long) As Long
  ListView_GetSelectionMark = SendMessage(hWnd, LVM_GETSELECTIONMARK, 0, 0)
End Function

Public Function ListView_SetSelectionMark(hWnd As Long, i As Long) As Long
  ListView_SetSelectionMark = SendMessage(hWnd, LVM_SETSELECTIONMARK, 0, ByVal i)
End Function

Public Function ListView_SetHoverTime(hwndLV As Long, dwHoverTimeMs As Long) As Long
  ListView_SetHoverTime = SendMessage(hwndLV, LVM_SETHOVERTIME, 0, ByVal dwHoverTimeMs)
End Function

Public Function ListView_GetHoverTime(hwndLV As Long) As Long
  ListView_GetHoverTime = SendMessage(hwndLV, LVM_GETHOVERTIME, 0, 0)
End Function

Public Function ListView_SetToolTips(hwndLV As Long, hwndNewHwnd As Long) As Long
  ListView_SetToolTips = SendMessage(hwndLV, LVM_SETTOOLTIPS, ByVal hwndNewHwnd, 0)
End Function

Public Function ListView_GetToolTips(hwndLV As Long) As Long
  ListView_GetToolTips = SendMessage(hwndLV, LVM_GETTOOLTIPS, 0, 0)
End Function

Public Function ListView_SetBkImage(hWnd As Long, plvbki As LVBKIMAGE) As Boolean
  ListView_SetBkImage = SendMessage(hWnd, LVM_SETBKIMAGE, 0, plvbki)
End Function

Public Function ListView_GetBkImage(hWnd As Long, plvbki As LVBKIMAGE) As Boolean
  ListView_GetBkImage = SendMessage(hWnd, LVM_GETBKIMAGE, 0, plvbki)
End Function

#End If     ' WIN32_IE >= &H400
