Attribute VB_Name = "mListviewDefs_471"
Option Explicit
'
' Brad Martinez http://www.mvps.org
'
' 11/25/98 - updated notification msg comments, NMFINDITEM to NMLVFINDITEM
'                 -  renamed a *lot* of the enums....
' 05-13-99  - added private window defs
'
' Set the following conditional compilation constant accordingly in the
' Make tab of the project Properties dialog box:
'
'  WIN32_IE < 768   (&H300, or not defined): don't include IE3 and IE4 listview definitions
'  WIN32_IE = 768   (&H300): include only the IE3 listview definitions
'  WIN32_IE = 1024 (&H400)  include both the IE3 and IE4 listview definitions
'
' =============================================================================

#Const WIN32_IE = &H400

'Private Type POINTAPI   ' pt
  'x As Long
  'y As Long
'End Type
'
'' The NMHDR structure contains information about a notification message. The pointer
'' to this structure is specified as the lParam member of the WM_NOTIFY message.
'Private Type NMHDR
  'hwndFrom As Long   ' Window handle of control sending message
  'idFrom As Long        ' Identifier of control sending message
  'code  As Long          ' Specifies the notification code
'End Type

#If (WIN32_IE >= &H400) Then
Private Const CCM_FIRST = &H2000
Private Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
#End If    ' WIN32_IE >= 0x0400

' =============================================================================
' Listview control macros extracted and translated directly from Commctrl.h v1.2  (Comctl32.dll v4.71)
' =============================================================================
' Creation

Public Const WC_LISTVIEW = "SysListView32"
 
Public Enum LVStyles
  LVS_ICON = &H0
  LVS_REPORT = &H1
  LVS_SMALLICON = &H2
  LVS_LIST = &H3
  LVS_TYPEMASK = &H3
  LVS_SINGLESEL = &H4
  LVS_SHOWSELALWAYS = &H8
  LVS_SORTASCENDING = &H10
  LVS_SORTDESCENDING = &H20
  LVS_SHAREIMAGELISTS = &H40
  LVS_NOLABELWRAP = &H80
  LVS_AUTOARRANGE = &H100
  LVS_EDITLABELS = &H200

#If (WIN32_IE >= &H300) Then
  LVS_OWNERDATA = &H1000
#End If
  
  LVS_NOSCROLL = &H2000
 
  LVS_TYPESTYLEMASK = &HFC00
 
  LVS_ALIGNTOP = &H0
  LVS_ALIGNLEFT = &H800
  LVS_ALIGNMASK = &HC00
 
  LVS_OWNERDRAWFIXED = &H400
  LVS_NOCOLUMNHEADER = &H4000
  LVS_NOSORTHEADER = &H8000&
End Enum   ' LVStyles

' ============================================
' Messages
   
Public Enum LVMessages
  LVM_FIRST = &H1000

  LVM_GETBKCOLOR = (LVM_FIRST + 0)
  LVM_SETBKCOLOR = (LVM_FIRST + 1)
  LVM_GETIMAGELIST = (LVM_FIRST + 2)
 
  LVM_SETIMAGELIST = (LVM_FIRST + 3)
  LVM_GETITEMCOUNT = (LVM_FIRST + 4)
 
#If UNICODE Then
  LVM_GETITEM = (LVM_FIRST + 75)
  LVM_SETITEM = (LVM_FIRST + 76)
  LVM_INSERTITEM = (LVM_FIRST + 77)
#Else
  LVM_GETITEM = (LVM_FIRST + 5)
  LVM_SETITEM = (LVM_FIRST + 6)
  LVM_INSERTITEM = (LVM_FIRST + 7)
#End If
 
  LVM_DELETEITEM = (LVM_FIRST + 8)
  LVM_DELETEALLITEMS = (LVM_FIRST + 9)

  LVM_GETCALLBACKMASK = (LVM_FIRST + 10)
  LVM_SETCALLBACKMASK = (LVM_FIRST + 11)
  
  LVM_GETNEXTITEM = (LVM_FIRST + 12)
 
#If UNICODE Then
  LVM_FINDITEM = (LVM_FIRST + 83)
#Else
  LVM_FINDITEM = (LVM_FIRST + 13)
#End If
 
  LVM_GETITEMRECT = (LVM_FIRST + 14)
  LVM_SETITEMPOSITION = (LVM_FIRST + 15)
  LVM_GETITEMPOSITION = (LVM_FIRST + 16)
 
#If UNICODE Then
  LVM_GETSTRINGWIDTH = (LVM_FIRST + 87)
#Else
  LVM_GETSTRINGWIDTH = (LVM_FIRST + 17)
#End If
 
  LVM_HITTEST = (LVM_FIRST + 18)
  LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
  LVM_SCROLL = (LVM_FIRST + 20)
  LVM_REDRAWITEMS = (LVM_FIRST + 21)
  LVM_ARRANGE = (LVM_FIRST + 22)
  
#If UNICODE Then
  LVM_EDITLABEL = (LVM_FIRST + 118)
#Else
  LVM_EDITLABEL = (LVM_FIRST + 23)
#End If
 
  LVM_GETEDITCONTROL = (LVM_FIRST + 24)
 
#If UNICODE Then
  LVM_GETCOLUMN = (LVM_FIRST + 95)
  LVM_SETCOLUMN = (LVM_FIRST + 96)
  LVM_INSERTCOLUMN = (LVM_FIRST + 97)
#Else
  LVM_GETCOLUMN = (LVM_FIRST + 25)
  LVM_SETCOLUMN = (LVM_FIRST + 26)
  LVM_INSERTCOLUMN = (LVM_FIRST + 27)
#End If
 
  LVM_DELETECOLUMN = (LVM_FIRST + 28)
  LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
 
  LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
 
#If (WIN32_IE >= &H300) Then
  LVM_GETHEADER = (LVM_FIRST + 31)
#End If
  
  LVM_CREATEDRAGIMAGE = (LVM_FIRST + 33)
  LVM_GETVIEWRECT = (LVM_FIRST + 34)
  LVM_GETTEXTCOLOR = (LVM_FIRST + 35)
  LVM_SETTEXTCOLOR = (LVM_FIRST + 36)
  LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
  LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
  LVM_GETTOPINDEX = (LVM_FIRST + 39)
  LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
  LVM_GETORIGIN = (LVM_FIRST + 41)
  LVM_UPDATE = (LVM_FIRST + 42)
  LVM_SETITEMSTATE = (LVM_FIRST + 43)
  LVM_GETITEMSTATE = (LVM_FIRST + 44)
  
#If UNICODE Then
  LVM_GETITEMTEXT = (LVM_FIRST + 115)
  LVM_SETITEMTEXT = (LVM_FIRST + 116)
#Else
  LVM_GETITEMTEXT = (LVM_FIRST + 45)
  LVM_SETITEMTEXT = (LVM_FIRST + 46)
#End If
 
  LVM_SETITEMCOUNT = (LVM_FIRST + 47)
 
  LVM_SORTITEMS = (LVM_FIRST + 48)
  LVM_SETITEMPOSITION32 = (LVM_FIRST + 49)
  LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
  LVM_GETITEMSPACING = (LVM_FIRST + 51)
  
#If UNICODE Then
  LVM_GETISEARCHSTRING = (LVM_FIRST + 117)
#Else
  LVM_GETISEARCHSTRING = (LVM_FIRST + 52)
#End If
 
#If (WIN32_IE >= &H300) Then
  LVM_SETICONSPACING = (LVM_FIRST + 53)
  LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
  LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
 
  LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
  LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
  LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)
  LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
  LVM_SETHOTITEM = (LVM_FIRST + 60)
  LVM_GETHOTITEM = (LVM_FIRST + 61)
  LVM_SETHOTCURSOR = (LVM_FIRST + 62)
  LVM_GETHOTCURSOR = (LVM_FIRST + 63)
  LVM_APPROXIMATEVIEWRECT = (LVM_FIRST + 64)
#End If  ' // WIN32_IE >= &H300
  
#If (WIN32_IE >= &H400) Then
  LVM_SETWORKAREAS = (LVM_FIRST + 65)
  LVM_GETWORKAREAS = (LVM_FIRST + 70)
  LVM_GETNUMBEROFWORKAREAS = (LVM_FIRST + 73)
  LVM_GETSELECTIONMARK = (LVM_FIRST + 66)
  LVM_SETSELECTIONMARK = (LVM_FIRST + 67)
  LVM_SETHOVERTIME = (LVM_FIRST + 71)
  LVM_GETHOVERTIME = (LVM_FIRST + 72)
  LVM_SETTOOLTIPS = (LVM_FIRST + 74)
  LVM_GETTOOLTIPS = (LVM_FIRST + 78)

#If UNICODE Then
  LVM_SETBKIMAGE = (LVM_FIRST + 138)
  LVM_GETBKIMAGE = (LVM_FIRST + 139)
#Else
  LVM_SETBKIMAGE = (LVM_FIRST + 68)
  LVM_GETBKIMAGE = (LVM_FIRST + 69)
#End If

#End If  ' // WIN32_IE >= &H400

#If (WIN32_IE >= &H400) Then
  LVM_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT
  LVM_GETUNICODEFORMAT = CCM_GETUNICODEFORMAT
#End If

End Enum   ' LVMessages

Public Const LV_MAX_WORKAREAS = 16

' LVM_GETIMAGELIST wParam
Public Enum LVSIL_Flags
  LVSIL_NORMAL = 0
  LVSIL_SMALL = 1
  LVSIL_STATE = 2
End Enum
  
' LVM_GETNEXTITEM flags
Public Enum LVNI_Flags
  LVNI_ALL = &H0
  LVNI_FOCUSED = &H1
  LVNI_SELECTED = &H2
  LVNI_CUT = &H4
  LVNI_DROPHILITED = &H8
 
  LVNI_ABOVE = &H100
  LVNI_BELOW = &H200
  LVNI_TOLEFT = &H400
  LVNI_TORIGHT = &H800
End Enum
 
' LVM_GETITEMRECT rct.Left
Public Enum LVIR_Flags
  LVIR_BOUNDS = 0
  LVIR_ICON = 1
  LVIR_LABEL = 2
  LVIR_SELECTBOUNDS = 3
End Enum
 
' LVM_ARRANGE wParam
Public Enum LVA_Flags
  LVA_DEFAULT = &H0
  LVA_ALIGNLEFT = &H1
  LVA_ALIGNTOP = &H2
  LVA_SNAPTOGRID = &H5
End Enum
 
 
' ============================================
' Structures and their flags
 
Public Type LVITEM   ' was LV_ITEM
  mask As LVITEM_mask
  iItem As Long
  iSubItem As Long
  state As LVITEM_state
  stateMask As Long
  'pszText As Long  ' if String, must be pre-allocated
  pszText As String
  cchTextMax As Long
  iImage As Long
  lParam As Long
#If (WIN32_IE >= &H300) Then
  iIndent As Long
#End If
End Type
 
#If (WIN32_IE >= &H300) Then
Public Const I_INDENTCALLBACK = (-1)   ' iIndent, 4.70
#End If

Public Enum LVITEM_mask
  LVIF_TEXT = &H1
  LVIF_IMAGE = &H2
  LVIF_PARAM = &H4
  LVIF_STATE = &H8
#If (WIN32_IE >= &H300) Then
  LVIF_INDENT = &H10
  LVIF_NORECOMPUTE = &H800
#End If
  LVIF_DI_SETITEM = &H1000   ' NMLVDISPINFO notification
End Enum
 
Public Enum LVITEM_state
  LVIS_FOCUSED = &H1
  LVIS_SELECTED = &H2
  LVIS_CUT = &H4
  LVIS_DROPHILITED = &H8
  LVIS_ACTIVATING = &H20
 
  LVIS_OVERLAYMASK = &HF00
  LVIS_STATEIMAGEMASK = &HF000
End Enum
  
Public Type LVFINDINFO   ' was LV_FINDINFO
  flags As LVFINDINFO_flags
  psz As Long  ' if String, must be pre-allocated
  lParam As Long
  pt As POINTAPI
  vkDirection As Long
End Type
 
Public Enum LVFINDINFO_flags
  LVFI_PARAM = &H1
  LVFI_STRING = &H2
  LVFI_PARTIAL = &H8
  LVFI_WRAP = &H20
  LVFI_NEARESTXY = &H40
End Enum
 
Public Type LVHITTESTINFO   ' was LV_HITTESTINFO
  pt As POINTAPI
  flags As LVHITTESTINFO_flags
  iItem As Long
#If (WIN32_IE >= &H300) Then
  iSubItem As Long    ' this is was NOT in win95.  valid only for LVM_SUBITEMHITTEST
#End If
End Type
 
Public Enum LVHITTESTINFO_flags
  LVHT_NOWHERE = &H1   ' in LV client area, but not over item
  LVHT_ONITEMICON = &H2
  LVHT_ONITEMLABEL = &H4
  LVHT_ONITEMSTATEICON = &H8
  LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)
 
  ' outside the LV's client area
  LVHT_ABOVE = &H8
  LVHT_BELOW = &H10
  LVHT_TORIGHT = &H20
  LVHT_TOLEFT = &H40
End Enum

Public Type LVCOLUMN   ' was LV_COLUMN
  mask As LVCOLUMN_mask
  fmt As LVCOLUMN_fmt
  cx As Long
  pszText As Long  ' if String, must be pre-allocated
  cchTextMax As Long
  iSubItem As Long
#If (WIN32_IE >= &H300) Then
  iImage As Long
  iOrder As Long
#End If
End Type
 
Public Enum LVCOLUMN_mask
  LVCF_FMT = &H1
  LVCF_WIDTH = &H2
  LVCF_TEXT = &H4
  LVCF_SUBITEM = &H8
#If (WIN32_IE >= &H300) Then
  LVCF_IMAGE = &H10
  LVCF_ORDER = &H20
#End If
End Enum
 
Public Enum LVCOLUMN_fmt
  LVCFMT_LEFT = &H0
  LVCFMT_RIGHT = &H1
  LVCFMT_CENTER = &H2
  LVCFMT_JUSTIFYMASK = &H3
#If (WIN32_IE >= &H300) Then
  LVCFMT_IMAGE = &H800
  LVCFMT_BITMAP_ON_RIGHT = &H1000
  LVCFMT_COL_HAS_IMAGES = &H8000&
#End If
End Enum

Public Enum LVM_SETCOLUMNWIDTH_lParam
  LVSCW_AUTOSIZE = -1
  LVSCW_AUTOSIZE_USEHEADER = -2
End Enum
 
#If (WIN32_IE >= &H300) Then
' // these flags only apply to LVS_OWNERDATA listviews in report or list mode
Public Enum LVM_SETITEMCOUNT_lParam
  LVSICF_NOINVALIDATEALL = &H1
  LVSICF_NOSCROLL = &H2
End Enum
#End If
 
#If (WIN32_IE >= &H300) Then
Public Enum LVM_SETEXTENDEDLISTVIEWSTYLE_lParam
  LVS_EX_GRIDLINES = &H1
  LVS_EX_SUBITEMIMAGES = &H2
  LVS_EX_CHECKBOXES = &H4
  LVS_EX_TRACKSELECT = &H8
  LVS_EX_HEADERDRAGDROP = &H10
  LVS_EX_FULLROWSELECT = &H20         ' // applies to report mode only
  LVS_EX_ONECLICKACTIVATE = &H40
  LVS_EX_TWOCLICKACTIVATE = &H80
#If (WIN32_IE >= &H400) Then
  LVS_EX_FLATSB = &H100
  LVS_EX_REGIONAL = &H200
  LVS_EX_INFOTIP = &H400              ' listview does InfoTips for you
  LVS_EX_UNDERLINEHOT = &H800
  LVS_EX_UNDERLINECOLD = &H1000
  LVS_EX_MULTIWORKAREAS = &H2000
#End If  ' // WIN32_IE >= &H400
End Enum
#End If  ' // WIN32_IE >= &H300

#If (WIN32_IE >= &H400) Then
Public Type LVBKIMAGE
  ulFlags As LVBKIMAGE_ulFlags
  hbm As Long
  pszImage As Long  ' if String, must be pre-allocated
  cchImageMax As Long
  xOffsetPercent As Long
  yOffsetPercent As Long
End Type

Public Enum LVBKIMAGE_ulFlags
  LVBKIF_SOURCE_NONE = &H0
  LVBKIF_SOURCE_HBITMAP = &H1
  LVBKIF_SOURCE_URL = &H2
  LVBKIF_SOURCE_MASK = &H3
  LVBKIF_STYLE_NORMAL = &H0
  LVBKIF_STYLE_TILE = &H10
  LVBKIF_STYLE_MASK = &H10
End Enum
#End If  ' // WIN32_IE >= &H400

' ============================================
' Notifications

Public Enum LVNotifications
  LVN_FIRST = -100&   ' &HFFFFFF9C   ' (0U-100U)
  LVN_LAST = -199&   ' &HFFFFFF39   ' (0U-199U)
                                                                          ' lParam points to:
  LVN_ITEMCHANGING = (LVN_FIRST - 0)            ' NMLISTVIEW, ?, rtn T/F
  LVN_ITEMCHANGED = (LVN_FIRST - 1)             ' NMLISTVIEW, ?
  LVN_INSERTITEM = (LVN_FIRST - 2)                  ' NMLISTVIEW, iItem
  LVN_DELETEITEM = (LVN_FIRST - 3)                 ' NMLISTVIEW, iItem
  LVN_DELETEALLITEMS = (LVN_FIRST - 4)         ' NMLISTVIEW, iItem = -1, rtn T/F

  LVN_COLUMNCLICK = (LVN_FIRST - 8)              ' NMLISTVIEW, iItem = -1, iSubItem = column
  LVN_BEGINDRAG = (LVN_FIRST - 9)                  ' NMLISTVIEW, iItem
  LVN_BEGINRDRAG = (LVN_FIRST - 11)              ' NMLISTVIEW, iItem

#If (WIN32_IE >= &H300) Then
  LVN_ODCACHEHINT = (LVN_FIRST - 13)           ' NMLVCACHEHINT
  LVN_ITEMACTIVATE = (LVN_FIRST - 14)           ' v4.70 = NMHDR, v4.71 = NMITEMACTIVATE
  LVN_ODSTATECHANGED = (LVN_FIRST - 15)  ' NMLVODSTATECHANGE, rtn T/F
#End If  ' // WIN32_IE >= &H300

#If (WIN32_IE >= &H400) Then
  LVN_HOTTRACK = (LVN_FIRST - 21)                 ' NMLISTVIEW, see docs, rtn T/F
#End If
 
#If UNICODE Then
  LVN_BEGINLABELEDIT = (LVN_FIRST - 75)
  LVN_ENDLABELEDIT = (LVN_FIRST - 76)
  LVN_GETDISPINFO = (LVN_FIRST - 77)
  LVN_SETDISPINFO = (LVN_FIRST - 78)

#If (WIN32_IE >= &H300) Then
  LVN_ODFINDITEM = (LVN_FIRST - 79)             ' NMLVFINDITEM
#End If   ' (WIN32_IE >= &H300)

#If (WIN32_IE >= &H400) Then
  LVN_GETINFOTIP = (LVN_FIRST - 58)              ' NMLVGETINFOTIP
#End If  ' (WIN32_IE >= &H400)

#Else
  LVN_BEGINLABELEDIT = (LVN_FIRST - 5)        ' NMLVDISPINFO, iItem, rtn T/F
  LVN_ENDLABELEDIT = (LVN_FIRST - 6)           ' NMLVDISPINFO, see docs
  LVN_GETDISPINFO = (LVN_FIRST - 50)            ' NMLVDISPINFO, see docs
  LVN_SETDISPINFO = (LVN_FIRST - 51)            ' NMLVDISPINFO, see docs

#If (WIN32_IE >= &H300) Then
  LVN_ODFINDITEM = (LVN_FIRST - 52)             ' NMLVFINDITEM
#End If   ' (WIN32_IE >= &H300)

#If (WIN32_IE >= &H400) Then
  LVN_GETINFOTIP = (LVN_FIRST - 57)             ' NMLVGETINFOTIP
#End If  ' (WIN32_IE >= &H400)

#End If   ' UNICODE
 
  LVN_KEYDOWN = (LVN_FIRST - 55)                 ' NMLVKEYDOWN

#If (WIN32_IE >= &H300) Then
  LVN_MARQUEEBEGIN = (LVN_FIRST - 56)       ' NMLISTVIEW, rtn T/F
#End If

End Enum   ' LVNotifications

Public Type NMLISTVIEW   ' was NM_LISTVIEW
  hdr As NMHDR
  iItem As Long
  iSubItem As Long
  uNewState As LVITEM_state
  uOldState As LVITEM_state
  uChanged As LVITEM_mask
  ptAction As POINTAPI
  lParam As Long
End Type

#If (WIN32_IE >= &H400) Then

'// NMITEMACTIVATE is used instead of NMLISTVIEW in IE >= 0x400
'// therefore all the fields are the same except for extra uKeyFlags
'// they are used to store key flags at the time of the single click with
'// delayed activation - because by the time the timer goes off a user may
'// not hold the keys (shift, ctrl) any more
Public Type NMITEMACTIVATE
  hdr As NMHDR
  iItem As Long
  iSubItem As Long
  uNewState As Long
  uOldState As Long
  uChanged As Long
  ptAction As POINTAPI
  lParam As Long
  uKeyFlags As Long
End Type

'// key flags stored in uKeyFlags
Public Enum NMITEMACTIVATE_uKeyFlags
  LVKF_ALT = &H1
  LVKF_CONTROL = &H2
  LVKF_SHIFT = &H4
End Enum

#End If  ' // WIN32_IE >= &H400

#If (WIN32_IE >= &H300) Then

'Public Type NMLVCUSTOMDRAW
'  nmcd As NMCUSTOMDRAW
'  clrText As Long
'  clrTextBk As Long
'#If (WIN32_IE >= &H400) Then
'  iSubItem As Long
'#End If     ' //  WIN32_IE >= &H400
'End Type

Public Type NMLVCACHEHINT
  hdr As NMHDR
  iFrom As Long
  iTo As Long
End Type

Public Type NMLVFINDITEM   ' was NM_LVFINDITEM
  hdr As NMHDR
  iStart As Long
  lvfi As LVFINDINFO
End Type

Public Type NMODSTATECHANGE   ' was NM_ODSTATECHANGE
  hdr As NMHDR
  iFrom As Long
  iTo As Long
  uNewState As Long
  uOldState As Long
End Type

#End If   ' // WIN32_IE >= &H300

Public Type NMLVDISPINFO   ' was LV_DISPINFO
  hdr As NMHDR
  item As LVITEM
End Type

Public Type NMLVKEYDOWN   ' was LV_KEYDOWN
  hdr As NMHDR
  wVKey As Integer   ' can't be KeyCodeConstants, enums are Longs!
  flags As Long   ' Always zero.
End Type

#If (WIN32_IE >= &H400) Then

Public Type NMLVGETINFOTIP
  hdr As NMHDR
  dwFlags As Long
  pszText As Long  ' if String, must be pre-allocated
  cchTextMax As Long
  iItem As Long
  iSubItem As Long
  lParam As Long
End Type

' // NMLVGETINFOTIPA.dwFlag values ("A" ?)
Public Const LVGIT_UNFOLDED = &H1

#End If   '  // WIN32_IE >= &H400
