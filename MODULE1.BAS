Attribute VB_Name = "Module1"
Option Explicit

Declare Function MoveWindow Lib "user32" _
                       (ByVal hwnd As Long, _
                        ByVal x As Long, ByVal y As Long, _
                        ByVal nWidth As Long, ByVal nHeight As Long, _
                        ByVal bRepaint As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                        (ByVal hwnd As Long, ByVal wMsg As Long, _
                         ByVal wParam As Long, lParam As Any) As Long

' This message helps speed up the initialization of list boxes that have a large number
' of items (more than 100). It preallocates the specified amount of memory so that
' subsequent LB_ADDSTRING, LB_INSERTSTRING, LB_DIR, and LB_ADDFILE
' messages take the shortest possible time. You can use estimates for the wParam and
' lParam parameters. If you overestimate, some extra memory is allocated; if you
' underestimate, the normal allocation is used for items that exceed the preallocated amount.
' wParam:          Specifies the number of items to add.
' lParam:           Specifies the amount of memory, in bytes, to allocate for item strings.
' Return Value:   The return value is the maximum number of items that the memory
                       ' object can store before another memory reallocation is needed, if
                       ' successful. It is LB_ERRSPACE if not enough memory is available.
Public Const LB_INITSTORAGE = &H1A8

' An application sends an LB_ADDSTRING message to add a string to a list box.
' If the list box does not have the LBS_SORT style, the string is added to the end
' of the list. Otherwise, the string is inserted into the list and the list is sorted.
Public Const LB_ADDSTRING = &H180

Public Const WM_SETREDRAW = &HB
Public Const WM_VSCROLL = &H115
Public Const SB_BOTTOM = 7

' If the function succeeds, the return value is a bitmask
' representing the currently available disk drives. Bit
' position 0 (the least-significant bit) is drive A, bit position
' 1 is drive B, bit position 2 is drive C, and so on.
' If the function fails, the return value is zero.
Declare Function GetLogicalDrives Lib "kernel32" () As Long

' If the function succeeds, the return value is a search handle
' used in a subsequent call to FindNextFile or FindClose
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
                        (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

'FindFirstFile failure rtn value
Public Const INVALID_HANDLE_VALUE = -1

' Rtns True (non zero) on succes, False on failure
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
                        (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

' Rtns True (non zero) on succes, False on failure
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Const MaxLFNPath = 260

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MaxLFNPath
        cShortFileName As String * 14
End Type
