Attribute VB_Name = "FolderMonitorGlobals"
Option Explicit

' FindFirstChangeNotification Constants:

' Any filename change causes a change notification wait operation to
' return. Changes include, renaming creating, or deleting a filename.
Public Const FILE_NOTIFY_CHANGE_FILE_NAME As Long = &H1

' Any directory-name change causes a change notification wait operation
' to return. Changes include creating or deleting a directory.
Public Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2 'not used

' Any attribute change causes a change notification wait
' operation to return.
Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4

' Any file-size change causes a change notification wait operation to
' return. File size changes are only detected when the file is written
' to disk. For operating systems that use extensive caching, detection
' occurs only when the cache is sufficiently flushed.
Public Const FILE_NOTIFY_CHANGE_SIZE As Long = &H8 'not used

' Any change to the last write-time of files causes a change notification
' wait operation to return. The change is detected when the file is written
' to disk. For operating systems that use extensive caching, detection occurs
' only when the cache is sufficiently flushed.
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE As Long = &H10
Public Const FILE_NOTIFY_CHANGE_LAST_ACCESS As Long = &H20 'not used
Public Const FILE_NOTIFY_CHANGE_CREATION As Long = &H40 'not used

' Any security-descriptor change in the watched directory or subtree causes
' a change notification wait operation to return.
Public Const FILE_NOTIFY_CHANGE_SECURITY As Long = &H100
Public Const FILE_NOTIFY_FLAGS = FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                          FILE_NOTIFY_CHANGE_FILE_NAME Or _
                          FILE_NOTIFY_CHANGE_LAST_WRITE

Public Const INVALID_HANDLE_VALUE = -1
Public Const SYNCHRONIZE = &H100000
Public Const WM_CLOSE = &H10

' Wait Function Declarations:
Public Const WAIT_OBJECT = &H0
Public Const WAIT_TIMEOUT = &H102
Public Const WAIT_TIME = 100
Public Const WAIT_FAILED = &HFFFFFFFF
Public Const INFINITE = -1&      '  Infinite timeout
Public Const MAXIMUM_WAIT_OBJECTS = &H40

' This function returns when one of the following occurs:
' Return Value    Reason
' ------------    ---------------------------------------------
' WAIT_OBJECT     The specified object is in the signaled state
' WAIT_TIMEOUT    The time-out interval elapses.

' Change Notification Functions:

' This function creates a change notification handle and sets up initial
' change notification filter conditions. A wait on a notification handle
' succeeds when a change matching the filter conditions occurs in the
' specified directory or subtree.
Public Declare Function FindFirstChangeNotification Lib "kernel32" _
    Alias "FindFirstChangeNotificationA" _
   (ByVal lpPathName As String, _
    ByVal bWatchSubtree As Long, _
    ByVal dwNotifyFilter As Long) As Long

' This function requests that the operating system signal a change
' notification handle the next time it detects an appropriate change.
Public Declare Function FindNextChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

' This function stops change notification handle monitoring.
Public Declare Function FindCloseChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

Public Declare Function WaitForMultipleObjects Lib "kernel32" ( _
    ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll _
    As Long, ByVal dwMilliseconds As Long) As Long

