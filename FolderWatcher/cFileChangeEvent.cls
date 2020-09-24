VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileChangeEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
'  Folder Watcher Project
'  Copyright (c) 2002 Vlad Vissoultchev (wqw@myrealbox.com)
'
'  This code is inspired by http://www.relisoft.com/Win32/watcher.html
'
'  Just had to translate the idea to VB. It turned out UGLY compared
'    to C++ implementation but at least it is working. So this is
'    actually pretty compact example of in-process multithreading
'    in VB6.
'
'  This code has an external reference to the threadapi.tlb (API calls
'    used for threading) which was orginally supplied by the "father"
'    of the VB6 in-process multithreading hack: Mathew Curland. Btw,
'    find a copy on "Advanced Visual Basic" by the same author it's
'    worth it!
'
'=========================================================================
'
'  cFileChangeEvent had to be a base class for a number of classes that
'    would determine what to be watched for. It turned out that this is
'    not so easy to implement in VB when the language is lacking
'    implementation inheritence, so the library has only the base class
'    Thanks god it's not a abstract one in the original and the
'    constructor is parameterised sufficiently.
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "cFileChangeEvent"

'=========================================================================
' API
'=========================================================================

Public Enum FILE_NOTIFY_CHANGE
    FNC_FILE_NAME = &H1
    FNC_DIR_NAME = &H2
    FNC_ATTRIBUTES = &H4
    FNC_SIZE = &H8
    FNC_LAST_WRITE = &H10
    FNC_SECURITY = &H100
End Enum

Private Const INVALID_HANDLE_VALUE = -1

Private Declare Function FindFirstChangeNotification Lib "kernel32" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As Long
Private Declare Function FindNextChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long
Private Declare Function FindCloseChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_handle            As Long

'=========================================================================
' Methods
'=========================================================================

Public Sub Create( _
            folder As String, _
            ByVal recursive As Boolean, _
            ByVal notifyFlags As FILE_NOTIFY_CHANGE)
    m_handle = FindFirstChangeNotification(folder, Abs(recursive), notifyFlags)
    If m_handle = INVALID_HANDLE_VALUE Then
        Err.Raise vbObjectError, , "Cannot create change notification handle"
    End If
End Sub

Property Get Handle() As Long
    Handle = m_handle
End Property

Public Function ContinueNotification() As BOOL
    ContinueNotification = FindNextChangeNotification(m_handle)
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    If m_handle <> INVALID_HANDLE_VALUE Then
        Call FindCloseChangeNotification(m_handle)
    End If
    Debug.Print LIB_NAME & "." & MODULE_NAME & ".Class_Terminate"
End Sub


