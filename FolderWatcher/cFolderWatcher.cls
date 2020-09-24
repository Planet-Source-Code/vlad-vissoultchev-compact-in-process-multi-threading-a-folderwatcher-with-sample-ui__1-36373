VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFolderWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'  cFolderWatcher is the only publicly visible class from the library.
'    It implements the main interface as the client sees it and also the
'    internal InitThread, DoLoop and FlushThread. Originally these were
'    implementaed as a separate interface but in the C++ this was nicely
'    inherited from ActiveObject class. It turned out ugly in VB so I had
'    to get rid of the base class and went straight for implementation.
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "cFolderWatcher"

'=========================================================================
' API
'=========================================================================

Private Const WAIT_OBJECT_0         As Long = 0
Private Const WM_KEYDOWN            As Long = &H100

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_hwndNotifySink        As Long
Private m_folder                As String
Private m_isDying               As Long
Private m_notifySource          As cFileChangeEvent
Private m_thread                As cThread
Private m_activeObject          As cFolderWatcher

'=========================================================================
' Methods
'=========================================================================

Public Sub Create(folder As String, ByVal hwnd As Long, Optional ByVal recursive As Boolean)
    m_hwndNotifySink = hwnd
    m_folder = folder
    m_isDying = 0
    '--- create notify
    Set m_notifySource = New cFileChangeEvent
    If recursive Then
        Call m_notifySource.Create(folder, True, FNC_FILE_NAME Or FNC_DIR_NAME)
    Else
        Call m_notifySource.Create(folder, False, FNC_FILE_NAME Or FNC_SIZE Or FNC_ATTRIBUTES Or FNC_LAST_WRITE Or FNC_SECURITY)
    End If
    '--- save pointer
    Set m_activeObject = Me
    '--- spin thread
    Set m_thread = New cThread
    Call m_thread.Create(AddressOf ThreadEntry, VarPtr(m_activeObject))
    Call m_thread.DoResume
End Sub

Public Sub Kill()
    If Not m_activeObject Is Nothing Then
        Set m_activeObject = Nothing
        m_isDying = m_isDying + 1
        If Not m_thread.WaitForDeath() Then
            m_thread.Kill
        End If
        Set m_thread = Nothing
    End If
End Sub

Friend Sub InitThread()
    '--- empty
End Sub

Friend Sub DoLoop()
    Do While True
        If m_isDying <> 0 Then
            Exit Sub
        End If
        If WaitForSingleObject(m_notifySource.Handle, 100) = WAIT_OBJECT_0 Then
            '--- notify sink for m_folder change
            Call SetWindowText(m_hwndNotifySink, m_folder)
            Call PostMessage(m_hwndNotifySink, WM_KEYDOWN, 32, &H10001)
            '--- capture next change
            Call m_notifySource.ContinueNotification
        End If
    Loop
End Sub

Friend Sub FlushThread()
    '--- empty
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    Call Kill
    Set m_notifySource = Nothing
    Debug.Print LIB_NAME & "." & MODULE_NAME & ".Class_Terminate"
End Sub
