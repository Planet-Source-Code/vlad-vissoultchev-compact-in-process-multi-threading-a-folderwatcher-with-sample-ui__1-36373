VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cThread"
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
'  cThread is a Resource Management class as described on the original
'    site: http://www.relisoft.com/resource/resmain.html. In C++ this
'    would usually be instantiated on the stack also in local scope
'    as to enforce the mechanics of the language to cleanup upon exit.
'    Unfortuanately in VB all the object are allocated on the heap
'    so we end up with such a performance penalty that it's not
'    justified to produce such a bulk of small classes.
'
'  This one is part of the VB implementation purely for the sake of
'    being as close to the original as possible.
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "cThread"

'=========================================================================
' API
'=========================================================================

Private Const WAIT_OBJECT_0     As Long = 0

Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_handle        As Long
Private m_tid           As Long

'=========================================================================
' Methods
'=========================================================================

Public Sub Create(ByVal pFun As Long, ByVal pArg As Long)
    m_handle = CreateThread(0, 0, pFun, pArg, CREATE_SUSPENDED, m_tid)
End Sub

Public Sub Kill()
    Call TerminateThread(m_handle, 0)
End Sub

Public Sub DoResume()
    Call ResumeThread(m_handle)
End Sub

Public Function WaitForDeath() As Boolean
    WaitForDeath = (WaitForSingleObject(m_handle, 2000) = WAIT_OBJECT_0)
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    Call CloseHandle(m_handle)
    Debug.Print LIB_NAME & "." & MODULE_NAME & ".Class_Terminate"
End Sub


