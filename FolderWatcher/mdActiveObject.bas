Attribute VB_Name = "mdActiveObject"
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
'  mdActiveObject implements the ThreadEntry function. This is where
'    the magic happend. Read on, it's less than 30 lines! Basicly
'    all the lines before the On Error Resume Next are NOT supposed to
'    happen. That is the VB run-time has never been tested and does not
'    support such kind of code. This is the reason for the use of
'    CoCreateInstance to create an object (usually we use New or
'    CreateObject for such a feast). Once we get our hand on an interface
'    to a VB object we are safe (saved) to go on with normal VB coding.
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "mdActiveObject"

'=========================================================================
' Constants and member variables
'=========================================================================

Public Const LIB_NAME           As String = "FolderWatcher"

'=========================================================================
' Functions
'=========================================================================

Public Function ThreadEntry(ByVal pArg As Long) As Long
    Dim IID_IUnknown        As VBGUID
    Dim CLSID_StdPicture    As VBGUID
    Dim pActive             As cFolderWatcher
    Dim pUnk                As IUnknown
    Dim hr                  As Long
    
    '--- initialize COM libs
    Call CoInitialize(0)
    '--- create an object
    IID_IUnknown = GUIDFromString("{00000000-0000-0000-C000-000000000046}")
    CLSID_StdPicture = CLSIDFromProgID(LIB_NAME & ".cDummy")
    hr = CoCreateInstance(CLSID_StdPicture, Nothing, CLSCTX_INPROC_SERVER, IID_IUnknown, pUnk)
    Set pUnk = Nothing
    On Error Resume Next
    '--- get cFolderWatcher interface
    Call CopyMemory(pUnk, ByVal pArg, 4)
    Set pActive = pUnk
    Call CopyMemory(pUnk, 0, 4)
    '--- call cFolderWatcher methods
    Call pActive.InitThread
    Call pActive.DoLoop
    Call pActive.FlushThread
    '--- clean up
    Set pActive = Nothing
    '--- uninitialize COM libs
    Call CoUninitialize
End Function
