Attribute VB_Name = "mdBrowse"
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
'  mdBrowse as seen in the Directory Monitor by Walter Brebels (Walter.Brebels@gmx.net)
'    on PSC here http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=36266&lngWId=1
'
'  I had no time to read it, nor to test it :-))
'
'=========================================================================

'This is code has been "borrowed" from Kamilche (www.Kamilche.com)
'and slightly modified.
'It can be found at: http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=9731&lngWId=1

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260


Private Declare Function SHBrowseForFolder Lib "shell32" _
    (lpbi As BrowseInfo) As Long


Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long


Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long


Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
    
Public Function GetFolderName(hWnd As Long, szTitle As String) As String
    'Opens a Treeview control that displays
    '     the directories in a computer
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo

    With tBrowseInfo
        .hWndOwner = hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
    
    GetFolderName = sBuffer
End Function
