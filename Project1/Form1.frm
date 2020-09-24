VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4404
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4404
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3192
      Top             =   3528
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "file"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0112
            Key             =   "folder"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAddress 
      Caption         =   "Address"
      Height          =   348
      Left            =   84
      TabIndex        =   5
      Top             =   84
      Width           =   852
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3780
      Top             =   3528
   End
   Begin VB.TextBox txtSink 
      Height          =   288
      Left            =   4284
      TabIndex        =   4
      Top             =   3612
      Visible         =   0   'False
      Width           =   1188
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   276
      Left            =   0
      TabIndex        =   3
      Top             =   4128
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   487
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10287
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   3288
      Left            =   84
      TabIndex        =   2
      Top             =   756
      Visible         =   0   'False
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   5800
      SortKey         =   2
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modified"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtFolder 
      Height          =   348
      Left            =   1008
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   84
      Width           =   4296
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   348
      Left            =   5376
      TabIndex        =   0
      Top             =   84
      Width           =   600
   End
   Begin VB.Label lblInfo 
      Caption         =   "Choose folder to monitor..."
      Height          =   264
      Left            =   84
      TabIndex        =   6
      Top             =   504
      Width           =   5892
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'  Form1, yes, this is actually meant as a name. An UI for the folder
'    watcher library. You have to compile the FolderWatcher dll first!
'    It will crash the IDE (hopelessly) if run inside it.
'
'  One drawback of the library is that it needs a handle to the window
'    of a textbox to raise the events. That is you get KeyDown on a
'    hidden textbox when rescan is needed. The contents of the textbox
'    is the folder that's been modifies. So a single textbox can be
'    used by multiple instances (objects) of the watcher.
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "Form1"
 
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long

Private Const STR_MONITORING        As String = " Monitoring..."
Private Const STR_RESCANNING        As String = " Re-scanning "

Private m_watcher           As cFolderWatcher
Private m_folder            As String

Private Function ShowError(sFunc As String) As VbMsgBoxResult
    ShowError = MsgBox(Error, vbCritical + vbAbortRetryIgnore, MODULE_NAME & "." & sFunc)
End Function

Private Sub RescanFolder()
    Const FUNC_NAME     As String = "RescanFolder"
    Dim file            As String
    Dim count           As Long
    Dim keySelected     As String
    Dim hWndFocus       As Long
    
    On Error GoTo EH
    lblInfo = STR_RESCANNING & m_folder & "..."
    lblInfo.BackColor = vbHighlight
    lblInfo.ForeColor = vbHighlightText
    UpdateWindow hWnd
    hWndFocus = GetFocus()
    If Not lvwFiles.SelectedItem Is Nothing Then
        keySelected = lvwFiles.SelectedItem.Key
        Set lvwFiles.SelectedItem = Nothing
    End If
    lvwFiles.Visible = False
    lvwFiles.ListItems.Clear
    file = Dir(m_folder & "\*.*", vbDirectory)
    Do While Len(file) > 0
        count = count + 1
        With lvwFiles.ListItems.Add(Key:="#" & file)
            .Text = file
            If Left(file, 1) <> "." Then
                .SubItems(3) = FileDateTime(m_folder & "\" & file)
                If (GetAttr(m_folder & "\" & file) And vbDirectory) <> 0 Then
                    .SmallIcon = "folder"
                Else
                    .SubItems(1) = Format(FileLen(m_folder & "\" & file), "#,##0")
                    .SubItems(2) = Right(m_folder & "\" & file, 4)
                    .SmallIcon = "file"
                End If
            Else
                .SmallIcon = "folder"
            End If
        End With
        file = Dir
    Loop
    lvwFiles.Visible = True
    On Error Resume Next
    Set lvwFiles.SelectedItem = lvwFiles.ListItems(keySelected)
    If hWndFocus = lvwFiles.hWnd Then
        lvwFiles.SetFocus
    End If
    lblInfo = STR_MONITORING
    lblInfo.BackColor = vbButtonFace
    lblInfo.ForeColor = vbWindowText
    sbMain.SimpleText = " " & count & " item(s) @ " & Time()
    Exit Sub
EH:
    lblInfo = STR_MONITORING
    lblInfo.BackColor = vbButtonFace
    lvwFiles.Visible = True
    sbMain.SimpleText = ""
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub cmdAddress_Click()
    Const FUNC_NAME     As String = "cmdAddress_Click"
    Dim folder          As String
    
    On Error GoTo EH
    folder = GetFolderName(hWnd, txtFolder)
    If Len(folder) > 0 Then
        txtFolder = folder
        cmdGo_Click
    End If
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub cmdGo_Click()
    Const FUNC_NAME     As String = "cmdGo_Click"
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    m_folder = txtFolder
    m_watcher.Kill
    m_watcher.Create m_folder, txtSink.hWnd, True
    Call RescanFolder
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Set m_watcher = New cFolderWatcher
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    m_watcher.Kill
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lvwFiles.Width = ScaleWidth - 2 * lvwFiles.Left
    lvwFiles.Height = ScaleHeight - lvwFiles.Top - sbMain.Height
    cmdGo.Left = ScaleWidth - cmdGo.Width - lvwFiles.Left
    txtFolder.Width = cmdGo.Left - lvwFiles.Left - txtFolder.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_watcher.Kill
    Set m_watcher = Nothing
End Sub

Private Sub lvwFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvwFiles.SortKey <> ColumnHeader.Index - 1 Then
        lvwFiles.SortKey = ColumnHeader.Index - 1
        lvwFiles.SortOrder = lvwAscending
    Else
        lvwFiles.SortOrder = 1 - lvwFiles.SortOrder
    End If
End Sub

Private Sub lvwFiles_DblClick()
    If Not lvwFiles.SelectedItem Is Nothing Then
        txtFolder = txtFolder & IIf(Right(txtFolder, 1) = "\", "", "\") & lvwFiles.SelectedItem.Text
        cmdGo_Click
    End If
End Sub

Private Sub Timer1_Timer()
    Screen.MousePointer = vbHourglass
    Timer1.Enabled = False
    RescanFolder
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtSink_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "txtSink_KeyDown "; KeyCode; Timer
    Timer1.Enabled = True
    lblInfo = STR_RESCANNING & m_folder & "..."
    lblInfo.BackColor = vbHighlight
    lblInfo.ForeColor = vbHighlightText
End Sub
