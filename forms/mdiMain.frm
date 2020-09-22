VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Code Collector"
   ClientHeight    =   6105
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9990
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   4800
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "XML Files (*.xml)|*.xml"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New XML"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open XML..."
         Shortcut        =   ^O
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
Dim sStr As String
Dim doc As New frmMain

    Me.Top = GetSetting(App.Title, "MainWindow", "Top", 0)
    Me.Left = GetSetting(App.Title, "MainWindow", "Left", 0)
    Me.Width = GetSetting(App.Title, "MainWindow", "Width", 600)
    Me.Height = GetSetting(App.Title, "MainWindow", "Height", 400)
    cd.Flags = cdlOFNFileMustExist
    
    'get commandline parameters. it should only be a file name
    sStr = Command$
    'trim quote marks
    If Left$(sStr, 1) = Chr$(34) Then sStr = Right$(sStr, Len(sStr) - 1)
    If Right$(sStr, 1) = Chr$(34) Then sStr = Left$(sStr, Len(sStr) - 1)

    If sStr <> "" And Dir(sStr) <> "" Then
        doc.OpenXML sStr
    End If
    doc.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    SaveSetting App.Title, "MainWindow", "Top", Me.Top
    SaveSetting App.Title, "MainWindow", "Left", Me.Left
    SaveSetting App.Title, "MainWindow", "Width", Me.Width
    SaveSetting App.Title, "MainWindow", "Height", Me.Height
End Sub

Public Sub mnuFileExit_Click()
    Unload Me
    'End
End Sub

Private Sub mnuFileNew_Click()
    NewFile
End Sub

Public Sub NewFile()
Dim doc As New frmMain
    doc.Show
End Sub

Private Sub mnuFileOpen_Click()
    OpenFile
End Sub

Public Sub OpenFile()
Dim doc As New frmMain

    On Error GoTo ErrExit
    cd.ShowOpen
    doc.OpenXML cd.FileName
    doc.Show
    
ErrExit:
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub
