VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Code Collector"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8670
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   8670
   Begin MSComctlLib.ImageList imgl 
      Left            =   3840
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06EA
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C84
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":121E
            Key             =   "selected"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3840
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "XML Files (*.xml)|*.xml"
   End
   Begin RichTextLib.RichTextBox txtXML 
      Height          =   5535
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9763
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":17B8
   End
   Begin MSComctlLib.TreeView tvXML 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9763
      _Version        =   393217
      HideSelection   =   0   'False
      PathSeparator   =   "/"
      Style           =   7
      ImageList       =   "imgl"
      BorderStyle     =   1
      Appearance      =   0
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
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsFont 
         Caption         =   "&Font..."
      End
   End
   Begin VB.Menu mnuNode 
      Caption         =   "&Node"
      Visible         =   0   'False
      Begin VB.Menu mnuNodeDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuNodeAdd 
         Caption         =   "&Add Child"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuTileHori 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuTileVert 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileName As String
Dim bFileDirty As Boolean, bInFocus As Boolean, bDirty As Boolean
Dim doc As DOMDocument40
Dim xNode As node
Dim inDrag As Boolean

Private Sub initDoc()
    'init doc object. Check out the docs
    Set doc = New DOMDocument40
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False
End Sub

'called when xml file opened
'to fill the tree with doc object and its child nodes
Private Sub AddAllNodes(iNode As DOMDocument40)
Dim iNodes As IXMLDOMNodeList
Dim i As Integer

    'this(document object) is the parent of root node, so just recurse (dont add)
    Set iNodes = iNode.childNodes
    'in case if there are more than 1 root node (not supported now)
    For i = 0 To iNodes.length - 1
        AddNodes iNodes(i)
    Next
    'Select the first node if present
    If tvXML.Nodes.Count > 0 Then tvXML_NodeClick tvXML.Nodes(1)
End Sub

'recursively adds child nodes to the tree
Private Sub AddNodes(iNode As IXMLDOMNode, Optional rel As node)
Dim iNodes As IXMLDOMNodeList
Dim i As Integer
Dim thisnode As node

    'Debug.Print iNode.nodeName & ": " & iNode.nodeType
    If iNode.nodeType <> 1 Then Exit Sub 'only of type element is accepted!

    If rel Is Nothing Then 'root node
        Set thisnode = tvXML.Nodes.Add(, tvwChild, , iNode.nodeName, , "selected")
    Else 'child of rel
        Set thisnode = tvXML.Nodes.Add(rel, tvwChild, , iNode.nodeName, , "selected")
    End If
    'set treenode images
    thisnode.Image = "closed"
    thisnode.ExpandedImage = "open"

    'recurse
    Set iNodes = iNode.childNodes
    For i = 0 To iNodes.length - 1
        AddNodes iNodes(i), thisnode
    Next
End Sub

Private Sub Form_Load()
    'load the font settings
    With txtXML.Font
        .Name = GetSetting(App.Title, "Font", "Name", "Arial")
        .Size = GetSetting(App.Title, "Font", "Size", 9)
        .Bold = GetSetting(App.Title, "Font", "Bold", False)
        .Italic = GetSetting(App.Title, "Font", "Italic", False)
        .Strikethrough = GetSetting(App.Title, "Font", "Strikethrough", False)
        .Underline = GetSetting(App.Title, "Font", "Underline", False)
    End With
    'initialize
    Call initDoc
    'if file is not opened then create new file
    If FileName = "" Then
        FileName = "Untitled.xml"
    Else
        'open the file
        doc.Load FileName
        AddAllNodes doc
    End If
    SetFileClean
End Sub

'just set the filename to the received file
'actual opening will be done in form load
Public Sub OpenXML(s As String)
    FileName = s
End Sub

'change the caption according to file's cleaness
Sub SetFileClean()
    bFileDirty = False
    Caption = FileName
End Sub

Sub SetFileDirty()
    bFileDirty = True
    Caption = FileName & "*"
End Sub

'the usual yes/no/cancel
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim ret As VbMsgBoxResult

    txtXML_LostFocus 'check if node text changed
    If bFileDirty Then
        ret = MsgBox("Save '" & FileName & "'", vbYesNoCancel, App.Title)
        If ret = vbYes Then
            Cancel = Not SaveFile 'if user clicked cancel in SaveAs Box
        ElseIf ret = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

'resize the controls
Private Sub Form_Resize()
    If ScaleHeight < 240 Then Exit Sub 'escape error
    'both treeview and txtbox take half of the form
    tvXML.Move 120, 120, ScaleWidth / 2 - 240, ScaleHeight - 240
    txtXML.Move ScaleWidth / 2, 120, ScaleWidth / 2 - 120, ScaleHeight - 240
End Sub

'clean up
Private Sub Form_Unload(Cancel As Integer)
    Set doc = Nothing
    'save the font settings
    With txtXML.Font
        SaveSetting App.Title, "Font", "Name", .Name
        SaveSetting App.Title, "Font", "Size", .Size
        SaveSetting App.Title, "Font", "Bold", .Bold
        SaveSetting App.Title, "Font", "Italic", .Italic
        SaveSetting App.Title, "Font", "Strikethrough", .Strikethrough
        SaveSetting App.Title, "Font", "Underline", .Underline
    End With
End Sub

Private Sub mnuFileClose_Click()
    Unload Me
End Sub

'transfer these to mdi
Private Sub mnuFileExit_Click()
    mdiMain.mnuFileExit_Click
End Sub

Private Sub mnuFileNew_Click()
    mdiMain.NewFile
End Sub

Private Sub mnuFileOpen_Click()
    mdiMain.OpenFile
End Sub

Function SaveFile() As Boolean
    If FileName = "Untitled.xml" Then
        'if untitled then save as
        SaveFile = SaveFileAs
    Else
        'save the file
        doc.save FileName
        SetFileClean
        SaveFile = True
    End If
End Function

Private Sub mnuFileSave_Click()
    SaveFile
End Sub

Function SaveFileAs() As Boolean
On Error GoTo CancelErr
    'get a new file location
    cd.Flags = cdlOFNOverwritePrompt
    cd.ShowSave
    FileName = cd.FileName
    'save the file with new name
    SaveFile
    SaveFileAs = True
CancelErr:
End Function

Private Sub mnuFileSaveAs_Click()
    SaveFileAs
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub

'Add new child node to the selected node
Private Sub mnuNodeAdd_Click()
Dim newNode As IXMLDOMNode, iNode As IXMLDOMNode
Dim tvNode As node, nNode As node
Dim nName As String

    On Error Resume Next
    Set tvNode = tvXML.SelectedItem
    Set iNode = doc.selectSingleNode(tvNode.FullPath) 'nodesCol(tvNode.Tag) 'the selected xmlnode

GetName:
    nName = InputBox$("Enter Valid name for new Node", , nName)
    If nName = "" Then Exit Sub 'user may have clicked cancel
    nName = Replace(nName, " ", "_") 'replace space with _ character
    
    Set newNode = doc.createElement(nName)
    If newNode Is Nothing Then 'This will trap an invalid name for new element
        MsgBox Err.Description
        GoTo GetName
    End If
    On Error GoTo 0

    'now add new node to XMLDOM, TreeView and NodesCollection
    If iNode Is Nothing Then 'root element
        doc.appendChild newNode
    Else
        iNode.appendChild newNode
    End If
    If tvNode Is Nothing Then
        Set nNode = tvXML.Nodes.Add(, , , nName, , "selected")
    Else
        Set nNode = tvXML.Nodes.Add(tvNode, tvwChild, , nName, , "selected")
    End If
    
    nNode.Image = "closed"
    nNode.ExpandedImage = "open"

    'select and click the newly created node
    nNode.Selected = True
    tvXML_NodeClick nNode

    txtXML.Locked = False 'the text in selected node can now be edited
    SetFileDirty
End Sub

'delete the selected node
Private Sub mnuNodeDelete_Click()
Dim tvNode As node
Dim iNode As IXMLDOMNode

    Set tvNode = tvXML.SelectedItem
    If tvNode Is Nothing Then Exit Sub 'routine check
    
    'Exit if unintentional
    If MsgBox("Are u sure u want to delete '" & tvNode.Text & "' and all its sub Nodes?", vbYesNo) = vbNo Then Exit Sub
    
    'get xml node from nodes collection
    Set iNode = doc.selectSingleNode(tvNode.FullPath)
    'detach it from its parent
    iNode.parentNode.removeChild iNode
    'dont remove the xml node from nedescol or all indexes will fail (tvNode.Tag)
    'x (nodescol.Remove tvNode.Tag)
    '
    'remove it from the tree view
    tvXML.Nodes.Remove tvNode.Index
    'pseudo click on the newly selected node
    Set tvNode = tvXML.SelectedItem
    If tvNode Is Nothing Then 'if the root node was deleted
        txtXML.Locked = True 'no selected node so text not editable
        txtXML.Text = ""
    Else
        tvXML_NodeClick tvNode
    End If
    SetFileDirty
End Sub

'the usual MDI stuff
Private Sub mnuTileHori_Click()
    mdiMain.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVert_Click()
    mdiMain.Arrange vbTileVertical
End Sub

Private Sub mnuToolsFont_Click()
    On Error GoTo ErrCheck
    With txtXML.Font
        cd.FontName = .Name
        cd.FontSize = .Size
        cd.FontBold = .Bold
        cd.FontItalic = .Italic
        cd.FontStrikethru = .Strikethrough
        cd.FontUnderline = .Underline
        cd.Flags = cdlCFScreenFonts
        cd.ShowFont
        .Name = cd.FontName
        .Size = cd.FontSize
        .Bold = cd.FontBold
        .Italic = cd.FontItalic
        .Strikethrough = cd.FontStrikethru
        .Underline = cd.FontUnderline
    End With
ErrCheck:
End Sub

Private Sub mnuWIcons_Click()
    mdiMain.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    mdiMain.Arrange vbCascade
End Sub

Private Sub tvXML_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim nXMLNode As IXMLDOMNode, oXMLNode As IXMLDOMNode
Dim oNode As node
Dim xmlNode As IXMLDOMNode

    'Create a xmlnode element with the new name
    On Error Resume Next
    NewString = Replace(NewString, " ", "_") 'replace space with _ character
    Set nXMLNode = doc.createElement(NewString)
    If nXMLNode Is Nothing Then
        MsgBox Err.Description
        Cancel = True
        Exit Sub
    End If
    On Error GoTo 0
    'copy the original node's attributes and child nodes
    Set oNode = tvXML.SelectedItem
    Set oXMLNode = doc.selectSingleNode(oNode.FullPath)
    For Each xmlNode In oXMLNode.Attributes
        'item have to be removed first before inserting them in new node
        oXMLNode.Attributes.removeNamedItem xmlNode.nodeName
        nXMLNode.Attributes.setNamedItem xmlNode
    Next
    For Each xmlNode In oXMLNode.childNodes
        nXMLNode.appendChild xmlNode
    Next
    'replace the new node with the orig node
    oXMLNode.parentNode.replaceChild nXMLNode, oXMLNode
    
    SetFileDirty
End Sub

Private Sub tvXML_DragDrop(Source As Control, x As Single, y As Single)
Dim iNode As IXMLDOMNode 'selected node
Dim iNodeT As IXMLDOMNode 'dropped to node
Dim pNode As IXMLDOMNode 'selected node's parent

   'On Error GoTo ErrCheck
   If tvXML.DropHighlight Is Nothing Then
      inDrag = False
      Exit Sub
   Else
      If xNode = tvXML.DropHighlight Then Exit Sub
      'Debug.Print xNode.Text & " dropped on " & tvXML.DropHighlight.Text
      'now move the node
      
        'get all info from treeview before changing it
        Set iNode = doc.selectSingleNode(xNode.FullPath)
        Set iNodeT = doc.selectSingleNode(tvXML.DropHighlight.FullPath)
        Set pNode = iNode.parentNode

        'add all treeview nodes(draggedto and its child) in new location
        AddNodes doc.selectSingleNode(xNode.FullPath), tvXML.DropHighlight
        'remove node from treeview
        tvXML.Nodes.Remove xNode.Index
        'move the node in xmldoc
        iNode.parentNode.removeChild iNode
        iNodeT.appendChild iNode
        
      inDrag = False
      SetFileDirty
   End If
   
   Exit Sub
ErrCheck:
    MsgBox Err.Description
    pNode.appendChild iNode
End Sub

Private Sub tvXML_DragOver(Source As Control, x As Single, y As Single, State As Integer)
   If inDrag = True Then
      ' Set DropHighlight to the mouse's coordinates.
      Set tvXML.DropHighlight = tvXML.HitTest(x, y)
      If Not tvXML.DropHighlight Is Nothing Then tvXML.DropHighlight.Expanded = True
   End If
End Sub

Private Sub tvXML_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set xNode = tvXML.HitTest(x, y)
End Sub

Private Sub tvXML_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If xNode Is Nothing Then Exit Sub
    If Button = vbLeftButton Then ' Signal a Drag operation.
        inDrag = True ' Set the flag to true.
        ' Set the drag icon with the CreateDragImage method.
        tvXML.DragIcon = tvXML.SelectedItem.CreateDragImage
        tvXML.Drag vbBeginDrag ' Drag operation.
    End If
End Sub

Private Sub tvXML_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'show the node(add/delete) popup menu
    If Button = 2 Then
        Me.PopupMenu mnuNode
    End If
End Sub

'Show the text related to this node in the textbox
Private Sub tvXML_NodeClick(ByVal node As MSComctlLib.node)
Dim iNode As IXMLDOMNode

    'node is clicked: get the text node associated with the xmlnode
    'show text in the text box if text node exists
    Set iNode = doc.selectSingleNode(node.FullPath)
    txtXML.Text = ""
    'I store the text as the first child node
    'the node type of text node is 3
    If Not iNode.firstChild Is Nothing Then
        If iNode.firstChild.nodeType = 3 Then
            txtXML.Text = iNode.firstChild.nodeValue
        End If
    End If
    txtXML.Locked = False
    node.Selected = True
    Set tvXML.DropHighlight = node
End Sub

Private Sub txtXML_Change()
    'text for the selected node has changed so it is dirty
    'text will also change when textbox in not in focus
    If bInFocus = True Then bDirty = True
End Sub

Private Sub txtXML_GotFocus()
    'starting to edit node text so...
    bInFocus = True
    bDirty = False
End Sub

Sub SaveSelNodeText()
Dim tvNode As node
Dim iNode As IXMLDOMNode
Dim newNode As IXMLDOMText

    'get the selected node
    Set tvNode = tvXML.SelectedItem
    If tvNode Is Nothing Then Exit Sub 'just be sure

    Set iNode = doc.selectSingleNode(tvNode.FullPath)
    'do we have a child node associated with the element node
    If iNode.firstChild Is Nothing Then
        'if no then create new textnode and append
        Set newNode = doc.createTextNode(txtXML.Text)
        iNode.appendChild newNode
    Else
        'if childnode is present check if it is a text node
        If iNode.firstChild.nodeType = 3 Then '3: text element
            iNode.firstChild.nodeValue = txtXML.Text
        Else
            'if no then add a text node as the first child node
            Set newNode = doc.createTextNode(txtXML.Text)
            iNode.insertBefore newNode, iNode.childNodes(0)
        End If
    End If
End Sub

Private Sub txtXML_LostFocus()
    bInFocus = False
    If Not bDirty Then Exit Sub
    'if node was edited then save it
    Call SaveSelNodeText
    bDirty = False
    SetFileDirty
End Sub

