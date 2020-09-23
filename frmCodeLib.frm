VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmCodeLib 
   Caption         =   "VBCodeLibrary Tool"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCodeLib.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar tbTools 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   17
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "OPENDB"
            Object.ToolTipText     =   "Open Database"
            Object.Tag             =   "OPENDB"
            ImageKey        =   "OPEN"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NEW"
            Object.ToolTipText     =   "Add New Code"
            Object.Tag             =   "NEW"
            ImageKey        =   "NEW"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "DELETE"
            Object.ToolTipText     =   "Delete Selected Code Item"
            Object.Tag             =   "DELETE"
            ImageKey        =   "DELETE"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "mnuFind"
            Object.ToolTipText     =   "FIND"
            Object.Tag             =   "FIND"
            ImageKey        =   "FIND"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PRINT"
            Object.ToolTipText     =   "Print Selected Code Window"
            Object.Tag             =   "PRINT"
            ImageKey        =   "PRINT"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "VIEWBOOKMARKS"
            Object.ToolTipText     =   "View Bookmarks"
            Object.Tag             =   "VIEWBOOKMARKS"
            ImageKey        =   "VIEWBOOKMARKS"
            Style           =   1
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BOOKMARK"
            Object.ToolTipText     =   "Add Bookmark to selected Code Item"
            Object.Tag             =   "BOOKMARK"
            ImageKey        =   "ADDBOOKMARK"
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SEP1"
            Object.Tag             =   "SEP1"
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Previous"
            Object.Tag             =   "Previous"
            ImageKey        =   "PREVIOUS"
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Next"
            Object.ToolTipText     =   "Next"
            Object.Tag             =   "Next"
            ImageKey        =   "NEXT"
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SETTINGS"
            Object.ToolTipText     =   "Settings"
            Object.Tag             =   "SETTINGS"
            ImageKey        =   "SETTINGS"
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ProgressBar prgBar 
      Height          =   270
      Left            =   6465
      TabIndex        =   7
      Top             =   4590
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   476
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmrDragTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7200
      Top             =   3885
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3180
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Left            =   3060
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.PictureBox picSysBar 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8715
      ScaleHeight     =   270
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   6225
      Visible         =   0   'False
      Width           =   375
   End
   Begin ComctlLib.TreeView tvCodeItems 
      DragIcon        =   "frmCodeLib.frx":030A
      Height          =   3255
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   5741
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   5280
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5768
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   5768
            TextSave        =   "19:36"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VBCodeLib.ctlBookmarks ctlBookMarkList 
      Height          =   1365
      Left            =   15
      TabIndex        =   3
      Top             =   3735
      Visible         =   0   'False
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   2408
   End
   Begin VBCodeLib.ctlCodeDetails ctlCodeItemDetails 
      Height          =   3000
      Left            =   3675
      TabIndex        =   2
      Top             =   555
      Width           =   4365
      _ExtentX        =   5980
      _ExtentY        =   4286
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   2625
      MousePointer    =   9  'Size W E
      Top             =   255
      Width           =   195
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3015
      Top             =   1470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":0614
            Key             =   "NEW"
            Object.Tag             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":0726
            Key             =   "CHILD"
            Object.Tag             =   "CHILD"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":0A40
            Key             =   "FOLDER"
            Object.Tag             =   "FOLDER"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":0D5A
            Key             =   "DELETE"
            Object.Tag             =   "DELETE"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":0E6C
            Key             =   "OPENFOLDER"
            Object.Tag             =   "OPENFOLDER"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":1186
            Key             =   "SETTINGS"
            Object.Tag             =   "SETTINGS"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":14A0
            Key             =   "PREVIOUS"
            Object.Tag             =   "PREVIOUS"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":17F2
            Key             =   "NEXT"
            Object.Tag             =   "NEXT"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":1B44
            Key             =   "BAS"
            Object.Tag             =   "BAS"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":1E96
            Key             =   "CLS"
            Object.Tag             =   "CLS"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":2158
            Key             =   "VB"
            Object.Tag             =   "VB"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":24AA
            Key             =   "VIEWBOOKMARKS"
            Object.Tag             =   "VIEWBOOKMARKS"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":27C4
            Key             =   "ADDBOOKMARK"
            Object.Tag             =   "ADDBOOKMARK"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":2ADE
            Key             =   "OPEN"
            Object.Tag             =   "OPEN"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":2BF0
            Key             =   "PRINT"
            Object.Tag             =   "PRINT"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCodeLib.frx":2D02
            Key             =   "FIND"
            Object.Tag             =   "FIND"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewDatabase 
         Caption         =   "&New VB CodeLibrary Database"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpenDatabase 
         Caption         =   "&Open VB Codelibrary Database"
      End
      Begin VB.Menu fsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAddCode 
         Caption         =   "&Add New Code Here"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Code from Here"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import Code here"
      End
      Begin VB.Menu mnuDeleteCode 
         Caption         =   "&Delete Selected Code Item"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename Code"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookMark 
         Caption         =   "&BookMark Here"
      End
      Begin VB.Menu mnuCount 
         Caption         =   "Coun tEm"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Application Settings"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewBookMarks 
         Caption         =   "&Bookmarks"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmCodeLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' VBCodeLib Form
'
' Written specially for the CodeGuru Site as a tool for developers
'
' http://www.codeguru.com/vb
'
' Chris Eastwood Feb.1999
'
' Updated March/April 1999 -
'
' Added Import/Export Items Functionality
' Added Import/Export Files Functionality
' Added Update Database Version Functionality
'
'

'
' Private Members
'
Private moSettings As CSettings             ' Our Settings Object
Private mnDragNode As Node                  ' Node being Dragged
Private mDB As Database                     ' Our Database Object (DAO for now)
Private miClipBoardFormat As Integer        ' Our Custom Clipboard Format
Private mbSplitting As Boolean              ' Are we splitting ?
Private mbShowBookmarks As Boolean          ' Are we showing Bookmarks at present ?
Private msDBFileName As String              ' Current Database File Name
Private miScrollDir As Integer              ' Direction the TreeView is scrolling

Private Const DEFAULTDB = "VBCodeLib.mdb"   ' Default CodeLibrary Database Name
Private Const lVSplitLimit As Long = 1500   ' Splitter side limits

Private Sub ctlBookMarkList_BookMarkRemoved(ByVal sCodeID As String)
'
' Controls is telling us that a bookmark has been removed
'
' Update the gui as necessary
'
End Sub

Private Sub ctlBookMarkList_ViewBookMark(oCodeItem As IDataObject)
    Dim nNode As Node
'
' Bookmark Control is telling us that it needs to link to the relevant item
' in the TreeView Control
'
    On Error Resume Next
    
    Set nNode = tvCodeItems.Nodes("C" & oCodeItem.Key)
    nNode.EnsureVisible
    nNode.Selected = True
    ctlCodeItemDetails.Initialise mDB, oCodeItem
    StatusBar1.Panels(1).Text = nNode.Text

End Sub

Private Sub ctlCodeItemDetails_RequestFileName(ByVal DialogType As eGetFileDialog, sFilename As String, ByVal sDialogTitle As String)
'
' CodeItemDetails control is asking us for a filename
'
    sFilename = GetFileName(DialogType, sFilename, sDialogTitle)
End Sub

Private Sub ctlCodeItemDetails_ViewChanged(ByVal CurrentView As eCurView)
'
' View has changed, update the necessary toolbar buttons
'
    tbTools.Buttons("PRINT").Enabled = (CurrentView <> vwFiles)
End Sub

Private Sub Form_Load()

    Dim oWaitCursor As CWaitCursor
    Dim oButton As Button
    
    
On Error GoTo vbErrorHandler
'
' Set Cursor to HourGlass
'
    Set oWaitCursor = New CWaitCursor
    oWaitCursor.SetCursor
    
'
' Register Our New Clipboard Format
'
    miClipBoardFormat = RegisterClipboardFormat("VBCodeLibTree")
'
' Create our new settings Object
'
    Set moSettings = New CSettings
'
' Execute Startup Procedures as defined in the Settings Object
'
    GetLastDBName
    
    DoStartUp
'
' Set the toolbar to 'flat' style, and TrackSelect on the TreeView
'
    InitControls
'
' Setup the SysTray Icon
'
    SetupSysTrayIcon
'
' Create our Link to the Database
'
    If SetupDBConnection = True Then
'
' Set ToolBar's ImageList
'
        Set tbTools.ImageList = ImageList1
'
' Fill the Tree with our code items from the DataBase
'
        FillTree
    
'
' Initialise the UserControls
'
        ctlBookMarkList.Initialise mDB
        ctlCodeItemDetails.Initialise mDB, Nothing

        ShowBookmarks mbShowBookmarks
        EnableControls True

    Else
        EnableControls False
        ShowBookmarks True
    End If
    
    Set oWaitCursor = Nothing
    
    Exit Sub

vbErrorHandler:
'
' Error handling could be nicer, but hey !
' it's only an example application.
'
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " frmCodeLib::Form_Load", , App.ProductName
    
    Set oWaitCursor = Nothing
End Sub

Private Sub FillTree()

On Error GoTo vbErrorHandler

'
' Populate our TreeView Control with the Data from our database
'
    Dim lCount As Long
    Dim rsSections As Recordset
    Dim sParent As String
    Dim sKey As String
    Dim sText As String
    Dim bBookMark As Boolean
    Dim nNode As Node
    
    Set rsSections = mDB.OpenRecordset("select * from codeitems order by parentid", dbOpenSnapshot)
    
    Set tvCodeItems.ImageList = Nothing
    Set tvCodeItems.ImageList = ImageList1

    If rsSections.BOF And rsSections.EOF Then
        tvCodeItems.Nodes.Add , , "ROOT", "Code Library", "VIEWBOOKMARKS"
        BoldTreeNode tvCodeItems.Nodes("ROOT")
        Exit Sub
    End If
        
    TreeRedraw tvCodeItems.hwnd, False
    
    rsSections.MoveFirst
    Set tvCodeItems.ImageList = Nothing
    Set tvCodeItems.ImageList = ImageList1
'
' Populate the TreeView Nodes
'

    With tvCodeItems.Nodes
        .Clear
        .Add , , "ROOT", "Code Library", "VIEWBOOKMARKS"
'
' Make our Root Item BOLD
'
        BoldTreeNode tvCodeItems.Nodes("ROOT")
'
' Now add all nodes into TreeView, but under the root item.
' We reparent the nodes in the next step
'
        Do Until rsSections.EOF
            sParent = rsSections("ParentID").Value
            sKey = rsSections("ID").Value
            sText = rsSections("Description").Value
            Set nNode = .Add("ROOT", tvwChild, "C" & sKey, sText, "FOLDER")
'
' Record parent ID
'
            nNode.Tag = "C" & sParent
            rsSections.MoveNext
        Loop
    
    End With
'
' Here's where we rebuild the structure of the nodes
'
    For Each nNode In tvCodeItems.Nodes
        sParent = nNode.Tag
        If Len(sParent) > 0 Then        ' Don't try and reparent the ROOT !
            If sParent = "C0" Then
                sParent = "ROOT"
            End If
            Set nNode.Parent = tvCodeItems.Nodes(sParent)
        End If
    Next
'
' Now setup the images for each node in the treeview & set each node to
' be sorted if it has children
'
    For Each nNode In tvCodeItems.Nodes
        If nNode.Children = 0 Then
            nNode.Image = "CHILD"
        Else
            nNode.Sorted = True
        End If
    Next
    
    Set rsSections = Nothing
'
' Expand the Root Node
'
    tvCodeItems.Nodes("ROOT").Sorted = True
    tvCodeItems.Nodes("ROOT").Expanded = True
    
    TreeRedraw tvCodeItems.hwnd, True
    
    Exit Sub

vbErrorHandler:
    
    TreeRedraw tvCodeItems.hwnd, True
    
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " frmCodeLib::FillTree", , App.ProductName

End Sub

Private Sub Form_Resize()
    On Error Resume Next
'
' Make sure that all of our Controls are resized appropriately
'
    SizeControls tvCodeItems.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo vbErrorHandler
'
' Clear the treeview
'
    ClearTreeView
'
' Kill of any handles to objects in the Controls
'
    ctlBookMarkList.Terminate
    ctlCodeItemDetails.Terminate
'
' Kill the SysTray Icon
'
    KillSysTrayIcon
'
' Close the Database Connection
'
    If Not (mDB Is Nothing) Then
        mDB.Close
        Set mDB = Nothing
    End If
'
' Execute our unload procedure
'
    DoUnload
'
' That's it !
'
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " frmCodeLib::Form_Unload", , App.ProductName

End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Handle Splitter Movement - taken straight from the VB Template code
'
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbSplitting = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Handle Splitter Movement - taken straight from the VB Template code
'
    Dim sglPos As Single
    
    If mbSplitting Then
        sglPos = x + imgSplitter.Left
        If sglPos < lVSplitLimit Then
            picSplitter.Left = lVSplitLimit
        ElseIf sglPos > Me.Width - lVSplitLimit Then
            picSplitter.Left = Me.Width - lVSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Handle Splitter Movement - taken straight from the VB Template code
'
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbSplitting = False
End Sub

Private Sub mnuAbout_Click()
    Dim mFrm As frmAbout

On Error GoTo vbErrorHandler

'
' Show the about form
'
    Set mFrm = New frmAbout
    
    Load mFrm
    mFrm.Show vbModal
    Unload mFrm
    Set mFrm = Nothing

    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "frmCodeLib::mnuAbout_Click", , App.ProductName

End Sub

Private Sub mnuAddCode_Click()
'
' Add Code at the selected poit
'
    AddCode
End Sub

Private Sub AddCode()
'
' Add Code at the selected poit
'
On Error GoTo vbErrorHandler

    Dim sTitle As String
    Dim nNode As Node
    Dim oCodeItem As CCodeItem
    Dim iDO As IDataObject
    Dim sParentKey As String
    Dim nParentNode As Node
    
    Set nNode = tvCodeItems.SelectedItem
    
    If nNode.Key = "ROOT" Then
        sParentKey = "0"
    Else
        sParentKey = Right$(nNode.Key, Len(nNode.Key) - 1)
    End If
    
    If sParentKey <> "0" Then
        Set nParentNode = tvCodeItems.Nodes("C" & sParentKey)
        nParentNode.Image = "FOLDER"
        nParentNode.ExpandedImage = "FOLDER"
    End If
    
    Set iDO = New CCodeItem
    Set oCodeItem = iDO
    
    iDO.Initialise mDB
    oCodeItem.Description = "New Code item"
    oCodeItem.ParentKey = sParentKey
    iDO.Commit
    
    ctlCodeItemDetails.Initialise mDB, iDO
    
    Set nNode = tvCodeItems.Nodes.Add(tvCodeItems.SelectedItem, tvwChild, "C" & iDO.Key, oCodeItem.Description, "CHILD")
    Set tvCodeItems.SelectedItem = nNode
    nNode.EnsureVisible
    
    tvCodeItems.StartLabelEdit
    

    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " frmCodeLib::AddCode", , App.ProductName

End Sub

Private Sub mnuBookMark_Click()
'
' Add BookMark at the selected poit
'
    AddBookMark

End Sub

Private Sub AddBookMark()
'
' Add BookMark at the selected poit
'
On Error GoTo vbErrorHandler
    
    Dim sKey As String
    Dim iDO As IDataObject
    Dim frmBookMark As frmAddBookmark
    
    sKey = tvCodeItems.SelectedItem.Key
    If sKey = "ROOT" Then Exit Sub
    
    sKey = Right$(sKey, Len(sKey) - 1)
    
    Set iDO = New CCodeItem
    iDO.Initialise mDB, sKey
    
    Set frmBookMark = New frmAddBookmark
    
    Load frmBookMark
    
    With frmBookMark
        .Initialise mDB, iDO
        .Show vbModal, Me
        If Not (.Cancelled) Then
            ctlBookMarkList.Initialise mDB
            ctlBookMarkList.FindBookmark tvCodeItems.SelectedItem.Text
        End If
    End With
    
    Unload frmBookMark
    
    Set frmBookMark = Nothing
    Set iDO = Nothing


    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " frmCodeLib::AddBookMark", , App.ProductName

End Sub


Private Sub mnuDeleteCode_Click()
'
' Delete the selected CodeItem and all it's children
'
    DeleteCodeItem
End Sub

Private Sub mnuEdit_Click()
    Dim bIsRoot As Boolean
'
' Make menu items enabled/disabled as appropriate
'
    bIsRoot = (StrComp(tvCodeItems.SelectedItem.Key, "ROOT", vbTextCompare) = 0)
    mnuRename.Enabled = Not (bIsRoot)
    mnuDeleteCode.Enabled = Not (bIsRoot)
    mnuBookMark.Enabled = Not (bIsRoot)
    
End Sub

Private Sub mnuExit_Click()
'
' Quit !
'
    Unload Me
End Sub

Private Sub mnuExport_Click()
'
' Export All the CodeItems from the selected node
'
    ExportCodeItems
End Sub

Private Sub mnuImport_Click()
'
' Import a list of codeitems at the selected Node
'
    ImportCodeItems
End Sub

Private Sub mnuOpenDatabase_Click()
'
' Open a different VBCodeLibrary Database
'
    SelectDataBase
End Sub

Private Sub mnuRename_Click()
'
' Change the Label - remember, we only allow 50 Characters
'
    tvCodeItems.StartLabelEdit
End Sub

Private Sub mnuSettings_Click()
'
' Show Application settings
'
    ShowSettings
End Sub

Private Sub mnuViewBookMarks_Click()
'
' View the Bookmarks control
'
    ShowBookmarks Not (mbShowBookmarks)
    mnuViewBookMarks.Checked = mbShowBookmarks
End Sub

Private Sub picSysBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Here's where we handle the Icon Tray Messages
'
    Dim lMsg As Long
    Static bInHere As Boolean
    
    On Error GoTo vbErrorHandler
    
    lMsg = x / Screen.TwipsPerPixelX
    
    If bInHere Then Exit Sub
    
    bInHere = True
    Select Case lMsg
        Case WM_LBUTTONDBLCLK:
'
' On Mouse DoubleClick - Restore the window
'
            On Error Resume Next
            Me.Show
            
            If Me.WindowState = vbMinimized Then
                Me.WindowState = vbDefault
            End If
            Me.ZOrder
    End Select
    
    bInHere = False
    
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & "  " & Err.Source & " frmCodeLib::picSysBar_MouseMove", , App.ProductName
End Sub


Private Sub tbTools_ButtonClick(ByVal Button As ComctlLib.Button)
'
' Handle a toolbar button click
'
On Error GoTo vbErrorHandler

    Dim oNodeTarget As Node
    
    Select Case UCase$(Button.Tag)
        Case "OPENDB"
            SelectDataBase
        Case "PRINT", "FIND"
            MsgBox "This will be implemented in a later version", vbInformation, App.ProductName
                        
        Case "NEW"
            AddCode
        Case "VIEWBOOKMARKS"
            ShowBookmarks Not (mbShowBookmarks)
            
        Case "BOOKMARK"
            AddBookMark
            
        Case "DELETE"
            DeleteCodeItem
        
        Case "PREVIOUS"
            Set oNodeTarget = tvCodeItems.SelectedItem.Previous
            If Not (oNodeTarget Is Nothing) Then
                Set tvCodeItems.SelectedItem = oNodeTarget
                SelectCodeItem oNodeTarget.Key
            Else
                Set oNodeTarget = tvCodeItems.SelectedItem.Parent
                If Not (oNodeTarget Is Nothing) Then
                    Set tvCodeItems.SelectedItem = oNodeTarget
                    SelectCodeItem oNodeTarget.Key
                End If
            End If
        
        Case "NEXT"
            SendMessageLong tvCodeItems.hwnd, TVM_SELECTITEM, TVGN_NEXT, 0&
            Set oNodeTarget = tvCodeItems.SelectedItem.Next
            If Not (oNodeTarget Is Nothing) Then
                Set tvCodeItems.SelectedItem = oNodeTarget
                SelectCodeItem oNodeTarget.Key
            Else
                Set oNodeTarget = tvCodeItems.SelectedItem.Child
                If Not (oNodeTarget Is Nothing) Then
                    Set tvCodeItems.SelectedItem = oNodeTarget
                    SelectCodeItem oNodeTarget.Key
                End If

            End If
        
        Case "SETTINGS"
            ShowSettings
            
    End Select
    

    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source, "frmCodeLib::tbTools_ButtonClick", , App.ProductName

End Sub

Private Sub DeleteCodeItem()

On Error GoTo vbErrorHandler
 
    Dim sKey As String
    Dim oNode As Node
    Dim sMessage As String
    Dim iDO As IDataObject
    Dim oCodeItem As CCodeItem
    Dim oParentNode As Node
    Dim oWait As CWaitCursor
    
    Set oNode = tvCodeItems.SelectedItem
    
    sKey = oNode.Key
    
    If sKey = "ROOT" Then Exit Sub
        
    If oNode Is Nothing Then
        MsgBox "No Selected Record", , App.ProductName
        Exit Sub
    End If
    
    sMessage = "Delete selected Code "
    
    If oNode.Children > 0 Then
        sMessage = sMessage & "and all child records ?"
    Else
        sMessage = sMessage & "?"
    End If
    
    If MsgBox(sMessage, vbYesNo + vbExclamation, "Delete Code Record") = vbNo Then
        Exit Sub
    End If
    
    Set oParentNode = oNode.Parent
    
    Set oWait = New CWaitCursor
    oWait.SetCursor
    
    BeginTrans
    
    RecursiveDeleteCode oNode
    
    CommitTrans
    
    tvCodeItems.Nodes.Remove sKey
    
    ctlBookMarkList.Initialise mDB
    
    SelectCodeItem tvCodeItems.SelectedItem.Key
    
    If oParentNode.Children = 0 Then
       oParentNode.Expanded = False
        If Not oParentNode.Key = "ROOT" Then
            oParentNode.Image = "CHILD"
        End If
        
    End If
    
    Set oWait = Nothing
    
    Exit Sub

vbErrorHandler:
    Set oWait = Nothing
    
    Rollback
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " frmCodeLib::DeleteCodeItem", , App.ProductName

End Sub

Private Sub tmrDragTimer_Timer()
    Dim nHitNode As Node
    Static lCount As Long
'
' This timer has two functions :
'
' 1 - It will scroll the TreeView when the user is dragging
'
' 2 - It will auto-expand a node when the user drags over it for more than
'     half a second.
'
' Both pieces of code stolen from the MDSN.
'

    If mnDragNode Is Nothing Then
        tmrDragTimer.Enabled = False
        Exit Sub
    End If
    
    lCount = lCount + 1
    If lCount > 10 Then
    
        Set nHitNode = tvCodeItems.DropHighlight
        If nHitNode Is Nothing Then Exit Sub
        
        If nHitNode.Expanded = False Then
            nHitNode.Expanded = True
        End If
        lCount = 0
    End If
    
    If miScrollDir <> 0 Then
        If miScrollDir = -1 Then
            SendMessageLong tvCodeItems.hwnd, WM_VSCROLL, 0, 0
        Else
            SendMessageLong tvCodeItems.hwnd, WM_VSCROLL, 1, 0
        End If
    End If
    
    
End Sub

Private Sub tvCodeItems_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim iDO As IDataObject
    Dim oCodeItem As CCodeItem
    Dim sKey As String

On Error GoTo vbErrorHandler

    If Len(NewString) = 0 Then
        MsgBox "You must enter some text for a description", vbInformation, App.ProductName
        Cancel = True
        Exit Sub
    End If
    
    Set iDO = New CCodeItem
    Set oCodeItem = iDO
    
    sKey = tvCodeItems.SelectedItem.Key
    sKey = Right$(sKey, Len(sKey) - 1)
    
    iDO.Initialise mDB, sKey
    oCodeItem.Description = NewString
    StatusBar1.Panels(1).Text = NewString
    
    iDO.Commit
    
    SelectCodeItem tvCodeItems.SelectedItem.Key
    
    Exit Sub

vbErrorHandler:

    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "frmCodeLib::tvCodeItems_AfterLabelEdit", , App.ProductName

End Sub

Private Sub tvCodeItems_BeforeLabelEdit(Cancel As Integer)
    Dim lEditHWND As Long
'
' Limit the text entry size to 50 characters (as defined in our database
'
    
'
' Get the handle of the Edit Box on the treeview
'
    lEditHWND = SendMessageLong(tvCodeItems.hwnd, TVM_GETEDITCONTROL, 0, 0)
'
' Now limit the size to 50 characters
'
    If lEditHWND > 0 Then
        SendMessageLong lEditHWND, EM_LIMITTEXT, 50, 0
    End If
    
End Sub

Private Sub tvCodeItems_Collapse(ByVal Node As ComctlLib.Node)
    If Not Node.Key = "ROOT" Then
        Node.Image = "FOLDER"
    End If
    StatusBar1.Panels(1).Text = Node.Text
End Sub

Private Sub tvCodeItems_Expand(ByVal Node As ComctlLib.Node)
    If Not Node.Key = "ROOT" Then
        Node.ExpandedImage = "OPENFOLDER"
    End If
    StatusBar1.Panels(1).Text = Node.Text
End Sub

Private Sub tvCodeItems_KeyUp(KeyCode As Integer, Shift As Integer)
'
' Check for Delete Key pressed (Delete) and Insert (addNew)
'
    If tvCodeItems.SelectedItem.Key <> "ROOT" Then
        If KeyCode = vbKeyDelete Then
            DeleteCodeItem
            Exit Sub
        End If
    End If
    
    If KeyCode = vbKeyInsert Then
        AddCode
    End If
    
End Sub

Private Sub tvCodeItems_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set mnDragNode = tvCodeItems.HitTest(x, y)
End Sub

Private Sub tvCodeItems_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mnDragNode Is Nothing Then Exit Sub
    
    If Button = vbLeftButton Then
        If mnDragNode.Key <> "ROOT" Then
'
' Start Dragging !
'
            Set tvCodeItems.SelectedItem = mnDragNode
            tmrDragTimer.Interval = 100
            tmrDragTimer.Enabled = True
            tvCodeItems.OLEDrag
        End If
    Else
        Set mnDragNode = Nothing
    End If
    
End Sub

Private Sub tvCodeItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sKey As String
    Dim bIsRoot As Boolean
'
' Show Popup Menu
'
   
    If Button = vbRightButton Then
        bIsRoot = (StrComp(tvCodeItems.SelectedItem.Key, "ROOT", vbTextCompare) = 0)
        mnuRename.Enabled = Not (bIsRoot)
        mnuDeleteCode.Enabled = Not (bIsRoot)
        mnuBookMark.Enabled = Not (bIsRoot)
        PopupMenu mnuEdit
    End If
    
End Sub

Private Sub tvCodeItems_NodeClick(ByVal Node As ComctlLib.Node)
    SelectCodeItem Node.Key
End Sub

Private Sub DoToolBarLogic()
    Dim nNode As Node
    
    Set nNode = tvCodeItems.SelectedItem

    If nNode.Key = "ROOT" Then
        tbTools.Buttons("DELETE").Enabled = False
        tbTools.Buttons("BOOKMARK").Enabled = False
    Else
        tbTools.Buttons("DELETE").Enabled = True
        tbTools.Buttons("BOOKMARK").Enabled = True
    End If
    
End Sub

Private Sub InitControls()
'
' Make Toolbar Flat-Style
'

    Dim lStyle As Long
    Dim hToolbar As Long
    
    hToolbar = FindWindowEx(tbTools.hwnd, 0&, "ToolbarWindow32", vbNullString)
    
    lStyle = SendMessageLong(hToolbar, TB_GETSTYLE, 0&, 0&)
    
    If lStyle And TBSTYLE_FLAT Then
    '
    ' It's already flat
    '
    Else
        lStyle = lStyle Or TBSTYLE_FLAT
    End If
    
    SendMessageLong hToolbar, TB_SETSTYLE, 0, lStyle
    tbTools.Refresh
'
' Setup Track Select on the TreeView
'
    lStyle = GetWindowLong(tvCodeItems.hwnd, GWL_STYLE)
    
    lStyle = lStyle Or TVS_TRACKSELECT
    
    SetWindowLong tvCodeItems.hwnd, GWL_STYLE, lStyle
    
End Sub

Private Sub SelectCodeItem(ByVal sNodeKey As String)
    Dim iDO As IDataObject
    Dim sKey As String
    Dim oCodeItem As CCodeItem
'
' Select the relevant code item into our controls
'
    DoToolBarLogic
    
    If sNodeKey = "ROOT" Then
        ctlCodeItemDetails.Initialise mDB, Nothing
    Else
        Set iDO = New CCodeItem
        sKey = Right$(sNodeKey, Len(sNodeKey) - 1)
        iDO.Initialise mDB, sKey
        Set oCodeItem = iDO
        
'
' Setup our code window control
'
        StatusBar1.Panels(1).Text = oCodeItem.Description
        Set oCodeItem = Nothing
        
        ctlCodeItemDetails.Initialise mDB, iDO
'
' Setup our Bookmark list control
'
        ctlBookMarkList.FindBookmark tvCodeItems.SelectedItem.Text
        Set iDO = Nothing
    End If
    
End Sub

Private Sub GetLastDBName()
    Dim sDefaultDB As String
    Dim sDBName As String
'
' Get previously opened database name
'
    
    sDefaultDB = App.Path & "\" & DEFAULTDB
    
    sDBName = GetSetting("VBCodeLib", "Settings", "LastDB")
    
    If Len(sDBName) = 0 Then
        sDBName = sDefaultDB
        SaveSetting "VBCodeLib", "Settings", "LastDB", sDBName
    End If
    
    If Len(Dir$(sDBName)) > 0 Then
        msDBFileName = sDBName
    Else
        msDBFileName = ""
    End If

End Sub

Private Sub DoStartUp()
'
' Get settings
'
    Dim sDBName As String
    

    Me.Left = GetSetting("VBCodeLib", "Settings", "MainLeft", 2055)
    Me.Top = GetSetting("VBCodeLib", "Settings", "MainTop", 2175)
    Me.Width = GetSetting("VBCodeLib", "Settings", "MainWidth", 11000)
    Me.Height = GetSetting("VBCodeLib", "Settings", "MainHeight", 6210)
    tvCodeItems.Width = GetSetting("VBCodeLib", "Settings", "TreeWidth", 3270)
    mbShowBookmarks = GetSetting("VBCodeLib", "Settings", "ViewBookMarks", True)
'
' Turn off delete & bookmark tools
'
    tbTools.Buttons("DELETE").Enabled = False
    tbTools.Buttons("BOOKMARK").Enabled = False
            
'
' Check that the user wants to backup the database
' at startup
'
    If Not (moSettings.BackupDatabaseAtStart) Then
        Exit Sub
    End If
    
    sDBName = App.Path & "\codebackup.mdb"
'
' Kill the backup if it already exists
'
    If Len(Dir$(sDBName)) > 0 Then
        Kill sDBName
    End If
    
    If Len(msDBFileName) = 0 Then
        MsgBox "Cannot find the last opened database : " & msDBFileName, vbInformation, App.ProductName
    Else
        DBEngine.CompactDatabase msDBFileName, sDBName
    End If
    
    
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " frmCodeLib::DoStartUp", , App.ProductName

End Sub

Private Sub DoUnload()
    Dim sBackupName As String
    Dim sDBName As String
    
    On Error GoTo vbErrorHandler
'
' Save settings if required
'
    If moSettings.SaveFormLayout Then
        If Me.WindowState <> vbMinimized Then
            SaveSetting "VBCodeLib", "Settings", "MainLeft", Me.Left
            SaveSetting "VBCodeLib", "Settings", "MainTop", Me.Top
            SaveSetting "VBCodeLib", "Settings", "MainWidth", Me.Width
            SaveSetting "VBCodeLib", "Settings", "MainHeight", Me.Height
            SaveSetting "VBCodeLib", "Settings", "TreeWidth", tvCodeItems.Width
        End If
        SaveSetting "VBCodeLib", "Settings", "ViewBookMarks", mbShowBookmarks
    End If
    
    ctlBookMarkList.Terminate
    ctlCodeItemDetails.Terminate
'
' Check if we want to compact the database
'
    If Not (moSettings.CompactDatabaseOnExit) Then
        Exit Sub
    End If
'
' Compact it now !
'
    sDBName = msDBFileName
    sBackupName = App.Path & "\dbbackup.mdb"
'
' Check if the temporary backup database already exists
'
    If Len(Dir$(sBackupName)) > 0 Then
        Kill sBackupName
    End If
'
' Here's where we compact the database - first copying
' it to a temporary db
'
    If Not (mDB Is Nothing) Then
        mDB.Close
        Set mDB = Nothing
    End If
    
    If Len(sDBName) > 0 Then
        DBEngine.CompactDatabase sDBName, sBackupName
'
' Now we remove the database
'
        Kill sDBName
'
' Now we compact the temporary DB back into our original
' database
'
        DBEngine.CompactDatabase sBackupName, sDBName
'
' And Kill the backup !
'
        Kill sBackupName
    End If
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & " frmCodeLib::DoUnload", , App.ProductName
    
End Sub

Private Sub RecursiveDeleteCode(nNode As Node)
'
' Recursively Delete Node Items
'
    Dim nNodeChild As Node
    Dim iIndex As Integer
    Dim iDO As IDataObject
    Dim sKey As String
    
    Set iDO = New CCodeItem
    sKey = nNode.Key
    sKey = Right$(sKey, Len(sKey) - 1)
'
' Delete affected data object - we could have done this all through Access, but
' this is intended to show recursion through TreeView Nodes
'

    iDO.Initialise mDB, sKey
    iDO.Delete
    iDO.Commit
    Set iDO = Nothing
    
    Set nNodeChild = nNode.Child
    
    ' Now walk through the current parent node's children
    Do While Not (nNodeChild Is Nothing)
    
    ' If the current child node has it's own children...
        RecursiveDeleteCode nNodeChild
    ' Get the current child node's next sibling
        Set nNodeChild = nNodeChild.Next
    Loop
End Sub


Private Sub SetupSysTrayIcon()
    On Error GoTo vbErrorHandler
'
' Setup the System Tray Icon
'
    Dim tTrayStuff As NOTIFYICONDATA
    
    With tTrayStuff
        .cbSize = Len(tTrayStuff)
        .hwnd = picSysBar.hwnd
        .uId = 1&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "VBCodeLibrary Tool" & vbNullChar
        Shell_NotifyIcon NIM_ADD, tTrayStuff
    End With
 
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & "  " & Err.Description & " " & Err.Source & "::frmBrowser_SetupSysTrayIcon", , App.ProductName
End Sub


Private Sub KillSysTrayIcon()
    Dim t As NOTIFYICONDATA
'
' Kill the icon in the system tray
'
    With t
        .cbSize = Len(t)
        .hwnd = picSysBar.hwnd
        .uId = 1&
    End With
    
    Shell_NotifyIcon NIM_DELETE, t

End Sub


Private Sub BoldTreeNode(nNode As Node)
'
' Make a tree node bold
'
' Many thanks to VBNet for this code
'

On Error GoTo vbErrorHandler

    Dim TVI As TVITEM
    Dim lRet As Long
    Dim hItemTV As Long
    Dim lHwnd As Long
    
    Set tvCodeItems.SelectedItem = nNode
    
    lHwnd = tvCodeItems.hwnd
    hItemTV = SendMessageLong(lHwnd, TVM_GETNEXTITEM, TVGN_CARET, 0&)
    
    If hItemTV > 0 Then
        With TVI
            .hItem = hItemTV
            .mask = TVIF_STATE
            .stateMask = TVIS_BOLD
            lRet = SendMessageAny(lHwnd, TVM_GETITEM, 0&, TVI)
            .State = TVIS_BOLD
        End With
        lRet = SendMessageAny(lHwnd, TVM_SETITEM, 0&, TVI)
    End If
    
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source, , "frmCodeLib::BoldTreeNode"

End Sub

Private Sub ShowSettings()

On Error GoTo vbErrorHandler
'
' Show the Settings Dialog and update any settings
'
    Dim frmOpt As frmOptions
    
    Set frmOpt = New frmOptions
    Load frmOpt
    
    With frmOpt
        .Initialise moSettings
        .Show vbModal, Me
    End With
    Unload frmOpt
'
' Clear the form from memory
'
    Set frmOpt = Nothing

    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source, , "frmCodeLib::ShowSettings"

End Sub

Private Sub tvCodeItems_OLECompleteDrag(Effect As Long)
    Screen.MousePointer = vbDefault
    tmrDragTimer.Enabled = False
End Sub

Private Sub tvCodeItems_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Handle the dragging and-a dropping of treeview nodes here
'
    Dim sTmpStr As String
    Dim iDO As IDataObject
    Dim oTargetNode As Node
    Dim sParentKey As String
    Dim sKey As String
    Dim oCodeItem As CCodeItem
    Dim oOldParentNode As Node
    
    On Error Resume Next
'
' Check whether the clipboard data is in our special defined format
'
    sTmpStr = Data.GetFormat(miClipBoardFormat)
    
    If Err Or sTmpStr = "False" Then    ' it's not, so don't allow dropping
        Set mnDragNode = Nothing
        Set tvCodeItems.DropHighlight = Nothing
        Err.Clear
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    On Error GoTo vbErrorHandler
    
    If mnDragNode Is Nothing Then
        Set mnDragNode = Nothing
        Set tvCodeItems.DropHighlight = Nothing
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    Set oTargetNode = tvCodeItems.DropHighlight

'
    If oTargetNode Is Nothing Then
        Set mnDragNode = Nothing
        Set tvCodeItems.DropHighlight = Nothing
        Effect = vbDropEffectNone
        Exit Sub
    End If
  
    Set oOldParentNode = mnDragNode.Parent
   
        
    Set mnDragNode.Parent = oTargetNode
    
'
' Here's where we handle the drop - don't forget that we have to reparent
' our data objects to point to the new data object (or 0 if root)
'
    sParentKey = oTargetNode.Key

    If sParentKey = "ROOT" Then
        sParentKey = "0"
    Else
        sParentKey = Right$(sParentKey, Len(sParentKey) - 1)
    End If

    sKey = mnDragNode.Key
    sKey = Right$(sKey, Len(sKey) - 1)

'
' Initialise the dataobject and set it's new parent key
'
    Set iDO = New CCodeItem
    iDO.Initialise mDB, sKey
    Set oCodeItem = iDO
    oCodeItem.ParentKey = sParentKey
    iDO.Commit
    Set iDO = Nothing
    Set oCodeItem = Nothing

    Set tvCodeItems.DropHighlight = Nothing

    Set mnDragNode = Nothing
    tmrDragTimer.Enabled = False
    If oTargetNode.Key <> "ROOT" Then
        oTargetNode.ExpandedImage = "OPENFOLDER"
    End If
    If oOldParentNode.Children <= 1 And oOldParentNode.Key <> oTargetNode.Key Then
        If oOldParentNode.Key <> "ROOT" Then
            oOldParentNode.ExpandedImage = "CHILD"
            oOldParentNode.Image = "CHILD"
        End If
    End If
    
    
    
    Exit Sub

vbErrorHandler:
    
    Set mnDragNode = Nothing
    Set tvCodeItems.DropHighlight = Nothing
'
' This will more than likely be 'would cause a loop' or whatever
'
    MsgBox Err.Description, , App.ProductName
    Effect = vbDropEffectNone
    
End Sub

Private Sub tvCodeItems_OLEDragOver(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
        
    Dim sTmpStr As String
    Dim nTargetNode As Node
    On Error Resume Next
'
' First check that we allow this type of data to be dropped here
'
    sTmpStr = Data.GetFormat(miClipBoardFormat)
    
    If Err Or sTmpStr = "False" Then
        Err.Clear
        Effect = vbDropEffectNone
        Exit Sub
    End If
        
    Set nTargetNode = tvCodeItems.HitTest(x, y)
    If nTargetNode Is Nothing Then
        Set tvCodeItems.DropHighlight = Nothing
        Exit Sub
    End If

    If nTargetNode.Key = mnDragNode.Key Then
        Set tvCodeItems.DropHighlight = Nothing
        Effect = vbDropEffectNone
    Else
        Set tvCodeItems.DropHighlight = nTargetNode
    End If
    If y > 0 And y < 300 Then
        miScrollDir = -1
    ElseIf (y < tvCodeItems.Height) And y > (tvCodeItems.Height - 500) Then
        miScrollDir = 1
    Else
        miScrollDir = 0
    End If
    
End Sub

Private Sub tvCodeItems_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
    Dim byt() As Byte
'
' Place the key of the dragged item into the clipboard in our own format
' declared in GetClipboardFormat api
'
    AllowedEffects = vbDropEffectMove
    byt = mnDragNode.Key
    
    Data.SetData byt, miClipBoardFormat
    
    
End Sub

Private Sub ClearTreeView()
'
' Very fast Clearing of treeview control
'
' Thanks to Brad Martinez for discovering this .
'
    Dim lHwnd As Long
    Dim hItem As Long
    
    lHwnd = tvCodeItems.hwnd
    
    TreeRedraw tvCodeItems.hwnd, False
    
    
    Do
        hItem = SendMessageLong(lHwnd, TVM_GETNEXTITEM, TVGN_ROOT, &O0)
        If hItem > 0 Then
            SendMessageLong lHwnd, TVM_DELETEITEM, &O0, hItem
        Else
            Exit Do
        End If
    Loop
    
    TreeRedraw tvCodeItems.hwnd, True

End Sub

Private Sub SizeControls(ByVal x As Long)
    On Error Resume Next
'
' Size all controls based on the splitter bar, and whether we're
' showing the Bookmarks control
'
    Dim lHeightOffSet As Long
   
    'set the width
    If x < 1500 Then x = 1500
    If x > (Me.Width - 1500) Then x = Me.Width - 1500
    
    If mbShowBookmarks Then
        ctlBookMarkList.Height = Me.ScaleHeight * (2 / 8)
        lHeightOffSet = ctlBookMarkList.Height
    Else
        lHeightOffSet = 0
    End If
    
    With imgSplitter
        .Left = x
        .Width = 150
        .ZOrder
    End With
    
    With tvCodeItems
        .Move ScaleLeft, tbTools.Height, x, Me.ScaleHeight - (StatusBar1.Height + tbTools.Height + lHeightOffSet)
    End With
    
    With ctlCodeItemDetails
        .Move x + 25, tvCodeItems.Top, Me.ScaleWidth - (tvCodeItems.Width + 50), tvCodeItems.Height
    End With
    
    If mbShowBookmarks Then
        With ctlBookMarkList
            .Move ScaleLeft, tvCodeItems.Top + tvCodeItems.Height, ScaleWidth, lHeightOffSet
        End With
    End If
   
    imgSplitter.Top = tvCodeItems.Top
    imgSplitter.Height = tvCodeItems.Height

End Sub

Private Sub ShowBookmarks(ByVal bShow As Boolean)
'
' Show / hide bookmarks
'
    mbShowBookmarks = bShow
    ctlBookMarkList.Visible = mbShowBookmarks
    Form_Resize
    mnuViewBookMarks.Checked = mbShowBookmarks
    tbTools.Buttons("VIEWBOOKMARKS").Value = IIf(mbShowBookmarks, tbrPressed, tbrUnpressed)
End Sub

Private Function SetupDBConnection() As Boolean
'
' Setup Database Connection
'
' This routine will also update any previous versions of the
' database to the new required version !
'

    Dim bValidDatabase As Boolean
    
    On Error GoTo vbErrorHandler
    
    If Not (mDB Is Nothing) Then
        mDB.Close
        Set mDB = Nothing
    End If
    
    If Len(msDBFileName) = 0 Then
        SelectDataBase
    End If
    
    If Len(msDBFileName) > 0 Then
        Set mDB = Workspaces(0).OpenDatabase(msDBFileName)
    '
    ' Here's where we setup the Version Specific Tables if they don't exist
    '
        SetupVersionTable
        SetupCodeFilesTable
            
        SetupDBConnection = True
    End If
    
    Exit Function


vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source
    
End Function

Private Sub SelectDataBase()
    Dim sDBName As String
    
    sDBName = GetFileName(eOpenFileName, "", "Select a CodeLibrary Database", "VBCodeLibrary Files|*.mdb" & vbNullChar & vbNullChar)
    
    If Len(sDBName) = 0 Then
    '
    ' No Change
    '
        Exit Sub
    End If
    
    DoUnload
    SaveSetting "VBCodeLib", "Settings", "LastDB", sDBName
    
    msDBFileName = sDBName
    If SetupDBConnection = True Then
        ClearTreeView
        FillTree
        ctlBookMarkList.Initialise mDB
        EnableControls True
    Else
    
    ' Disable appropriate controls
    
        EnableControls False
        
    End If
    
End Sub

Private Function GetFileName(ByVal DialogType As eGetFileDialog, _
        ByRef sFilename As String, _
        ByVal sDialogTitle As String, _
        Optional sFilter As String) As String
    
    On Error GoTo vbErrorHandler
    
    If Len(sFilter) = 0 Then
        sFilter = "All Files|*.*"
    End If
    
    If Len(CommonDialog1.InitDir) = 0 Then
        CommonDialog1.InitDir = App.Path
    End If
    
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = sDialogTitle
    
    If Len(sFilename) > 0 Then
        CommonDialog1.filename = sFilename
    Else
        CommonDialog1.filename = ""
    End If
    If Len(sFilter) > 0 Then
        CommonDialog1.Filter = sFilter
    Else
        CommonDialog1.Filter = ""
    End If
    
    CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly
    
    If DialogType = eOpenFileName Then
        CommonDialog1.ShowOpen
    Else
        CommonDialog1.Flags = CommonDialog1.Flags + cdlOFNOverwritePrompt
        CommonDialog1.ShowSave
    End If
    sFilename = CommonDialog1.filename
    
    If Len(sFilename) > 0 Then
        GetFileName = sFilename
    End If
    Exit Function
    
vbErrorHandler:
    If Err.Number = 32755 Then
        GetFileName = ""
        Exit Function
    Else
        MsgBox Err.Number & " " & Err.Source & " " & Err.Description, vbCritical, App.ProductName
    End If
    
End Function

Private Sub SetupVersionTable()
    Dim oRS As Recordset
    Dim fldField As Field
    Dim tblVersion As TableDef
    Dim sDBVersion As String
'
' Make sure that the user has the CodeDBVersion Table
'
' This will be used for future Backwards Compatibility
'

    On Error GoTo vbErrorHandler
    
    Set oRS = mDB.OpenRecordset("select version from CodeDBVersion")
    
    If oRS.EOF And oRS.BOF Then
        oRS.AddNew
        oRS.Fields("version").Value = App.Major & "." & App.Minor & "." & App.Revision
        oRS.Update
        oRS.Bookmark = oRS.LastModified
    End If
    
    sDBVersion = oRS.Fields("version")
    
    Me.Caption = "VBCodeLibrary Tool (" & msDBFileName & " - DB Version " & sDBVersion & ")"
    oRS.Close
    
    Exit Sub

vbErrorHandler:
    
    If Err.Number = 3078 Then
    '
    ' Add in the CodeDBVersion Table
    '
        Set tblVersion = New TableDef
        Set fldField = New Field
        
        Set tblVersion = mDB.CreateTableDef("CodeDBVersion")
        
        tblVersion.Fields.Append tblVersion.CreateField("Version", dbText)
        mDB.TableDefs.Append tblVersion
        
        Set oRS = mDB.OpenRecordset("CodeDBVersion")
        oRS.AddNew
        oRS.Fields("version").Value = App.Major & "." & App.Minor & "." & App.Revision
        oRS.Update
        oRS.Close
        
        Resume
        
    ElseIf Err.Number = 0 Then
    
    
    End If
    
End Sub

Private Sub SetupCodeFilesTable()
    Dim oRS As Recordset
    Dim tblCodeFiles As TableDef
    Dim fldField As Field
    
'
' Make sure that the database contains the CodeFiles Table
'
' If it doesn't, then create it inside the database.
'
' This ensures that the user can continue using their existing database.
'
'

    On Error GoTo vbErrorHandler
    
    Set oRS = mDB.OpenRecordset("codefiles")
    oRS.Close
    
    Exit Sub

vbErrorHandler:
    
    If Err.Number = 3078 Then
    '
    ' Add in the CodeFiles Table
    '
        
        Set tblCodeFiles = mDB.CreateTableDef("CodeFiles")
        
        tblCodeFiles.Fields.Append tblCodeFiles.CreateField("ID", dbLong)
        tblCodeFiles.Fields("ID").Attributes = dbAutoIncrField
        tblCodeFiles.Fields.Append tblCodeFiles.CreateField("CodeID", dbLong)
        tblCodeFiles.Fields.Append tblCodeFiles.CreateField("Description", dbText, 50)
        tblCodeFiles.Fields.Append tblCodeFiles.CreateField("File", dbLongBinary)
        tblCodeFiles.Fields.Append tblCodeFiles.CreateField("OrigDateTime", dbDate)
        tblCodeFiles.Fields.Append tblCodeFiles.CreateField("DateAdded", dbDate)
        
        mDB.TableDefs.Append tblCodeFiles
        
        Resume
        
    ElseIf Err.Number = 0 Then
    
    
    End If

End Sub

Private Sub ImportCodeItems()
'
' This routine imports items in the VCL file into the Database
'
    Dim nNode As Node
    Dim iFile As Integer
    Dim sUseFileName As String
    Dim oCodeItem As CCodeItem
    Dim iDO As IDataObject
    Dim lCount As Long
    Dim oImport As ImportData
    Dim sParentKey As String
    Dim sTopParentKey As String
    Dim oColl As Collection
    Dim oWait As CWaitCursor
    Dim lNumCodeItems As Long
    Dim sTmp As String
    
    Dim oHeader As FileHeader
    
On Error GoTo vbErrorHandler
'
' Get selected Node
'
    Set nNode = tvCodeItems.SelectedItem
'
' If No Node Selected (very unlikely) then exit
'
    If nNode Is Nothing Then Exit Sub
    
'
' Get Import File Name
'
    sUseFileName = GetFileName(eOpenFileName, "", "Import Data From File :", "VBCodeLibrary Export|*.vcl")
'
' If no name selected then quit
'
    If Len(sUseFileName) = 0 Then Exit Sub
    
'
' Get FileHandle
'
    iFile = FreeFile
    
'
' Get Top Parent Key
'
    If nNode.Key = "ROOT" Then
        sTopParentKey = "0"
    Else
        sTopParentKey = Right$(nNode.Key, Len(nNode.Key) - 1)
    End If
'
' Set Cursor to HourGlass
'
    Set oWait = New CWaitCursor
    oWait.SetCursor
'
' Setup Our Collection Internally
'
    Set oColl = New Collection
    
'
' Place all of the Import into a Transaction for Speed & rollback opportunity
'
    BeginTrans
    
'
' Open the file
'
    Open sUseFileName For Binary Access Read As iFile
    
    lCount = 1
    
    Get #iFile, , oHeader
    
    prgBar.Min = 1
    prgBar.Max = oHeader.lNumberOfRecords '+ 5
    
    StatusBar1.Panels(1).Text = "Importing Items...."
    DoEvents
    ShowProgressInStatusBar True
'
' Now loop through the records in the file
'
    For lCount = 1 To oHeader.lNumberOfRecords
        
'
' Get each record until empty
'
        Get #iFile, , oImport
        
        If oImport.sName = "" Then Exit For
        
'
' Create a new CodeItem for the record
'
        Set iDO = New CCodeItem
        Set oCodeItem = iDO
        iDO.Initialise mDB
        
'
' Setup the CodeItems values
'
        oCodeItem.Code = oImport.sStoredCode
        oCodeItem.Description = oImport.sName
        oCodeItem.Example = oImport.sUsage
        oCodeItem.Notes = oImport.sNotes
'
' If this is the first one, then set it's parent to the selected Node database key
'
        If lCount = 1 Then
            oCodeItem.ParentKey = sTopParentKey
        End If
'
' Write the new record away
'
        iDO.Commit
'
' Now build up our key object for recreating the Tree Structure
'
'        Set oKeys = New CImportKey
'        oKeys.sNewID = iDO.Key
'        oKeys.sOldID = oImport.sOriginalID
    
'
' Add it to the collection - indexed by Original Key
'
    '    oColl.Add oKeys, oKeys.sOldID
        oColl.Add iDO.Key, oImport.sOriginalID
        
        
'
' If we're not on the first item to be imported, restructure the items
'
        If lCount > 1 Then
            sParentKey = oImport.sParentID
                    
            If Len(sParentKey) > 0 And sParentKey <> "0" Then
                oCodeItem.ParentKey = oColl.Item(sParentKey) '.sNewID
            Else
                oCodeItem.ParentKey = sTopParentKey
            End If
            iDO.Commit
        End If
        Set iDO = Nothing
        Set oCodeItem = Nothing
        prgBar.Value = lCount
        sParentKey = ""
    Next
    
'
' Close the file
'
    Close iFile
'
' Commit all of our database work
'
    ShowProgressInStatusBar False
    StatusBar1.Panels(1).Text = ""
    CommitTrans
'
' Fill the tree with all records from the database
'
    FillTree
'
' Now, get the original Node that was the TopParent, and make sure
' that it's expanded, and visible
'
    If Len(sTopParentKey) > 0 And sTopParentKey <> "0" Then
        Set nNode = tvCodeItems.Nodes("C" & sTopParentKey)
        Set tvCodeItems.SelectedItem = nNode
        nNode.Expanded = True
        nNode.EnsureVisible
    End If
'
' Restore the cursor
'
    Set oWait = Nothing
    
'
' Notify the User of success
'
    MsgBox "Successfully imported " & lCount - 1 & " Code snippets.", vbInformation, App.ProductName
    
    Exit Sub

vbErrorHandler:
'
' Restore the cursor
'
    Set oWait = Nothing
'
' Rollback the database work
'
    Rollback
    ShowProgressInStatusBar False
    StatusBar1.Panels(1).Text = ""
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & vbCrLf & vbCrLf & "frmCodeLib::ImportCodeItems"

End Sub

Private Sub ExportCodeItems()
    Dim nNode As Node
    Dim iFile As Integer
    Dim sUseFileName As String
    Dim oWait As CWaitCursor
    Dim lNumToExport As Long
    Dim oHeader As FileHeader
    
'
' Here's where we export the items to a file
'

'
' Get Selected Node
'
    Set nNode = tvCodeItems.SelectedItem
'
' Check if node exists (very unlikely that it doesn't)
'
    If nNode Is Nothing Then Exit Sub
'
' Get File Name to Export to
'
    sUseFileName = GetFileName(eSaveFileName, "", "Export Data To File :", "VBCodeLibrary Export|*.vcl")
    
    If Len(sUseFileName) = 0 Then Exit Sub
'
' Show Wait Cursor
'
    Set oWait = New CWaitCursor
    oWait.SetCursor
    
'
' Get File Handle
'
    iFile = FreeFile
    
    On Error Resume Next
    Kill sUseFileName
    On Error GoTo 0
    
    Open sUseFileName For Binary Access Write As iFile
    
    oHeader.lNumberOfRecords = RecursiveCountNodes(tvCodeItems.SelectedItem, True)
    
    If nNode.Key <> "ROOT" Then
        oHeader.lNumberOfRecords = oHeader.lNumberOfRecords + 1
    End If
    
    prgBar.Min = 1
    prgBar.Max = oHeader.lNumberOfRecords
    prgBar.Value = 1
    StatusBar1.Panels(1).Text = "Exporting Code Items...."
    DoEvents
    ShowProgressInStatusBar True
    
    Put #iFile, , oHeader
    
'
' Recursively Write Each item Away to the File
'
    RecursiveExportCode nNode, iFile
    
    Close iFile
    Set oWait = Nothing
    ShowProgressInStatusBar False
    StatusBar1.Panels(1).Text = ""
    
End Sub

Private Sub RecursiveExportCode(nNode As Node, ByVal iFileNumber As Integer)
'
' Recursively Delete Node Items
'
    Dim nNodeChild As Node
    Dim iIndex As Integer
    Dim iDO As IDataObject
    Dim oCodeItem As CCodeItem
    
    Dim sKey As String
    Dim oExport As ImportData
    
    Set iDO = New CCodeItem
    Set oCodeItem = iDO
    
    On Error Resume Next
    prgBar.Value = prgBar.Value + 1
    
    On Error GoTo 0
    
    sKey = nNode.Key
'
' Get Details for item (as long as it's not the Root Item)
'
    If StrComp(sKey, "ROOT", vbTextCompare) <> 0 Then
        sKey = Right$(sKey, Len(sKey) - 1)
'
        iDO.Initialise mDB, sKey
        oExport.sOriginalID = iDO.Key
        oExport.sParentID = oCodeItem.ParentKey
        oExport.sName = oCodeItem.Description
        oExport.sNotes = oCodeItem.Notes
        oExport.sParentName = nNode.Parent.Key
        oExport.sStoredCode = oCodeItem.Code
        oExport.sUsage = oCodeItem.Example
        
        Put #iFileNumber, , oExport
        
        Set iDO = Nothing
        Set oCodeItem = Nothing
    End If
    
    
    Set nNodeChild = nNode.Child
'
' Now walk through the current parent node's children
'
    Do While Not (nNodeChild Is Nothing)
'
' If the current child node has it's own children...
'
        RecursiveExportCode nNodeChild, iFileNumber
'
' Get the current child node's next sibling
'
        Set nNodeChild = nNodeChild.Next
    Loop
End Sub

Private Function RecursiveCountNodes(nNode As Node, Optional bResetToZero As Boolean = False) As Long
'
    Dim nNodeChild As Node
    Dim iIndex As Integer
    Static lCount As Long
    
    If bResetToZero Then
        lCount = 0
    End If
    
'
' Get Details for item (as long as it's not the Root Item)
'
    Set nNodeChild = nNode.Child
'
' Now walk through the current parent node's children
'
    Do While Not (nNodeChild Is Nothing)
        lCount = lCount + 1
'
' If the current child node has it's own children...
'
        RecursiveCountNodes nNodeChild, False
'
' Get the current child node's next sibling
'
        Set nNodeChild = nNodeChild.Next
            
    Loop
    RecursiveCountNodes = lCount
    
End Function

Private Sub TreeRedraw(ByVal lHwnd As Long, ByVal bRedraw As Boolean)
'
' Utility Routine for TreeRedraw on/of
'
    SendMessageLong lHwnd, WM_SETREDRAW, bRedraw, 0

End Sub

Private Sub EnableControls(ByVal bEnable As Boolean)
    Dim oTool As Button
    
    For Each oTool In tbTools.Buttons
        oTool.Enabled = bEnable
    Next
    tbTools.Enabled = bEnable
    tvCodeItems.Enabled = bEnable
    mnuEdit.Enabled = bEnable
    mnuView.Enabled = bEnable
    If bEnable Then
        ctlCodeItemDetails.Initialise mDB, Nothing
    Else
        ctlCodeItemDetails.Initialise Nothing, Nothing
    End If

End Sub

Private Sub ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean)

    Dim tRC As RECT
    
    If bShowProgressBar Then
'
' Get the size of the Panel (2) Rectangle from the status bar
' remember that Indexes in the API are always 0 based (well,
' nearly always) - therefore Panel(2) = Panel(1) to the api
'
'
        SendMessageAny StatusBar1.hwnd, SB_GETRECT, 1, tRC
'
' and convert it to twips....
'
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
'
' Now Reparent the ProgressBar to the statusbar
'
        With prgBar
            SetParent .hwnd, StatusBar1.hwnd
            .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
            .Value = 1
        End With
        
    Else
'
' Reparent the progress bar back to the form and hide it
'
        SetParent prgBar.hwnd, Me.hwnd
        prgBar.Visible = False
    End If

End Sub

