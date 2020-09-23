VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl ctlBookmarks 
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3060
   ScaleWidth      =   5235
   Begin ComctlLib.ListView lvBookMarks 
      Height          =   2040
      Left            =   75
      TabIndex        =   0
      Top             =   330
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   3598
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Caption         =   " Bookmarks :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   15
      TabIndex        =   1
      Top             =   -15
      Width           =   5160
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveBookMark 
         Caption         =   "&Remove Bookmark"
      End
   End
End
Attribute VB_Name = "ctlBookmarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Bookmarks Control
'
' Chris Eastwood Feb 1999
'
' This control displays a list of the bookmarks in the system.
'
Private mDB As Database ' Our pointer to the Database

'
' Public events raised from the control
'
Public Event ViewBookMark(oCodeItem As IDataObject)
Public Event BookMarkRemoved(ByVal sCodeID As String)

Public Sub Initialise(oDB As Database)
'
' Entry Point
'

    Dim tRS As Recordset
    Dim sSql As String
    Dim lCount As Long
    Dim li As ListItem
    Dim sTitle As String
    
'
' Setup the Controls internals
'
    SetupControl
'
' Clear any existing listview items
'
    SendMessageLong lvBookMarks.hwnd, WM_SETREDRAW, False, &O0
    
    lvBookMarks.ListItems.Clear
'
' Record Database object
'
    Set mDB = oDB
    
    If mDB Is Nothing Then Exit Sub
    
'
' Build SQL string to get all details from Database
'
    sSql = "SELECT Bookmarks.ID, Bookmarks.CodeID, CodeItems.Description, " & _
            "Bookmarks.Description FROM Bookmarks INNER JOIN " & _
            "CodeItems ON Bookmarks.codeID = CodeItems.ID"
    
    Set tRS = mDB.OpenRecordset(sSql, dbOpenSnapshot)
   
    If Not (tRS.BOF And tRS.EOF) Then
'
' Add items to the listview
' Listitem Key = Bookmark Key
' Listitem Tag = CodeItem Key
'
        tRS.MoveFirst
        
        Do While Not (tRS.EOF)
            Set li = lvBookMarks.ListItems.Add(, "B" & tRS.Fields("ID"), tRS.Fields("CodeItems.Description").Value)
            li.Tag = "C" & tRS.Fields("CodeID").Value ' CodeID
            sTitle = tRS.Fields("bookmarks.Description")
            ReplaceAll sTitle, vbCrLf, " "
            li.SubItems(1) = sTitle
            tRS.MoveNext
        Loop
        
    End If
    SendMessageLong lvBookMarks.hwnd, WM_SETREDRAW, True, &O0
    
    AutoSizeListViewColumns lvBookMarks, True
    
    
End Sub

Public Sub Terminate()
'
' Release our database object
'
    lvBookMarks.ListItems.Clear
    Set mDB = Nothing
End Sub

Private Sub SetupControl()
    Dim lStyle As Long
    Dim lHeaderHwnd As Long
    Dim llvHwnd As Long
    
    llvHwnd = lvBookMarks.hwnd

'
' Give the ListView full-row-select capability
'
    SendMessageLong llvHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, True
End Sub

Private Function HasHorizontalScrollBar(ByVal lHwnd As Long) As Boolean
'
' General purpose routine to see if ANY control has a Horizontal ScrollBar
'
    Dim lStyle As Long
    
    lStyle = GetWindowLong(hwnd, GWL_STYLE)
    HasHorizontalScrollBar = lStyle And WS_HSCROLL
End Function


Public Function RemoveBookMark(ByVal sCodeID As String)
'
' Remove an item from the ListView
'
    On Error Resume Next ' it may not exist in the listview
    lvBookMarks.ListItems.Remove ("C" & sCodeID)
    
End Function

Private Sub lvBookMarks_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'
' Sort the listview by the clicked column
'
    lvBookMarks.SortKey = ColumnHeader.Index - 1
    lvBookMarks.Sorted = True
End Sub

Private Sub SynchItem(li As ListItem)
    Dim sKey As String
    Dim oCodeItem As IDataObject
'
' Synchronize the ListView item with the treeview in the parent form
'
    sKey = Right$(li.Tag, Len(li.Tag) - 1)
    If Len(sKey) = 0 Then
    '
    ' Should never happen, but you never know !
    '
        MsgBox "No Related Code !"
        Exit Sub
    End If
    
    Set oCodeItem = New CCodeItem
    oCodeItem.Initialise mDB, sKey
    
    RaiseEvent ViewBookMark(oCodeItem)
    Set oCodeItem = Nothing
    
End Sub

Private Sub lvBookMarks_ItemClick(ByVal Item As ComctlLib.ListItem)
    SynchItem Item
End Sub

Private Sub lvBookMarks_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Display the popup menu
'
    If Button = vbRightButton And lvBookMarks.ListItems.Count > 0 Then
        PopupMenu mnuPopup
    End If
    
End Sub

Private Sub mnuRemoveBookMark_Click()
    Dim iDO As IDataObject
    Dim sBMKey As String
    Dim sCodeKey As String
'
' Delete a bookmark
'
    
'
' Check for Selected item
'
    If lvBookMarks.SelectedItem Is Nothing Then Exit Sub
'
' Get Keys of Bookmark and Code Item
'
    sBMKey = lvBookMarks.SelectedItem.Key
    sBMKey = Right$(sBMKey, Len(sBMKey) - 1)
    sCodeKey = lvBookMarks.SelectedItem.Tag
    sCodeKey = Right$(sCodeKey, Len(sCodeKey) - 1)
'
' Delete the Bookmark Item
'
    Set iDO = New CBookmark
    iDO.Initialise mDB, sBMKey
    iDO.Delete
    iDO.Commit
    
    Set iDO = Nothing
'
' Tell our Main Form that a book mark has been deleted - it may
' want to know
'
    RaiseEvent BookMarkRemoved(sCodeKey)
'
' Remove the selected item from the listview
'
    lvBookMarks.ListItems.Remove lvBookMarks.SelectedItem.Key
    
End Sub


Private Sub UserControl_Initialize()
'
' Add the relevant columns to the listview control
'
    With lvBookMarks
        .View = lvwReport
        .ColumnHeaders.Add , , "Code Section"
        .ColumnHeaders.Add , , "Bookmark Description"
    End With
    
End Sub

Public Sub FindBookmark(ByVal sCodeName As String)
    Dim li As ListItem
'
' Find requested string in column 1
'
    With lvBookMarks
        Set li = .FindItem(sCodeName)
        If Not li Is Nothing Then
            li.EnsureVisible
            Set lvBookMarks.SelectedItem = li
        End If
    End With
    
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Static bInHere As Boolean
'
' Make sure we resize correctly
'
    If bInHere Then Exit Sub ' to stop potential stack problems
    
    bInHere = True
    
    Label1.Move ScaleLeft + 30, Label1.Top, ScaleWidth - 30, Label1.Height
    lvBookMarks.Move ScaleLeft + 15, Label1.Top + Label1.Height + 15, ScaleWidth - 15, ScaleHeight - (Label1.Top + Label1.Height + 15)
'
' Now autosize the Description column so it fills the remainder
' of the listview (it just looks nicer imho)
'
    AutoSizeLastColumn lvBookMarks
    
    bInHere = False
    
End Sub

