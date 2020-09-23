VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl ctlFileDetails 
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3960
   ScaleWidth      =   6045
   Begin VB.Frame fraFile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   5460
      Begin ComctlLib.ListView lvFiles 
         Height          =   2865
         Left            =   60
         TabIndex        =   1
         Top             =   165
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   5054
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         _Version        =   327682
         Icons           =   "imgFiles"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
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
         NumItems        =   0
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuAddFile 
         Caption         =   "Add File"
      End
      Begin VB.Menu mnuExportFile 
         Caption         =   "Export File"
      End
      Begin VB.Menu mnuDeleteFile 
         Caption         =   "Delete File"
      End
   End
End
Attribute VB_Name = "ctlFileDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' ctlFileDetails
'
' Chris Eastwood (mailto:chris.eastwood@codeguru.com)
'
' For the VBCodeLibrary Project
'
' This control consists of a listview which displays all files that
' have been stored in the database associated with a CodeItem object.
'

Private moCodeItem As CCodeItem     ' Our Referenced CodeItem
Private mDataObject As IDataObject  ' Pointer to Referenced CodeItem IDataObject Interface
Private mDB As Database             ' Pointer to the DataBase
'
' Bubble up events to ask for filename
'

Public Event RequestFileName(ByVal DialogType As eGetFileDialog, ByRef sFilename As String, ByVal sDialogTitle As String)

Public Sub Initialise(oDB As Database, iDO As IDataObject)
'
' Initialise the Control (not the same as UserControl Initialise)
'

'
' Clear any listitems from the listview
'
    lvFiles.ListItems.Clear
'
' Record the associated CodeItem Object
'
    Set mDataObject = iDO
    Set moCodeItem = iDO
    Set mDB = oDB
'
' Populate the ListView
'
    PopulateListView
'
' Autosize the listview columns
'
    AutoSizeListViewColumns lvFiles, False
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "::ctlFileDetails_Initialise"
    
End Sub

Public Sub Terminate()
'
' Terminate the internal data of this control
'

'
' Clear the listview
'
    lvFiles.ListItems.Clear
'
' Clear our object references
'
    Set mDataObject = Nothing
    Set moCodeItem = Nothing
    Set mDB = Nothing
    
End Sub


Private Sub lvFiles_DblClick()
'
' Bring up the Export Function as Double-Click on the ListView
'
    If Not (lvFiles.SelectedItem Is Nothing) Then
        If lvFiles.SelectedItem.Selected Then
            ExportFile
        End If
    End If
End Sub

Private Sub lvFiles_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        DeleteFile
    End If
    
End Sub

Private Sub lvFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim li As ListItem
'
' Show Popup Menu for the selected Item in the listview
'
    
'
' Was the RightMouseButton Pressed ?
'
    If Button = vbLeftButton Then Exit Sub
    
'
' Get Selected ListViewItem
'
    Set li = lvFiles.HitTest(x, y)
'
' Setup Appropriate Menus
'
    If li Is Nothing Then
        mnuAddFile.Enabled = True
        mnuDeleteFile.Enabled = False
        mnuExportFile.Enabled = False
    Else
        mnuAddFile.Enabled = True
        mnuDeleteFile.Enabled = True
        mnuExportFile.Enabled = True
    End If
'
' Show the Popupmenu mnuPopup with 'Export File' as the default
'
    PopupMenu mnuPopup, , , , mnuExportFile

    
End Sub

Private Sub lvFiles_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Handle files dragged & dropped on this listview
'
    Dim lFiles As Long
    Dim sFilename As String
    Dim iDO As IDataObject
    Dim oFO As CFileObject
    Dim lCount As Long
    
    On Error GoTo vbErrorHandler
'
' Check whether it was a file / group of files dropped on to the listview
'
    If Data.GetFormat(vbCFFiles) = True Then
        lFiles = Data.Files.Count
                    
        If lFiles > 0 Then
'
' Add each file into the database for our associated CodeItem Object
'
            For lCount = 1 To lFiles
                sFilename = Data.Files(lCount)
                AddCodeFile sFilename
            Next
        End If
    End If
    
'
' Populate the listview
'
    PopulateListView
    
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "::lvFiles_OLEDragDrop"
    
End Sub

Private Sub mnuAddFile_Click()
'
' Add File clicked
'
    AddFile
End Sub

Private Sub mnuDeleteFile_Click()
'
' Delete File Clicked
'
    DeleteFile
End Sub

Private Sub mnuExportFile_Click()
'
' Export File Clicked
'
    ExportFile
End Sub


Private Sub UserControl_Initialize()
    Dim lHeaderHwnd As Long
    Dim lStyle As Long
    Dim llvHwnd As Long
    
'
' Setup the ListView Columns
'
    llvHwnd = lvFiles.hwnd
    
    lvFiles.View = lvwReport
    lvFiles.ColumnHeaders.Add , , "Original File Name              "
    lvFiles.ColumnHeaders.Add , , "File Date/Time", , 1
    lvFiles.ColumnHeaders.Add , , "Date Added", , 1
    lvFiles.ColumnHeaders.Add , , "Size (KB)", , 1
'
' Setup Full Row Select on the listview
'
    SendMessageLong llvHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, True
'
' Setup Flat Headers on the ListView
'
    lHeaderHwnd = SendMessageLong(llvHwnd, LVM_GETHEADER, 0, ByVal 0&)
    
    lStyle = GetWindowLong(lHeaderHwnd, GWL_STYLE)
    If lStyle And HDS_BUTTONS Then
        lStyle = lStyle Xor HDS_BUTTONS
    End If
'
' Set the new ListView Style
'
    If lStyle > 0 Then
        SetWindowLong lHeaderHwnd, GWL_STYLE, lStyle
    End If
    
End Sub

Private Sub UserControl_Resize()
'
' Do on-error-resume-next to ignore invalid control sizing
'
    On Error Resume Next
    fraFile.Move UserControl.ScaleLeft, UserControl.ScaleTop, UserControl.ScaleWidth, UserControl.ScaleHeight
    lvFiles.Move fraFile.Left + 50, lvFiles.Top, fraFile.Width - 80, fraFile.Height - (lvFiles.Top + 55)
'
' Resize all columns so they fit in the listview
'
    AutoSizeLastColumn lvFiles
End Sub

Private Sub AddFile()
'
' Add an associated file to the codeitem into the database
'
    Dim sFilename As String

    On Error GoTo vbErrorHandler
'
' Bubble up events to get filename
'
    RaiseEvent RequestFileName(eOpenFileName, sFilename, "Select A File To Import")
'
' If chosen filename = "" then user clicked cancel
'
    If Len(sFilename) = 0 Then Exit Sub
'
' Add the file to the codeitem
'
    AddCodeFile sFilename
'
' Repopulate the listview
'
    PopulateListView
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "::ctlFileDetails_AddFile"
End Sub

Private Sub DeleteFile()
    Dim sKey As String
    Dim iDO As IDataObject
    Dim iFO As CFileObject
'
' Delete the associated file from the database
'
    On Error GoTo vbErrorHandler
'
' Get Key from Selected Item
'
    sKey = lvFiles.SelectedItem.Key
    sKey = Right$(sKey, Len(sKey) - 3)
        
    If MsgBox("Do you really want to delete this file from the database ?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete File from Database") = vbYes Then
        Set iDO = New CFileObject
        Set iFO = iDO
'
' Delete It
'
        iDO.Initialise mDB, sKey
        iDO.Delete
        iDO.Commit
        
        Set iDO = Nothing
        Set iFO = Nothing
        lvFiles.ListItems.Remove (lvFiles.SelectedItem.Index)
    End If
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "::ctlFileDetails_DeleteFile"
End Sub

Private Sub ExportFile()
    Dim sKey As String
    Dim iFO As CFileObject
    Dim iDO As IDataObject
    Dim sFilename As String
'
' Export the file from the database
'

'
' Get the Key from the Selected Item
'
    sKey = lvFiles.SelectedItem.Key
    sKey = Right$(sKey, Len(sKey) - 3)
'
' Ask user for Save File Name
'
    sFilename = lvFiles.SelectedItem.Text
    
    RaiseEvent RequestFileName(eSaveFileName, sFilename, "Save File As : ")
    
    If Len(sFilename) = 0 Then Exit Sub
    
    On Error Resume Next
'
' Initialise the Object
'
    Set iDO = New CFileObject
    Set iFO = iDO
    
    iDO.Initialise mDB, sKey
'
' Save it to the selected File/PathName
'
    iFO.SaveToFile sFilename
    Set iFO = Nothing
    Set iDO = Nothing
'
' Notify User that File Was Exported
'
    MsgBox "File Exported to " & sFilename, vbOKOnly + vbInformation, "VBCodeLibrary"
    
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "::ctlFileDetails_ExportFile"
End Sub

Private Sub PopulateListView()
    
    Dim sSql As String
    Dim rs As Recordset
    Dim lCount As Long
    Dim li As ListItem
'
' Populate the listview with the associated files for this codeitem
'
    On Error Resume Next
    
'
' Remove any existing items from the listview
'
    lvFiles.ListItems.Clear
'
' Build the SQL statement
'
    sSql = "select * from codefiles where codeid = " & mDataObject.Key
    
    Set rs = mDB.OpenRecordset(sSql)
    
    If rs.BOF And rs.EOF Then
        Exit Sub
    End If
'
' Make sure we have all the items from the cursor
'
    rs.MoveFirst
    rs.MoveLast
    rs.MoveFirst
        
    For lCount = 1 To rs.RecordCount
        Set li = lvFiles.ListItems.Add(, "ID=" & (rs.Fields("id").Value), rs.Fields("description").Value)
        li.SubItems(1) = Format$(rs.Fields("origdatetime").Value)
        li.SubItems(2) = Format$(rs.Fields("dateadded").Value)
        li.SubItems(3) = Format$(rs.Fields("file").FieldSize, "#,###,###")
        rs.MoveNext
    Next
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "::ctlFileDetails_PopulateListView"
    
End Sub

Private Sub AddCodeFile(ByVal sFilename As String)
'
    Dim iFO As CFileObject
    Dim iDO As IDataObject
'
' Create a new CFileObject in the database with the chosen file
'
    On Error GoTo vbErrorHandler
    
    Set iDO = New CFileObject
    Set iFO = iDO
    
    iDO.Initialise mDB
'
' Set the Parent Object ID on the CFileObject Object
'
    iFO.CodeID = mDataObject.Key
'
' Set the Stored File Property
'
    iFO.StoredFile = sFilename
'
' Write it away to the database
'
    iDO.Commit
    
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & "::ctlFileDetails_AddCodeFile"

End Sub
