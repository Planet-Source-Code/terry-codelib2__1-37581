Attribute VB_Name = "CodeModule"
Option Explicit
'
' Global Constants and Declarations for the VBCodeLibrary Project
'
' http://www.codeguru.com/vb
'
'
' Chris Eastwood Feb. 1998
'

'
' Our Application Generated Errors
'
Public Enum AppErrors
    errAwaitingDelete = vbObjectError + 513
    errObjectDeleted
    errObjectNotCreated
End Enum

Public Enum eGetFileDialog
    eOpenFileName           ' Used in Generic Routines to Get File Names
    eSaveFileName
End Enum

'
' Our Exported / Imported Data Type
'
Public Type FileHeader
    lNumberOfRecords As Long
End Type
'
Public Type ImportData
    sName As String
    sOriginalID As String
    sParentID As String
    sNewID As String
    sParentName As String
    sStoredCode As String
    sNotes As String
    sUsage As String
End Type

'
' Win API Types
'
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long

'
' API Types

'
' API Messages
'
Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)

'
' ListView Types/Messages/Styles
'

Public Const LVSCW_AUTOSIZE As Long = -1
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 55
Public Const LVM_GETCOLUMNWIDTH As Long = LVM_FIRST + 29
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30

Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_TRACKSELECT = &H8
Public Const LVS_EX_FLATSB  As Long = &H100
'
' Misc Windows Messages and Styles
'
Public Const SM_CXVSCROLL As Long = 2 ' Get Width Of Vertical ScrollBar
Public Const WS_HSCROLL As Long = &H100000
Public Const HDS_BUTTONS As Long = &H2
Public Const GWL_STYLE As Long = (-16)
'Public Const SWP_DRAWFRAME As Long = &H20
'Public Const SWP_NOMOVE As Long = &H2
'Public Const SWP_NOSIZE As Long = &H1
'Public Const SWP_NOZORDER As Long = &H4
'Public Const SWP_FLAGS As Long = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Public Const WM_SETREDRAW As Long = &HB
Public Const SW_SHOWNORMAL As Long = 1

'
' Toolbar State Messages
'
Public Const TB_SETSTYLE As Long = WM_USER + 56
Public Const TB_GETSTYLE As Long = WM_USER + 57
Public Const TBSTYLE_FLAT As Long = &H800

'
' System Tray Messages and Structures
'
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD As Long = &H0
Public Const NIM_DELETE As Long = &H2
Public Const WM_MOUSEMOVE As Long = &H200
Public Const NIF_MESSAGE As Long = &H1
Public Const NIF_ICON As Long = &H2
Public Const NIF_TIP As Long = &H4
'
' Mouse Messages Captured from the System Tray
'
Public Const WM_LBUTTONDBLCLK As Long = &H203

'
' Treeview Messages and styles
'
Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_GETEDITCONTROL As Long = (TV_FIRST + 15)
Public Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Public Const TVM_GETITEM As Long = (TV_FIRST + 12)
Public Const TVM_SETITEM As Long = (TV_FIRST + 13)
Public Const TVM_SELECTITEM As Long = (TV_FIRST + 11)
'
Public Const TVIF_STATE As Long = &H8
Public Const TVS_TRACKSELECT As Long = &H200&
Public Const TVS_FULLROWSELECT As Long = &H1000
Public Const TVIS_BOLD As Long = &H10
'
Public Const TVGN_ROOT As Long = &H0
Public Const TVGN_NEXT As Long = &H1
Public Const TVGN_CARET As Long = &H9
Public Const EM_LIMITTEXT = &HC5
Public Const WM_VSCROLL = &H115

'
' Treeview Item Structure
'
Public Type TVITEM
   mask As Long
   hItem As Long
   State As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   cChildren As Long
   lParam As Long
End Type

'
' WinAPI Declarations
'
Public Const TCS_FLATBUTTONS = &H8
Public Const GWL_EXSTYLE = (-20)

'
' Declarations for SHGETFILEINFO & associated routines
' - from a posting by Brad Martinez
'
Public Const MAX_PATH = 260

Public Type SHFILEINFO   ' shfi
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

'
' ShellGetFileInfo Flags Enum stolen from the Net - Brad Martinez I think ?
'
Public Enum SHGFI_FLAGS
    SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
    SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
    SHGFI_OPENICON = &H2              ' sfi.hIcon is open icon
    SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
    SHGFI_PIDL = &H8                        ' pszPath is pidl, rtns BOOL
    SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
    SHGFI_ICON = &H100                    ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
    SHGFI_DISPLAYNAME = &H200    ' isf.szDisplayName is filled, rtns BOOL
    SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
    SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
    SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
                                ' containing the icon, rtns BOOL
    SHGFI_EXETYPE = &H2000            ' rtns two ASCII chars of exe type
    SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
    SHGFI_LINKOVERLAY = &H8000    ' add shortcut overlay to sfi.hIcon
    SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
End Enum

Public Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbFileInfo As Long, _
    ByVal uFlags As SHGFI_FLAGS) As Long

'
' Declares for other WINAPI Stuff
'
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Integer
Public Declare Function GetTempFileName Lib "KERNEL32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "KERNEL32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Public Sub Main()
'
' Ensure only one instance is running
'
    If App.PrevInstance Then
        MsgBox "An Instance of the VBCodeLibrary Tool is already Running", , App.ProductName
        Exit Sub
    End If
    
    frmCodeLib.Show

End Sub


Public Sub ReplaceAll(ByRef sOrigStr As String, ByVal sFindStr As String, ByVal sReplaceWithStr As String, Optional bWholeWordsOnly As Boolean)
'
' Replaces all occurances of sFindStr with sReplaceWithStr
' (as included with this project database!)

    Dim lPos As Long
    Dim lPos2 As Long
    Dim sTmpStr As String
    Dim bReplaceIt As Boolean
    Dim lFindStr As Long
    
    On Error GoTo vbErrorHandler
    
    lFindStr = Len(sFindStr)
    
    lPos2 = 1
    bReplaceIt = True
    sTmpStr = sOrigStr
    
    Do
        lPos = InStr(lPos2, sOrigStr, sFindStr)
        If lPos = 0 Then
            Exit Do
        End If
        If bWholeWordsOnly Then
            On Error Resume Next
            If lPos = 1 Or (Mid$(sOrigStr, lPos - 1, 1) = " ") Then
                If (Mid$(sOrigStr, lPos + lFindStr, 1) = " ") Or Mid$(sOrigStr, lPos + lFindStr + 1, 1) = "" Then
                    bReplaceIt = True
                Else
                    bReplaceIt = False
                End If
            End If
        End If
        If bReplaceIt Then
            If lPos > 1 Then
                sTmpStr = Left$(sOrigStr, lPos - 1)
            Else
                sTmpStr = ""
            End If
            sTmpStr = sTmpStr & sReplaceWithStr
            sTmpStr = sTmpStr & Mid$(sOrigStr, lPos + lFindStr, Len(sOrigStr) - (lPos + lFindStr - 1))
            sOrigStr = sTmpStr
        End If
        lPos2 = lPos + 1
    Loop
    sOrigStr = sTmpStr
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description, , "CodeModule::ReplaceAll"

    
End Sub

Public Sub AutoSizeListViewColumns(lvListView As ListView, Optional bAutoSizeLastColumn As Boolean = False)
    Dim lCount As Long
'
' Turn off Redrawing at this point to speed up / hide the visible changes
'
    
    SendMessageLong lvListView.hwnd, WM_SETREDRAW, False, &O0
    
    For lCount = 0 To lvListView.ColumnHeaders.Count - 1
        Call SendMessageLong(lvListView.hwnd, LVM_SETCOLUMNWIDTH, lCount, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next
'
' Turn Redrawing back on
'
    SendMessageLong lvListView.hwnd, WM_SETREDRAW, True, &O0
    
    If bAutoSizeLastColumn Then
        AutoSizeLastColumn lvListView
    End If
    
End Sub

Public Sub AutoSizeLastColumn(lvListView As ListView)
    Dim lCount As Long
    Dim lNoColumns As Long
    Dim lTotSize As Long
    Dim lRet As Long
    Dim lSize As Long
    Dim lHScrollBarWidth As Long

On Error GoTo vbErrorHandler

'
' Get Number of columns in this listview
'
    lNoColumns = lvListView.ColumnHeaders.Count
'
' Get ScrollBar Width
'
    lHScrollBarWidth = GetSystemMetrics(SM_CXVSCROLL)

    For lCount = 0 To lNoColumns - 2
'
' Get the total size of all the columns except the last one we want to resize
'
        lSize = SendMessageLong(lvListView.hwnd, LVM_GETCOLUMNWIDTH, lCount, 0)
        lTotSize = lTotSize + lSize
    Next
'
' Now determine how big to make the last columm in pixels
'

    lSize = (lvListView.Width / Screen.TwipsPerPixelX) - (lTotSize + lHScrollBarWidth + 10)
'
' Now set the column width
'
    SendMessageLong lvListView.hwnd, LVM_SETCOLUMNWIDTH, lNoColumns - 1, lSize

    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source, , "Common::AutoSizeLastColumn"

End Sub

