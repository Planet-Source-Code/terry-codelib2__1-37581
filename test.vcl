	    Miscellaneous 274 0   ROOT      
 Status Bar 237 274   C274      $ Using A Progress Bar in a Status Bar 238 237   C237�1. Create statusbar with few panels
2. Create invisible progressbar
3. Add test button which calls ShowProgress sub shown below

'Module Declares
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long,
ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA"
(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As
Any) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Const WM_USER = &H400
Public Const SB_GETRECT = (WM_USER + 10)


Private Sub ShowProgress(Mode As Boolean)

Dim rc As RECT

    StatusBar1.Panels("keyProgress").Visible = Mode
    
    If Mode Then
'2 => Panel index (0 based)
        SendMessageAny StatusBar1.hwnd, SB_GETRECT, 2, rc
    
        With rc
            .Top = .Top * Screen.TwipsPerPixelY
            .Left = .Left * Screen.TwipsPerPixelX
            .Bottom = .Bottom * Screen.TwipsPerPixelY - .Top
            .Right = .Right * Screen.TwipsPerPixelX - .Left
        End With
    
        With ProgressBar1
            SetParent .hwnd, StatusBar1.hwnd
            .Move rc.Left, rc.Top, rc.Right, rc.Bottom
            .Visible = True
            .Value = 0
        End With
    Else
        SetParent ProgressBar1.hwnd, Me.hwnd
        ProgressBar1.Visible = False
    End If
    
End Sub     Hiding Applications 239 274   C274       Hide App From Task Manager (NT) 241 239   C239r 'Hide the app from the task manager.
lRtn = GetWindow(frmHide.hwnd, GW_OWNER)
lRtn = ShowWindow(lRtn, SW_HIDE)
    % Hide App From Task Manager (win95/98) 240 239   C239
Declare Function RegisterServiceProcess lib "kernel32" (byval ProcessID
as long, byval ServiceFlags as Long) as long
Declare Function GetCurrentProcessId lib "kernel32" () as Long

' ServiceFlags: 0 means unregister as a service, 1 means to register.

ret = RegisterServiceProcess (GetCurrentProcessId, 1)

'When you register the service process it becomes hidden.

ret = RegisterServiceProcess (GetCurrentProcessId, 0)

'And the unhidden when you unregister it when you're finished with hiding it.
     Reading Items in InBox 221 274   C2744[1]    Add the reference to the Outlook Object Library
 
Dim myOLApp As New Outlook.Application
Dim olNameSpace As Outlook.NameSpace
Dim myItem As New Outlook.AppointmentItem
Dim myRequest As New Outlook.MailItem
Dim myFolder As Outlook.MAPIFolder

Public myResponse
Dim L As String
Dim i As Integer
Dim SearchSub As String
Dim strSubject As String
Dim myFolder As Outlook.MAPIFolder
Dim strSender As String
Dim strBody As String
Dim olMapi As Object
Dim strOwnerBox As String
Dim sbOLApp
    
    
    Set myOLApp = CreateObject("Outlook.Application")
    Set olNameSpace = myOLApp.GetNamespace("MAPI")
    Set myFolder = olNameSpace.GetDefaultFolder(olFolderInbox)
    
    
    'Dim mailfolder As Outlook.MAPIFolder
    
    Set olMapi = GetObject("", "Outlook.Application").GetNamespace("MAPI")
    
    For i = 1 To myFolder.Items.Count
            strSubject = myFolder.Items(i).Subject
            strBody = myFolder.Items(i).Body
            strSender = myFolder.Items(i).SenderName
            strOwnerBox = myFolder.Items(i).ReceivedByName

 
            ' Now Mail it to somebody
            Set sbOLAPp = CreateObject("Outlook.Application")
            Set myRequest = myOLApp.CreateItem(olMailItem)
            With myRequest
                .Subject = strSubject
                .Body = strBody
                .To = "anybody@anywhere.com"
                .Send
                
                
            End With
            Set sbOLAPp = Nothing
            
    Next
    
    Set myOLApp = Nothing
    Exit Sub
 
     Moving Controls at RunTime 251 274   C274'
' A Really Neat little bit of code !
'
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA"
(ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As
Single, Y As Single)
    
    ReleaseCapture
    SendMessage Picture1.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0

End Sub     ListView Coloured Items 275 274   C274Option Explicit
'
' Custom Draw Listview Processing
'
' This form will sublcass it's own message queue and listen out for any WM_NOTIFY
' messages from the Listview. When one is found, we tell the listview to notify us on
' every item paint.
'
' Every second item in the listview should be painted in red
'
'
' Chris Eastwood May 1998
'
' Feel free to use any of this code as you wish
'

'
' Custom Draw Message to intercept
'
Public Enum WinNotifications
    NM_FIRST = -0& ' (0U- 0U) ' // generic to all controls
    NM_LAST = -99& ' (0U- 99U)
    NM_OUTOFMEMORY = (NM_FIRST - 1)
    NM_CLICK = (NM_FIRST - 2)
    NM_DBLCLK = (NM_FIRST - 3)
    NM_RETURN = (NM_FIRST - 4)
    NM_RCLICK = (NM_FIRST - 5)
    NM_RDBLCLK = (NM_FIRST - 6)
    NM_SETFOCUS = (NM_FIRST - 7)
    NM_KILLFOCUS = (NM_FIRST - 8)
    NM_CUSTOMDRAW = (NM_FIRST - 12)
    NM_HOVER = (NM_FIRST - 13)
End Enum
'
' Win API Rect structure
'
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


'
' Custom Draw Structures
'
' The NMHDR structure contains information about a notification message. The pointer
' to this structure is specified as the lParam member of the WM_NOTIFY message.
'
Private Type NMHDR
    hwndFrom As Long        ' Window handle of control sending message
    idFrom As Long               ' Identifier of control sending message
    code As Long                  ' Specifies the notification code
End Type

Private Type NMCUSTOMDRAWINFO
    hdr As NMHDR
    dwDrawStage As Long
    hdc As Long
    rc As RECT
    dwItemSpec As Long
    iItemState As Long
    lItemLParam As Long
End Type

Private Type NMLVCUSTOMDRAW
    nmcmd As NMCUSTOMDRAWINFO
    clrText As Long
    clrTextBk As Long
    iSubItem As Integer
End Type
'
' Notify Message
'
Private Const WM_NOTIFY& = &H4E
'
' Custom Draw Messages
'
Private Const CDDS_PREPAINT& = &H1
Private Const CDDS_POSTPAINT& = &H2
Private Const CDDS_PREERASE& = &H3
Private Const CDDS_POSTERASE& = &H4
Private Const CDDS_ITEM& = &H10000
Private Const CDDS_ITEMPREPAINT& = CDDS_ITEM Or CDDS_PREPAINT
Private Const CDDS_ITEMPOSTPAINT& = CDDS_ITEM Or CDDS_POSTPAINT
Private Const CDDS_ITEMPREERASE& = CDDS_ITEM Or CDDS_PREERASE
Private Const CDDS_ITEMPOSTERASE& = CDDS_ITEM Or CDDS_POSTERASE
Private Const CDDS_SUBITEM& = &H20000

Private Const CDRF_DODEFAULT& = &H0
Private Const CDRF_NEWFONT& = &H2
Private Const CDRF_SKIPDEFAULT& = &H4
Private Const CDRF_NOTIFYPOSTPAINT& = &H10
Private Const CDRF_NOTIFYITEMDRAW& = &H20
Private Const CDRF_NOTIFYSUBITEMDRAW = &H20     ' flags are the same, we can distinguish by context
Private Const CDRF_NOTIFYPOSTERASE& = &H40
Private Const CDRF_NOTIFYITEMERASE& = &H80
'
' Win API Declarations
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_GETFONT = &H31


Private Sub Form_Load()
    Dim li As ListItem
    Dim lCount As Long
'
' Setup the listview with 1 column and 25 listitems
'

    With ListView1.ListItems
        ListView1.ColumnHeaders.Add 1, "Test", "Column 1"

        For lCount = 1 To 25
            Set li = .Add(lCount, "key" & lCount, "This is line " & lCount)
        Next
    End With
'
' Now subclass the form and watch for WM_NOTIFY messages coming from the listview
'
' I'm using the Softcircuits subclass control here, although you can use any other
' (they all work in pretty much the same way). You could also do any AddressOf processing
' if you so wish. I just use the SC control because it's a lot quicker (and you don't have
' to remember to un-subclass your window afterwards)
'
    With Subclass1
        .hwnd = Me.hwnd
        .Messages(WM_NOTIFY) = True
    End With

End Sub

Private Sub Subclass1_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)

    Dim tMessage As NMHDR
    Dim lCode As Long
    Dim tLVRedrawMessage As NMLVCUSTOMDRAW

    Select Case Msg
'
' Should only be WM_NOTIFY (that's all we've subclassed)
'
        Case WM_NOTIFY
'
' Same as in C : tMessage = (NMHDR) lParam;
'
'
' The .code section of the NMHDR notify structure contains the submessage
'
            CopyMemory tMessage, ByVal lParam, Len(tMessage)
            lCode = tMessage.code

            Select Case lCode
                Case NM_CUSTOMDRAW
'
' Make sure it's our listview raising the Custom Redraw message
'
                    If tMessage.hwndFrom <> ListView1.hwnd Then
'
' It's not ! - Return default processing to windows
'
                        Result = Subclass1.CallWndProc(Msg, wParam, lParam)
                        Exit Sub
                    End If
'
' Copy the message into our local structure
'
                    CopyMemory tLVRedrawMessage, ByVal lParam, Len(tLVRedrawMessage)
'
' Now process the Custom Redraw Messages in Order :
'
' CDDS_PREPAINT is at the beginning of the paint cycle.
' You must return the property value to get Custom painting
' to work correctly. In this example, we're only looking for
' item specific painting - although theoretically, you should
' be able to paint just about anything on the control, from
' bitmap backgrounds to changing fonts etc.
'
' (Just don't ask me how to do it (yet)).
'
                    If tLVRedrawMessage.nmcmd.dwDrawStage = CDDS_PREPAINT Then
'
' Request a notification for each item being painted
'
                        Result = CDRF_NOTIFYITEMDRAW
                        Exit Sub
                    End If
'
' Because we returned CDRF_NOTIFYITEMDRAW in the above code, CDDS_ITEMPREPAINT is now sent
' when the control is ready to paint an Item
'
                    If tLVRedrawMessage.nmcmd.dwDrawStage = CDDS_ITEMPREPAINT Then
'
' The item's about to be repainted - Here's where you can trap to see which item is being
' painted and so set the color accordingly
'
' To see which item is about to be painted, check :
'
' if tLVRedrawMessage.nmcm.dwItemSpec = required listview item number Then
'
' To Change the text and background colours in a list view control,
' set the clrText and clrTextBk members of the NMLVCUSTOMDRAW structure to the
' required color. Most other controls rely on the SetTextColor and SetBkColor API
' calls on the passed in hdc
'
' In this code I'm setting every second listitem to be red
'
'
                        With tLVRedrawMessage
                            If .nmcmd.dwItemSpec / 2 = CInt(.nmcmd.dwItemSpec / 2) Then
                                .clrTextBk = vbWhite
                                .clrText = vbRed
'
' You must remember to copy back the changes made in tLVRedrawMessage to the LPARAM value
'
                                CopyMemory ByVal lParam, tLVRedrawMessage, Len(tLVRedrawMessage)
                                Exit Sub
                            Else
'
' This is standard painting stuff - let windows do it for us
'
                                Result = CDRF_DODEFAULT
                                Exit Sub
                            End If
                        End With
                    End If

                Case Else
'
' Other messages from the listview which we're not interested in should be passed back
'
                    Result = Subclass1.CallWndProc(Msg, wParam, lParam)
                    Exit Sub
        End Select
    End Select
End Sub
    