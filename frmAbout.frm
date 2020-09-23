VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VBCodeLibrary"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4905
      TabIndex        =   5
      Top             =   1065
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   -15
      Picture         =   "frmAbout.frx":030A
      Top             =   -60
      Width           =   2490
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   4380
      TabIndex        =   4
      Top             =   1425
      Width           =   1755
   End
   Begin VB.Label lblLink 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.codeguru.com/vb"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   60
      MouseIcon       =   "frmAbout.frx":09F4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1155
      Width           =   2235
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CodeGuru - The WebSite for Developers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   885
      Width           =   4320
   End
   Begin VB.Label lblMailTo 
      BackStyle       =   0  'Transparent
      Caption         =   "Chris Eastwood Jan/Feb 1999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   2640
      MouseIcon       =   "frmAbout.frx":0CFE
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   570
      Width           =   2625
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VBCodeLibrary Tool"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   2625
      TabIndex        =   0
      Top             =   225
      Width           =   1980
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' About Form for the VBCodeLibrary project
'
' http://www.codeguru.com/vb
'
' Chris Eastwood Feb. 1999
'
' - Updated April 1999 - new CodeGuru Bitmap
'

Private Sub ExecuteLink(ByVal sLinkTo As String)
'
' Execute the link to http://www.codeguru.com/vb
' (if possible) - or the new 'mailto:ME!'
'
    On Error Resume Next
    
    Dim lRet As Long
    Dim lOldCursor As Long
    
    lOldCursor = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    lRet = ShellExecute(0, "open", sLinkTo, "", vbNull, SW_SHOWNORMAL)
    
    If lRet >= 0 And lRet <= 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Error Opening Link to " & sLinkTo & vbCrLf & vbCrLf & Err.LastDllError, , "frmAbout::ExecuteLink"
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub


Private Sub lblLink_Click()
    ExecuteLink lblLink.Caption
    Unload Me
End Sub

Private Sub lblMailTo_Click()
    ExecuteLink "mailto:chris.eastwood@codeguru.com"
    Unload Me
End Sub
