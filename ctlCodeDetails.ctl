VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl ctlCodeDetails 
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4305
   ScaleWidth      =   4620
   Begin VB.TextBox txtCode 
      Height          =   2565
      Left            =   60
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "ctlCodeDetails.ctx":0000
      Top             =   960
      Width           =   4350
   End
   Begin VBCodeLib.ctlFileDetails ctlFileDetails1 
      Height          =   1035
      Left            =   300
      TabIndex        =   2
      Top             =   1695
      Visible         =   0   'False
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   1826
   End
   Begin ComctlLib.TabStrip tbsTabs 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      Top             =   495
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   5477
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Stored Code"
            Key             =   "CODE"
            Object.Tag             =   "CODE"
            Object.ToolTipText     =   "Show/Add Code"
            ImageVarType    =   8
            ImageKey        =   "CODE"
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Notes"
            Key             =   "Notes"
            Object.Tag             =   "Notes"
            Object.ToolTipText     =   "Show / Add Notes"
            ImageVarType    =   8
            ImageKey        =   "NOTES"
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Usage"
            Key             =   "Usage"
            Object.Tag             =   "Usage"
            Object.ToolTipText     =   "Show / Add Usage Instructions"
            ImageVarType    =   8
            ImageKey        =   "USAGE"
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fil&es"
            Key             =   "FILES"
            Object.Tag             =   "FILES"
            Object.ToolTipText     =   "Show / Add / Export Files for this Code Item"
            ImageVarType    =   8
            ImageKey        =   "FILES"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgLabel 
      Height          =   330
      Left            =   4110
      Stretch         =   -1  'True
      Top             =   360
      Width           =   360
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4035
      Top             =   2895
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ctlCodeDetails.ctx":0006
            Key             =   "USAGE"
            Object.Tag             =   "USAGE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ctlCodeDetails.ctx":0320
            Key             =   "NOTES"
            Object.Tag             =   "NOTES"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ctlCodeDetails.ctx":063A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ctlCodeDetails.ctx":074C
            Key             =   "CODE"
            Object.Tag             =   "CODE"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ctlCodeDetails.ctx":0A66
            Key             =   "FILES"
            Object.Tag             =   "FILES"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Caption         =   "Label Caption Goes Here ...."
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
      Height          =   330
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "ctlCodeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Control for display data from the CodeItem object
'
' Chris Eastwood Jan/Feb 1999
'

'
' Our constants - used in trapping on KeyDown / Key Press routines
'
Public Enum eCurView
    vwCode = 1
    vwNotes = 2
    vwUsage = 3
    vwFiles = 4
End Enum


Private Const kpBold = 2            ' CTRL & B  - Bold
Private Const kpItalic = 9          ' CTRL & I  - Italic
Private Const kpUnderline = 21      ' CTRL & U  - Underline

Private meCurrentView As eCurView

Private mCodeObject As CCodeItem        ' Internal Pointer to our CodeItem
Private mDataObject As IDataObject      ' Pointer to CodeItem's IDataObject Interface
Private mDB As Database

Public Event ViewChanged(ByVal CurrentView As eCurView)

Public Event RequestFileName(ByVal DialogType As eGetFileDialog, ByRef sFilename As String, ByVal sDialogTitle As String)

Public Sub Initialise(oDB As Database, oCodeObject As IDataObject)
'
' Entry Point
'
    Dim tTab As Object
    
    Set mDB = oDB
    
    If oCodeObject Is Nothing Or oDB Is Nothing Then
'
' User most likely chose 'root'
'
        DisplayDefaults
    
    Else
'
' Record object internally & populate control
'
        Set mCodeObject = oCodeObject
        Set mDataObject = oCodeObject
        
        PopulateControl
        ctlFileDetails1.Initialise mDB, mCodeObject
        tbsTabs.Visible = True
        tbsTabs.Enabled = True
        txtCode.Enabled = True
    End If
    UserControl_Resize
    
End Sub

Public Sub Terminate()
'
' Release references to any data objects
'
    Set mDataObject = Nothing
    Set mCodeObject = Nothing
    Set mDB = Nothing
End Sub

Private Sub DisplayDefaults()
'
' Setup Defaults for Display when user chooses 'ROOT' item
'
    txtCode.Text = ""
    Label1.Caption = "Welcome to the VB Code Library"
    tbsTabs.Visible = False
    tbsTabs.Enabled = False
    txtCode.Enabled = False
    ctlFileDetails1.Visible = False
End Sub

Private Sub PopulateControl()
'
' Populate the control depending on the selected tab
'
    Select Case UCase$(tbsTabs.SelectedItem.Key)
        Case "CODE"
            DisplayCode
        Case "NOTES"
            DisplayNotes
        Case "USAGE"
            DisplayUsage
        Case "FILES"
            DisplayFiles
            
    End Select
    
    With mCodeObject
        Label1.Caption = " " & .Description
    End With
    
End Sub

Private Sub DisplayCode()
    txtCode.Text = mCodeObject.Code
    Set imgLabel.Picture = ImageList1.ListImages("CODE").Picture
    RaiseEvent ViewChanged(vwCode)
End Sub

Private Sub DisplayNotes()
    txtCode.Text = mCodeObject.Notes
    Set imgLabel.Picture = ImageList1.ListImages("NOTES").Picture
    RaiseEvent ViewChanged(vwNotes)
    
End Sub

Private Sub DisplayUsage()
    txtCode.Text = mCodeObject.Example
    Set imgLabel.Picture = ImageList1.ListImages("USAGE").Picture
    RaiseEvent ViewChanged(vwUsage)
    
End Sub

Private Sub DisplayFiles()
    imgLabel.Picture = ImageList1.ListImages("FILES").Picture
    ctlFileDetails1.Initialise mDB, mCodeObject
    ctlFileDetails1.Visible = True
    ctlFileDetails1.ZOrder
    RaiseEvent ViewChanged(vwFiles)
End Sub

Private Sub ctlFileDetails1_RequestFileName(ByVal DialogType As eGetFileDialog, sFilename As String, ByVal sDialogTitle As String)
    RaiseEvent RequestFileName(DialogType, sFilename, sDialogTitle)
End Sub

Private Sub tbsTabs_BeforeClick(Cancel As Integer)
'
' Here is where we would normally copy text back into
' our DataObject and write away if required
'
' However, when you set the TabStrip to buttons with a
' flat style. This BeforeClick event doesnt seem to
' get fired with NT4 at least (will try it under 95/98
' at a later date).
'
' I'm leaving this code in to show how it should work
' If it does get called - great !
'
'
    Select Case UCase$(tbsTabs.SelectedItem.Key)
        Case "CODE"
            mCodeObject.Code = txtCode.Text
        Case "NOTES"
            mCodeObject.Notes = txtCode.Text
        Case "USAGE"
            mCodeObject.Example = txtCode.Text
'
' The 'File' stuff is all handled by the control
'
    End Select
'
' Commit the changes
'
    mDataObject.Commit
    
End Sub

Private Sub tbsTabs_Click()
'
' Display Required Tab
'
    Static sLastTabKey As String
    Dim sKey As String
    
'
' The following code is a work around for the BeforeClick bug mentioned above
'
    sKey = tbsTabs.SelectedItem.Key
    
    If StrComp(sKey, sLastTabKey, vbTextCompare) = 0 Then
        Exit Sub
    Else
        Select Case UCase$(sLastTabKey)
            Case "CODE", "" ' coz sLastTabKey will be "" the first time around !
                mCodeObject.Code = txtCode.Text
            Case "NOTES"
                mCodeObject.Notes = txtCode.Text
            Case "USAGE"
                mCodeObject.Example = txtCode.Text
            Case "FILES"
                ctlFileDetails1.Terminate
        End Select
        mDataObject.Commit
    End If
    
    sLastTabKey = sKey
    
'
' Resume Normal Play !
'
    txtCode.Enabled = False
    ctlFileDetails1.Visible = False
    
    Select Case UCase$(sKey)
        Case "CODE"
            DisplayCode
        Case "NOTES"
            DisplayNotes
        Case "USAGE"
            DisplayUsage
        Case "FILES"
            DisplayFiles
    End Select
    txtCode.Enabled = True
    
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
'
' First, allow tabs in our text box
'
    
    If KeyCode = Asc(vbTab) Then
        txtCode.SelText = vbTab
        KeyCode = 0
        Exit Sub
    End If
End Sub

Private Sub txtCode_LostFocus()
'
' Record changes if Focus is lost (just in case)
'
    On Error Resume Next ' because we could have opened a new database
    
    Select Case UCase$(tbsTabs.SelectedItem.Key)
        Case "CODE"
            mCodeObject.Code = txtCode.Text
        Case "NOTES"
            mCodeObject.Notes = txtCode.Text
        Case "USAGE"
            mCodeObject.Example = txtCode.Text
    End Select
    
    mDataObject.Commit
    
End Sub

'
Private Sub UserControl_Initialize()
    Dim lRet As Long
    Dim lStyle As Long
    
    lStyle = GetWindowLong(tbsTabs.hwnd, GWL_STYLE)
    lStyle = lStyle Or TCS_FLATBUTTONS
    SetWindowLong tbsTabs.hwnd, GWL_STYLE, lStyle
    
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Static bInHere As Boolean
'
' Make sure we resize our constituent controls correctly
'
    If bInHere Then Exit Sub ' to stop potential stack problems
    
    bInHere = True
    
    Label1.Move ScaleLeft + 15, ScaleTop + 15, ScaleWidth - 30, Label1.Height
    imgLabel.Move UserControl.ScaleWidth - imgLabel.Width, ScaleTop + 15, imgLabel.Width, Label1.Height - 2
    
    With tbsTabs
        .Move UserControl.ScaleLeft, UserControl.ScaleTop + Label1.Height + 65, UserControl.ScaleWidth, UserControl.ScaleHeight - (Label1.Height + 15)
        If .Visible Then
            txtCode.Move UserControl.ScaleLeft, .ClientTop + 30, UserControl.ScaleWidth, .ClientHeight - 30
        Else
            txtCode.Move UserControl.ScaleLeft, UserControl.ScaleTop + Label1.Height + 65, UserControl.ScaleWidth, UserControl.ScaleHeight - (Label1.Height + 75)
        End If
        
        ctlFileDetails1.Move UserControl.ScaleLeft, .ClientTop + 30, UserControl.ScaleWidth, .ClientHeight - 30
    End With
    bInHere = False
    
End Sub

Public Property Get CodeWindowText() As String
    CodeWindowText = txtCode.Text
End Property

Public Property Get CurrentView() As eCurView
    CurrentView = meCurrentView
End Property
