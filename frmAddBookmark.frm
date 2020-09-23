VERSION 5.00
Begin VB.Form frmAddBookmark 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Bookmark"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddBookmark.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   3480
      TabIndex        =   5
      Top             =   1980
      Width           =   1140
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   360
      Left            =   2160
      TabIndex        =   4
      Top             =   1980
      Width           =   1140
   End
   Begin VB.TextBox txtBookmark 
      Height          =   1140
      Left            =   45
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   750
      Width           =   4560
   End
   Begin VB.Label lblBookmark 
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmark Description :"
      Height          =   225
      Left            =   60
      TabIndex        =   3
      Top             =   465
      Width           =   1935
   End
   Begin VB.Label lblCodeItem 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1635
      TabIndex        =   1
      Top             =   30
      Width           =   2970
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Adding bookmark to :"
      Height          =   270
      Left            =   30
      TabIndex        =   0
      Top             =   75
      Width           =   1680
   End
End
Attribute VB_Name = "frmAddBookmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' frmAddBookmark
'
' Add's a bookmark to the database for a relevant codeitem
'
' http://www.codeguru.com/vb
'
' Chris Eastwood Feb. 1999
'

'
' Private Members
'
Private moCodeItem As CCodeItem ' Pointer to the CodeItem Data Object
Private mbCancelled As Boolean  ' Was the Form Cancelled
Private mDB As Database         ' Reference to the Database


Public Property Get Cancelled() As Boolean
'
' Was the Form's Cancel Button Selected ?
'
    Cancelled = mbCancelled
End Property

Public Sub Initialise(oDB As Database, oCodeItem As IDataObject)
'
' Entry Point - Record Database & Code Item Object
'
    Set moCodeItem = oCodeItem
    Set mDB = oDB
'
' Set Caption to that of the CodeItem Name/Description
'
    lblCodeItem.Caption = " " & moCodeItem.Description
    mbCancelled = False
    
End Sub

Private Sub cmdCancel_Click()
'
' User Clicked Cancel - Set Flag and hide form
'
    Set moCodeItem = Nothing
    mbCancelled = True
    Me.Hide
End Sub

Private Sub cmdOk_Click()

On Error GoTo vbErrorHandler

'
' User Clicked OK - Add the Book Mark to the Database
'
    Dim oBookMark As CBookmark
    Dim iDO As IDataObject
    
'
' Create New Bookmark Object
'
    Set iDO = New CBookmark
    Set oBookMark = iDO
    
    iDO.Initialise mDB
'
' Point Bookmark to Code item
'
    Set oBookMark.CodeItem = moCodeItem
    oBookMark.Description = txtBookmark.Text
'
' Write it to the Database
'
    iDO.Commit
    
    Set moCodeItem = Nothing
'
' Hide the Form
'
    Me.Hide
    

    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source, , "frmAddBookmark::cmdOk_Click"

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Kill Database Connection and any other objects
'

    Set mDB = Nothing
    Set moCodeItem = Nothing
End Sub
