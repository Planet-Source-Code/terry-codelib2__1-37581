VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VBCodeLibrary Options"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3585
      TabIndex        =   2
      Top             =   1890
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2445
      TabIndex        =   1
      Top             =   1890
      Width           =   1050
   End
   Begin VB.ListBox lstOptions 
      Height          =   1410
      Left            =   45
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   390
      Width           =   4605
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Options :"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   3465
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Options Form
'
' http://www.codeguru.com/vb
'
'
' Chris Eastwood Feb. 1999
'

Private moSettings As CSettings   ' Settings object

Public Sub Initialise(oSettings As CSettings)
'
' Entry point
' - Record Settings object
'
    Set moSettings = oSettings
'
' populate list box with available options
'
    With lstOptions
        .Clear
        .AddItem "Backup Database On Startup"
        .Selected(.NewIndex) = oSettings.BackupDatabaseAtStart
        .AddItem "Compact Database On Exit"
        .Selected(.NewIndex) = oSettings.CompactDatabaseOnExit
        .AddItem "Save Form Layout on Exit"
        .Selected(.NewIndex) = oSettings.SaveFormLayout
        .ListIndex = 0
    End With
    
End Sub

Private Sub cmdCancel_Click()
'
' User Cancelled
'
    Set moSettings = Nothing
'
' Hide the form
'
    Me.Hide
End Sub

Private Sub cmdOk_Click()
'
' Change Settings Options
'
    With lstOptions
        moSettings.BackupDatabaseAtStart = .Selected(0)
        moSettings.CompactDatabaseOnExit = .Selected(1)
        moSettings.SaveFormLayout = .Selected(2)
    End With
    Set moSettings = Nothing
'
' Hide the form
'
    Me.Hide
    
End Sub

