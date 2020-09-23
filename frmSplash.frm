VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2265
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5655
   ControlBox      =   0   'False
   Enabled         =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Processing..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   2130
      TabIndex        =   0
      Top             =   1965
      Width           =   1155
   End
   Begin VB.Image imgLogo 
      Enabled         =   0   'False
      Height          =   1515
      Left            =   -15
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   5400
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
