VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWaitCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Wait Cursor Class
'
' Chris Eastwood January 1999
'
' Stolen from an Idea by SoftCircuits (http://www.softcircuits.com)
'
' This class encapsulates the setting of the cursor to ensure that it is
' always returned to the correct pointer, even if errors occur in the calling
' routine.
'
' Usage :
'
' Sub MySub(..)
'
'   Dim cWait As WaitCursor
'
'   Set cWait = New cWait
'
'   cWait.SetCursor
'
'   <do whatever processing>
'
' When cWait goes out of scope, the cursor is automatically returned to the
' default (or what the cursor was before cWait.SetCursor was called)

Private mlPointer As MousePointerConstants

Private Sub Class_Initialize()
   mlPointer = Screen.MousePointer
End Sub

Private Sub Class_Terminate()
    Restore
End Sub

Public Sub SetCursor(Optional nPointer As MousePointerConstants = vbHourglass)
    Screen.MousePointer = nPointer
End Sub

Public Sub Restore()
    Screen.MousePointer = mlPointer
End Sub
