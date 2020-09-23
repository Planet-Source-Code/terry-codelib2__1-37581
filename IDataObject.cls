VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Interface Definition for DataObject
'
Public Enum doState
    doStored
    doAwaitingInsert
    doAwaitingUpdate
    doAwaitingDelete
    doDeleted
End Enum

Public Property Get Key() As String
'
' Interface to get the Records Key
'
End Property

Public Sub Initialise(oDB As Database, Optional sKey As String)
'
' Interface to initialise the Data Object
'
End Sub

Public Sub Delete()
'
' Interface to mark the record as deleted
'
End Sub

Public Sub Commit()
'
' Interface to commit any changes to the dataobject to the database
'
End Sub

