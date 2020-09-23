VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBookmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDataObject

'
' Private members
'
Private msKey As String
Private mState As doState
Private mbRead As Boolean
Private msCodeKey As String
Private msBookMarkText As String

Private mDB As Database

Public Property Get CodeItemKey() As String
    GetAttributes
    CodeItemKey = msCodeKey
End Property

Public Property Let CodeItemKey(ByVal sKey As String)
    GetAttributes
    msCodeKey = sKey
    SetStateForLet
End Property

Public Property Get CodeItem() As CCodeItem
    Dim iDO As IDataObject
    
    GetAttributes
    
    If Len(msCodeKey) > 0 Then
        Set iDO = New CCodeItem
        iDO.Initialise mDB, msCodeKey
    End If
    
    Set CodeItem = iDO
    
End Property

Public Property Set CodeItem(oCodeItem As IDataObject)
    GetAttributes
    If Not (oCodeItem Is Nothing) Then
        msCodeKey = oCodeItem.Key
    Else
        msCodeKey = oCodeItem.Key
    End If
    SetStateForLet
End Property
Public Property Get Description() As String
    GetAttributes
    Description = msBookMarkText
End Property

Public Property Let Description(ByVal sBookMarkText As String)
    GetAttributes
    msBookMarkText = sBookMarkText
    SetStateForLet
End Property

Private Sub InitialiseProperties()
    mState = doStored
    msKey = ""
    msCodeKey = ""
    msBookMarkText = ""
    mbRead = False
End Sub

Private Sub GetAttributes()
    Dim mRS As Recordset
    
    On Error GoTo vbErrorHandler
    
    Select Case mState
        Case doStored, doAwaitingUpdate
            If Not mbRead Then
                'get details
                Set mRS = mDB.OpenRecordset("select * from BookMarks where id = " & msKey)
                If Not (mRS.BOF And mRS.EOF) Then
                    msBookMarkText = mRS.Fields("description").Value
                    msCodeKey = mRS.Fields("codeid").Value
                End If
                mRS.Close
                Set mRS = Nothing
                
                mbRead = True
            End If
        Case Else
            ' do nothing
    End Select
    
    Exit Sub
vbErrorHandler:
    Err.Raise Err.Number, "CBookmark:GetAttributes", Err.Description
    
End Sub

Private Sub Class_Terminate()
    Set mDB = Nothing
End Sub

Private Sub IDataObject_Commit()
    Select Case mState
        Case doAwaitingUpdate
            UpdateObject
        Case doAwaitingInsert
            InsertObject
        Case doAwaitingDelete
            DeleteObject
    End Select
    
End Sub

Private Sub IDataObject_Delete()
    mState = doAwaitingDelete
End Sub


Private Sub IDataObject_Initialise(oDB As DAO.Database, Optional sKey As String)
    
    InitialiseProperties
    
    Set mDB = oDB
    
    If Len(sKey) > 0 Then
        msKey = sKey
        mState = doStored
    Else
        mState = doAwaitingInsert
    End If
    mbRead = False

End Sub

Private Property Get IDataObject_Key() As String
    IDataObject_Key = msKey
End Property

Private Sub SetStateForLet()

    Select Case mState
        Case doAwaitingInsert, doAwaitingUpdate
            ' State doesn't change
        Case doStored
            mState = doAwaitingUpdate
        Case doAwaitingDelete
            Err.Raise AppErrors.errAwaitingDelete, "CBookmark::SetStateForLet", "This Record is About to be deleted"
        Case doDeleted
            Err.Raise AppErrors.errObjectDeleted, "CBookmark::SetStateForLet", "This record has been deleted"
        Case Else
            Err.Raise AppErrors.errObjectNotCreated, "CBookmark::SetStateForLet", "This record has not been created yet"
    End Select

End Sub

Private Sub UpdateObject()
    Dim rs As Recordset

On Error GoTo vbErrorHandler

    Set rs = mDB.OpenRecordset("select * from bookmarks where id = " & msKey)
    rs.Edit
    rs.Fields("codeid").Value = msCodeKey
    rs.Fields("description").Value = msBookMarkText
    rs.Update
        
    rs.Close
    
    mState = doStored
    mbRead = True
    

    Exit Sub

vbErrorHandler:
    Err.Raise Err.Number, "CBookmark::UpdateObject", Err.Description

End Sub

Private Sub DeleteObject()
    Dim sSql As String
    
    sSql = "delete from Bookmarks where id = " & msKey
    mDB.Execute sSql
    mState = doDeleted
    
End Sub

Private Sub InsertObject()
    Dim rs As Recordset

On Error GoTo vbErrorHandler

    Set rs = mDB.OpenRecordset("select * from bookmarks where id = 0")
    With rs
        .AddNew
        .Fields("codeid").Value = msCodeKey
        .Fields("description").Value = msBookMarkText
        .Update
    End With
    rs.Bookmark = rs.LastModified
    msKey = rs.Fields("id")
    rs.Close
    Set rs = Nothing
    
    mState = doStored
    

    Exit Sub

vbErrorHandler:
    rs.Close
    Set rs = Nothing
    
    
    Err.Raise Err.Number, "CBookmark::InsertObject", Err.Description

End Sub
