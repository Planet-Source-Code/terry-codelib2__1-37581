VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCodeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Data Object for Code in the DataBase
'
' Chris Eastwood February 1999
'
Implements IDataObject
'
' Private members
'
Private msKey As String
Private msParentKey As String
Private msDescription As String


Private msNotes As String
Private msCode As String
Private msExample As String

Private mState As doState
Private mbRead As Boolean

Private moFileObject As CFileObject

Private mDB As Database

Public Property Get Notes() As String
    
    GetAttributes
    Notes = msNotes
    
End Property

Public Property Let Notes(ByVal sNotesText As String)
    
    GetAttributes
    msNotes = sNotesText
    SetStateForLet
    
End Property

Public Property Get Example() As String
    GetAttributes
    Example = msExample
End Property

Public Property Let Example(ByVal sExampleText As String)
    GetAttributes
    msExample = sExampleText
    SetStateForLet
End Property

Public Property Get ParentKey() As String
    GetAttributes
    ParentKey = msParentKey
End Property

Public Property Let ParentKey(ByVal sParentKey As String)
    GetAttributes
    msParentKey = sParentKey
    If Len(msParentKey) = 0 Then
        msParentKey = "0"
    End If
    SetStateForLet
End Property

Public Property Get Description() As String
    GetAttributes
    Description = msDescription
End Property

Public Property Let Description(ByVal sDesc As String)
    GetAttributes
    msDescription = sDesc
    SetStateForLet
End Property

Public Property Get Code() As String
    
    GetAttributes
    Code = msCode
End Property

Public Property Let Code(ByVal sCodeText As String)
    GetAttributes
    msCode = sCodeText
    SetStateForLet
End Property

Private Sub InitialiseProperties()
    mState = doStored
    msKey = ""
    msParentKey = "0"
    msDescription = ""
    msCode = ""
    msExample = ""
    msNotes = ""
    mbRead = False
End Sub

Private Sub GetAttributes()
    Dim mRS As Recordset
    Dim iDO As IDataObject
    
    On Error GoTo vbErrorHandler
    
    Select Case mState
        Case doStored, doAwaitingUpdate
            If Not mbRead Then
                'get details
                Set mRS = mDB.OpenRecordset("select * from codeitems where id = " & msKey)
                If Not (mRS.BOF And mRS.EOF) Then
                    msParentKey = mRS.Fields("parentid").Value & ""
                    msDescription = mRS.Fields("description").Value & ""
                    On Error Resume Next
                    msNotes = mRS.Fields("Notes").GetChunk(0, 32768) & ""
                    msCode = mRS.Fields("Code").GetChunk(0, 32768) & ""
                    msExample = mRS.Fields("Example").GetChunk(0, 32768) & ""
                    On Error GoTo vbErrorHandler
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
    Err.Raise Err.Number, "CCodeItem:GetAttributes", Err.Description
    
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
        Case Else
            UpdateObject
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
            Err.Raise AppErrors.errAwaitingDelete, "CCodeItem::SetStateForLet", "This Record is About to be deleted"
        Case doDeleted
            Err.Raise AppErrors.errObjectDeleted, "CCodeItem::SetStateForLet", "This record has been deleted"
        Case Else
            Err.Raise AppErrors.errObjectNotCreated, "CCodeItem::SetStateForLet", "This record has not been created yet"
    End Select

End Sub

Private Sub UpdateObject()
    Dim rs As Recordset

On Error GoTo vbErrorHandler

    Set rs = mDB.OpenRecordset("select * from codeitems where id = " & msKey)
    rs.Edit
    rs.Fields("code").Value = ""
    rs.Fields("code").AppendChunk msCode
    rs.Fields("description").Value = msDescription
    rs.Fields("parentid") = msParentKey
    rs.Fields("notes").Value = ""
    rs.Fields("notes").AppendChunk msNotes
    rs.Fields("example").Value = ""
    rs.Fields("example").AppendChunk msExample
    rs.Update
        
    rs.Close
'
' Refresh the DB Engine
'
    DBEngine.Idle dbRefreshCache
    
    mState = doStored
    mbRead = True

    Exit Sub

vbErrorHandler:
    Err.Raise Err.Number, "CCodeItem::UpdateObject", Err.Description

End Sub

Private Sub DeleteObject()
    Dim sSql As String
    
    sSql = "delete from codeitems where id = " & msKey
    mDB.Execute sSql
    sSql = "delete from bookmarks where codeid = " & msKey
    mDB.Execute sSql
    sSql = "delete from codefiles where codeid = " & msKey
    mDB.Execute sSql
    mState = doDeleted
'
' Refresh the DB Engine
'
    DBEngine.Idle dbRefreshCache
    
End Sub

Private Sub InsertObject()
    Dim rs As Recordset

On Error GoTo vbErrorHandler

    Set rs = mDB.OpenRecordset("select * from codeitems where id = 0")
    With rs
        .AddNew
        .Fields("parentid").Value = msParentKey
        .Fields("description").Value = msDescription
        .Fields("code").AppendChunk msCode
        .Fields("notes").AppendChunk msNotes
        .Fields("example").AppendChunk msExample
        .Update
        .Bookmark = .LastModified
        msKey = .Fields("id")
        .Close
    End With
    
    mState = doStored
'
' Refresh the DB Engine
'
    DBEngine.Idle dbRefreshCache
    
    Exit Sub

vbErrorHandler:
    rs.Close
    Set rs = Nothing
    
    
    Err.Raise Err.Number, "CCodeItem::InsertObject", Err.Description

End Sub
