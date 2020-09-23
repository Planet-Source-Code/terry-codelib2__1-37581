VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Object for storing datafiles inside the database
'

Implements IDataObject

Private msKey As String ' key
Private msCodeID As String ' code id
Private msDescription As String ' description
Private msFileName As String ' original file name
Private mdteOrigDate As Date ' original date of file
Private mdteDateAdded As Date ' date file added to database

Private mDB As Database ' database object

Private Const CHUNKSIZE As Long = 16384 ' internal chunksize
Private mState As doState ' internal state
Private mbRead As Boolean ' internal state
Private mbFileChanged As Boolean ' internal state again

Private Sub InitialiseProperties()
'
' Initialise the properties for this object
'
    mState = doStored
    mbFileChanged = False
    msKey = ""
    msDescription = ""
    msCodeID = ""
    mbRead = False
End Sub

Public Property Get CodeID() As String
'
' Return the Parent ID (CodeItem ID)
'
    GetAttributes
    CodeID = msCodeID
End Property

Public Property Let CodeID(ByVal sCodeID As String)
'
' Set the Parent ID (codeitem ID)
'
    GetAttributes
    msCodeID = sCodeID
    SetStateForLet
End Property

Public Property Get Description() As String
'
' Get the description of this file (original File Name)
'
    GetAttributes
    Description = msDescription
End Property

Public Property Let StoredFile(ByVal sFilename As String)
'
' Set the stored file name
'
    GetAttributes
    If Len(msDescription) > 0 Then
'
' This internal flag tells us to copy the file into the database when the
' object is committed.
'
        mbFileChanged = True
    End If
    msFileName = sFilename
    BuildFile sFilename
    SetStateForLet
End Property

Private Sub Class_Terminate()
'
' Kill our Database object reference
'
    Set mDB = Nothing
End Sub

Private Sub IDataObject_Commit()
'
' Commit the object to the database - update, delete, or insert
'
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
'
' Initialise properties
'
    InitialiseProperties
'
' Record database reference
'
    Set mDB = oDB
'
' If we are passed a key - assume this object exists in database
'
    If Len(sKey) > 0 Then
        msKey = sKey
        mState = doStored
    Else
'
' Otherwise, assume it's to be inserted
'
        mState = doAwaitingInsert
    End If
    mbRead = False
    
End Sub

Private Property Get IDataObject_Key() As String
'
' Return object's key
'
    IDataObject_Key = msKey
End Property

Private Sub GetAttributes()
    Dim mRS As Recordset
    Dim iDO As IDataObject
'
' Get the attributes from the database
'
    On Error GoTo vbErrorHandler
    
    Select Case mState
        Case doStored, doAwaitingUpdate
            If Not mbRead Then
                'get details
                Set mRS = mDB.OpenRecordset("select * from codefiles where id = " & msKey)
                If Not (mRS.BOF And mRS.EOF) Then
                    msCodeID = mRS.Fields("codeid").Value
                    msDescription = mRS.Fields("description").Value
                    mdteDateAdded = mRS.Fields("DateAdded").Value
                    mdteOrigDate = mRS.Fields("OrigDateTime").Value
                End If
                mRS.Close
                Set mRS = Nothing
                
                mbRead = True
                
            End If
        Case Else
            ' do nothing
    End Select
    DBEngine.Idle dbRefreshCache
    
    Exit Sub

vbErrorHandler:
    Err.Raise Err.Number, "CFileObject:GetAttributes", Err.Description
    
End Sub

Private Sub InsertObject()
    Dim rs As Recordset
'
' Insert the object into the database
'
On Error GoTo vbErrorHandler

'
' Record the original date added to the database
'
    mdteDateAdded = Now()
    
    Set rs = mDB.OpenRecordset("select * from codefiles where id = 0")
    
    With rs
        .AddNew
        .Fields("codeid").Value = msCodeID
        .Fields("description").Value = msDescription
'
' Copy the file into the recordset field#
'
        BuildRSFile rs
'
' Record the required date/times
'
        .Fields("origdatetime").Value = mdteOrigDate
        .Fields("dateadded").Value = mdteDateAdded
        .Update
        .Bookmark = .LastModified
        msKey = .Fields("id")
        .Close
    End With
'
' Refresh the DB Engine
'
    DBEngine.Idle dbRefreshCache
    
    mState = doStored

    Exit Sub

vbErrorHandler:
    rs.Close
    Set rs = Nothing
    
    Err.Raise Err.Number, "CFileObject::InsertObject", Err.Description

End Sub

Private Sub SetStateForLet()
'
' Set the internal state of the object
'
    Select Case mState
        Case doAwaitingInsert, doAwaitingUpdate
            ' State doesn't change
        Case doStored
            mState = doAwaitingUpdate
        Case doAwaitingDelete
            Err.Raise AppErrors.errAwaitingDelete, "CFileObject::SetStateForLet", "This Record is About to be deleted"
        Case doDeleted
            Err.Raise AppErrors.errObjectDeleted, "CFileObject::SetStateForLet", "This record has been deleted"
        Case Else
            Err.Raise AppErrors.errObjectNotCreated, "CFileObject::SetStateForLet", "This record has not been created yet"
    End Select

End Sub

Private Sub UpdateObject()
    Dim rs As Recordset
'
' Update the object in the database
'
On Error GoTo vbErrorHandler
    
    Set rs = mDB.OpenRecordset("select * from codefiles where id = " & msKey)

    rs.Edit
    rs.Fields("codeid").Value = msCodeID
    rs.Fields("description").Value = msDescription
'
' Only rebuild the file if it's been changed to a different one
'
    If mbFileChanged Then
        BuildRSFile rs
    End If
'
' Shouldn't update the dates here!
'
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
    Err.Raise Err.Number, "CFileObject::UpdateObject", Err.Description

End Sub

Private Sub DeleteObject()
    Dim sSql As String
'
' Delete the object from the database
'
    sSql = "delete from codefiles where id = " & msKey
    mDB.Execute sSql
'
' Refresh the DB Engine
'
    DBEngine.Idle dbRefreshCache
'
' Set the internal state of the object
'
    mState = doDeleted
    
End Sub


Private Sub BuildFile(ByVal sFilename As String)
    Dim iFileNum As Integer
    Dim mVar As Variant
    Dim tSFI As SHFILEINFO
    Dim sDescription As String
    Dim lPos As Long
    Dim sExtension As String
    Dim lCount As Long
'
' Build the file
'
' Get the Display name, ie. NotePad from c:\windows\system\notepad.exe
'
    If SHGetFileInfo(sFilename, 0, tSFI, Len(tSFI), SHGFI_USEFILEATTRIBUTES Or SHGFI_DISPLAYNAME) Then
        sDescription = tSFI.szDisplayName
    Else
        sDescription = sFilename
    End If
    
'
' Record the Extension of this file (useful for when we need to export)
'
    lPos = InStr(1, sFilename, ".")
    If lPos > 0 Then
        sExtension = Right$(sFilename, (Len(sFilename) - lPos))
    End If
    
    lPos = InStr(1, sDescription, vbNullChar)
    If lPos > 0 Then
        sDescription = Left$(sDescription, lPos - 1)
    End If
'
' Now build the complete filename (minus Path)
'
    If InStr(1, sDescription, "." & sExtension) Then
        msDescription = sDescription
    Else
        msDescription = sDescription & "." & sExtension
    End If
End Sub

Private Sub BuildRSFile(rs As Recordset)
    
    Dim lLen As Long
    Dim lCount As Long
    Dim lFragment As Long
    Dim lChunks As Long
    Dim bChunk() As Byte
    Dim iFileNum As Integer
    Dim oField As Field
'
' Copy the File into the recordset field
'
    On Error GoTo vbErrorHandler
        
    iFileNum = FreeFile
'
' Open the file for binary access so we can read it in chunks
'
    Open msFileName For Binary Access Read As iFileNum
'
' Get Original Date/Time of the File for storing in the Database
'
    mdteOrigDate = FileDateTime(msFileName)
    
    lLen = LOF(iFileNum)
'
' Get the number of chunks
'
    lChunks = lLen \ CHUNKSIZE
'
' Get the small fragment size
'
    lFragment = lLen Mod CHUNKSIZE
    
    ReDim bChunk(lFragment)
    
    Get iFileNum, , bChunk
    Set oField = rs("file")
    
    oField.Value = ""
'
' Append the first chunk
'
    oField.AppendChunk bChunk
    
    ReDim bChunk(CHUNKSIZE)
'
' Now read in the rest of the file into the field
'
    For lCount = 1 To lChunks
        Get iFileNum, , bChunk()
        oField.AppendChunk bChunk
    Next
'
' Close the file
'
    Close iFileNum
    
    Exit Sub
    
vbErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub SaveToFile(ByVal sFilename As String)
'
' Export the file from the database to the passed filename
'
    Dim iFileNum As Integer
    Dim lFileLen As Long
    Dim lChunks As Long
    Dim lFragment As Long
    Dim bChunk() As Byte
    Dim lCount As Long
    Dim oField As Field
    Dim oRS As Recordset
'
' Check that everythings been read in !
'
    GetAttributes
   
    On Error GoTo vbErrorHandler
'
' Get the field from the database
'
    Set oRS = mDB.OpenRecordset("select file from codefiles where id = " & msKey)
    
    If oRS.BOF Or oRS.EOF Then Exit Sub
    
    iFileNum = FreeFile
'
' Create the Named File
'
    Open sFilename For Binary Access Write As iFileNum
    Set oField = oRS.Fields("file")
    
'
' Get the length of the file and the number of chunks required
'
    lFileLen = oField.FieldSize
    lChunks = lFileLen \ CHUNKSIZE
    lFragment = lFileLen Mod CHUNKSIZE
'
' Write away the chunks to the file
'
    For lCount = 1 To lChunks
        ReDim bChunk(CHUNKSIZE)
        bChunk() = oField.GetChunk(((lCount - 1) * CHUNKSIZE), CHUNKSIZE)
        Put iFileNum, , bChunk()
    Next
'
' Write the final (or first if lChunks = 0) chunk
'
    ReDim bChunk(lFragment)
    bChunk() = oField.GetChunk(lChunks * CHUNKSIZE, lFragment)
    
    Put iFileNum, , bChunk()
    Close iFileNum
    
    oRS.Close
    Set oRS = Nothing
    
    Exit Sub

vbErrorHandler:
    Err.Raise Err.Number, Err.Source & "::CFileObject_SaveToFile", Err.Description
End Sub
