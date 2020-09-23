VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Settings Object for this project
'
' Chris Eastwood Feb.1999
'
' TODO:
'
' Implement a method to backup to other datasources
' besides the registry.
'

Private mbBackupDatabaseOnStartup As Boolean
Private mbCompactDatabaseOnExit As Boolean
Private mbSaveFormLayout As Boolean

Public Property Get SaveFormLayout() As Boolean
    SaveFormLayout = mbSaveFormLayout
End Property

Public Property Let SaveFormLayout(ByVal bSave As Boolean)
    mbSaveFormLayout = bSave
End Property

Public Property Get BackupDatabaseAtStart() As Boolean
    BackupDatabaseAtStart = mbBackupDatabaseOnStartup
End Property

Public Property Let BackupDatabaseAtStart(ByVal bBackup As Boolean)
    mbBackupDatabaseOnStartup = bBackup
End Property

Public Property Get CompactDatabaseOnExit() As Boolean
    CompactDatabaseOnExit = mbCompactDatabaseOnExit
End Property

Public Property Let CompactDatabaseOnExit(ByVal bCompact As Boolean)
    mbCompactDatabaseOnExit = bCompact
End Property

Private Sub Class_Initialize()
    mbCompactDatabaseOnExit = GetSetting("VBCodeLib", "Settings", "CompactOnExit", True)
    mbBackupDatabaseOnStartup = GetSetting("VBCodeLib", "Settings", "BackupAtStart", False)
    mbSaveFormLayout = GetSetting("VBCodeLib", "Settings", "SaveFormLayout", True)
End Sub

Private Sub Class_Terminate()
    SaveSetting "VBCodeLib", "Settings", "CompactOnExit", mbCompactDatabaseOnExit
    SaveSetting "VBCodeLib", "Settings", "BackupAtStart", mbBackupDatabaseOnStartup
    SaveSetting "VBCodeLib", "Settings", "SaveFormLayout", mbSaveFormLayout
End Sub
