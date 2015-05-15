Option Compare Database
Option Explicit

''''''''''''''''''''''''''''''
' Backup related functions   '
''''''''''''''''''''''''''''''


'''
' Auto backup of a chosen exported report
'
' Copies a given file from the DB's [ROOTPATH] and pastes a timestamped version in [BACKUPPATH]
'
' Dependencies:
' - Constant [ROOTPATH] defining the path to any given filenames
' - Constant [BACKUPPATH] defining the path to the 'backups' folder
' - Procedure log_event() to log and/or output error messages and events
'
' Parameters:
' - report_name (string) Defines the name of the report to be backed up
'
'''
Public Function auto_backup(ByVal fromPath As String, toPath As String) As Boolean
On Error GoTo err_handler
    
    Dim FSO As Object
    
    auto_backup = False
    
    ' ensure we don't have double slashes (except at the beginning of the path to indicate server names)
    fromPath = Replace(fromPath, "\", "\")
    toPath = Replace(toPath, "\", "\")
    If Left(fromPath, 1) = "\" Then fromPath = "\" & fromPath
    If Left(toPath, 1) = "\" Then toPath = "\" & toPath
    
    
    ' This copies a file from FromPath to ToPath.
    ' Note: If ToPath already exists it will overwrite existing file
    ' if ToPath does not exist it will be created
       
    If Right(fromPath, 1) = "\" Then fromPath = Left(fromPath, Len(fromPath) - 1)
    If Right(toPath, 1) = "\" Then toPath = Left(toPath, Len(toPath) - 1)

    Set FSO = CreateObject("scripting.filesystemobject")

    ' if FromPath is specified incorrectly (i.e. doesn't exist)
    If Not FSO.FileExists(fromPath) And Not FSO.folderexists(fromPath) Then
        log_event fromPath & " does not exist", ioError
        Exit Function
    End If

    ' create folder housing toPath if it doesn't already exist
    If Not FSO.folderexists(dirname(toPath, "\") & "\") Then MkDir (dirname(toPath, "\"))
    
    ' if able to copy and paste the file, return True
    FSO.CopyFile Source:=fromPath, Destination:=toPath
    Set FSO = Nothing
    log_event "Backup of " & basename(fromPath) & " complete", ioSuccess
    auto_backup = True
    
    ' error catching
err_handler:
    If Err.Number <> 0 Then
        log_event "Error with auto backup: " & Err.Description, ioError
        auto_backup = False
        Exit Function
    End If
End Function


'''
' Auto Backup of database file itself and dependent spreadsheets
'
' Dependencies:
' - Procedure auto_backup()
'''
Public Function self_backup() As Boolean
    ' catch errors
    On Error GoTo err_handler
    
    If ENVIRONMENT = "DEV" Then
        log_event "Self backup prevented due to being in DEV environment.", ioMessage
        self_backup = True ' return true even though no backup takes place - to ensure there are no errors later in testing
        Exit Function
    End If
    
    Dim strBackupPath As String
    Dim strBackupDate As String
    Dim strFiles2Backup As String
    Dim strArrFiles2Backup() As String
    Dim vCurFile As Variant
    Dim strSplitStr() As String
    Dim strTeam As String
    Dim strCurPath As Variant
    Dim bSuccess As Boolean
    
    ' get backup info
    strBackupPath = SM_getSetting("backup_path", "system")
    strBackupDate = Format(Now, "yyyy-mm-dd-hh-nn-ss ")
    strFiles2Backup = SM_getSetting("filesToBackup", "system")

    If "" = strBackupPath Or "" = strFiles2Backup Then
        log_event "Unable to retrieve settings for self backup", ioError
        self_backup = False
        Exit Function
    End If
    
    strArrFiles2Backup = Split(strFiles2Backup, ",")
    strFiles2Backup = ""
    
    bSuccess = True
    For Each vCurFile In strArrFiles2Backup
        If Not bSuccess Then Exit For
        
        Select Case vCurFile
            Case "self"
                strCurPath = CurrentProject.FullName
            
            ' compare vCurFile with itself only if it contains 'KPIs*'
            Case IIf(vCurFile Like "KPIs*", vCurFile, "")
                '' Future improvement: loop through all teams for 'KPIs' rather than having to specify each team in turn
                On Error Resume Next
                strSplitStr = Split(vCurFile, "_")
                strTeam = strSplitStr(1)
                On Error GoTo err_handler
                ' get the daily KPI path for the relevant team (all teams must be set up in the "filesToBackup" setting in order to be backed up)
                strCurPath = SM_getSetting("dailyKPI_path", strTeam)
            Case Else
                strCurPath = SM_getSetting("dailyImpPath", vCurFile)
        End Select
        
        bSuccess = auto_backup(strCurPath, strBackupPath & strBackupDate & basename(strCurPath, "\"))
    Next vCurFile
    
    If Not bSuccess Then
        log_event "Error with self backup - please check the logs...", ioError
        self_backup = False
        Exit Function
    End If
    
    ' only true if it works
    log_event "Self backup completed", ioSuccess
    self_backup = True
    
    ' error catching
err_handler:
    If Err.Number <> 0 Then
        log_event "Error with self backup: " & Err.Description, ioError
        self_backup = False
        Exit Function
    End If
End Function
