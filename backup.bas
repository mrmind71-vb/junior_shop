Attribute VB_Name = "backup"
Private Declare Function PathIsNetworkPath Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Long
Public Function isNetWorkPath() As Boolean
'Dim isNetwork As Long
Dim appFolderPath As String

' App.Path returns the directory the executable is in.
' Note: When running in the VB IDE, App.Path refers to the IDE's path,
' so test with a compiled EXE.
appFolderPath = App.Path

' PathIsNetworkPath returns True (non-zero) if it is a network path.
isNetWorkPath = PathIsNetworkPath(appFolderPath) = 1
End Function
Public Function myBackup(con As ADODB.Connection, Optional nCount As Integer = 10, Optional pIgExist As Boolean = False) As Boolean
Dim fs As New FileSystemObject
Dim sSource As String, sDest As String

On Error GoTo myerror
'sLastDrive = LastDrive(True)
sLastDrive = "D"
MyCreateFolder sLastDrive & ":\backup"

sDest = sLastDrive & ":\backup\" & retBackUpFileName

'sSource = App.Path & "\backup\"
'FixFiles sSource, nCount
'If UCase(RetSetting("MAIN")) <> "YES" Then
'    FixFiles sDest, nCount
'End If


sSource = App.Path & "\backup\" & retBackUpFileName
If (Not fs.FileExists(sSource)) Or pIgExist Then
    Inform "”Ì „ ⁄„· ‰”Œ… «Õ Ì«ÿÌ… ⁄·Ì «·”Ì—›—"
    Dim sSourceMdf As String
    sSourceMdf = getFolderBackup(con) & "\backup\" & retBackUpFileName
    createBackup sSourceMdf, pIgExist, con
End If

If isNetWorkPath Then
    If (Not fs.FileExists(sDest)) And fs.FileExists(sSource) Then
        Inform "”Ì „ ‰”Œ ‰”Œ… «Õ Ì«ÿÌ… „‰ «·”Ì—›—"
        fs.CopyFile sSource, sDest
        MsgBox " „ ‰”Œ «·„·› »‰Ã«Õ"
    End If
End If

If isNetWorkPath Then
    FixFiles sDest
Else
    FixFiles sSource
End If
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function retBackUpFileName() As String
retBackUpFileName = sCatalog & "_" & Format(Date, "yymmdd") & ".bak"
End Function
Private Function createBackup(pFileName As String, pIgExist As Boolean, con As ADODB.Connection) As Boolean
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "catalog", sCatalog)
aPrm = AddFlag(aPrm, "BackupPath", pFileName)
aPrm = AddFlag(aPrm, "Ig_Exist", IIf(pIgExist, 1, 0))

Dim command As New ADODB.command
Set command = cmd("[dbo].[BackupIfNotExists]", con, adStoredProc, aPrm)
command.Execute
If Not IsNull(command.Parameters("@MSG")) Then
    If command.Parameters("@MSG") <> "ok" Then
        MsgBox command.Parameters("@MSG") & vbCrLf & _
               "·„ Ì „ ⁄„· „·› Backup"
        Exit Function
    Else
        MsgBox " „ ⁄„· „·› backup"
    End If
End If
createBackup = True
End Function
Function FixFiles(pDir As String, Optional nMaxFiles As Integer = 10) As Boolean
Dim fs As New FileSystemObject
Dim aret As Variant, nDelete As Long
On Error Resume Next
aret = retFArray(pDir, "bak")
nDelete = (UBound(aret) + 1) - nMaxFiles
For i = 0 To (nDelete)
    fs.DeleteFile pDir & "\" & aret(i)
Next
Err.Clear
End Function
Private Function retFArray(pFolder As String, sExt As String) As Variant
Dim fso As New FileSystemObject, FileCount As Long
Dim fNames()
ReDim fNames(0)
If Not fso.FolderExists(pFolder) Then
    retFArray = fNames
    Exit Function
End If
Set fold = fso.GetFolder(pFolder)
For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = "." & sExt And Len(File.Name) > 4 Then
        If IsNumeric(Mid(File.Name, Len(sCatalog) + 2, 6)) Then FileCount = FileCount + 1
    End If
Next


ReDim fNames(FileCount)
cFcount = 0

For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = "." & sExt And Len(File.Name) > 4 Then
        If IsNumeric(Mid(File.Name, Len(sCatalog) + 2, 6)) Then
            cFcount = cFcount + 1
            fNames(cFcount) = LCase(File.Name)
        End If
    End If
Next

For tName = 1 To FileCount
    For nName = (tName + 1) To FileCount
        If StrComp(fNames(tName), fNames(nName), 0) = 1 Then
            buffer = fNames(nName)
            fNames(nName) = fNames(tName)
            fNames(tName) = buffer
        End If
    Next
Next
retFArray = fNames
End Function
Private Function getFolderBackup(con As ADODB.Connection) As String
Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "SELECT  SUBSTRING(physical_name, 1, (Len(physical_name) - Len(DB_NAME(database_id)) - 5)) As DataFilePath" & _
          " From sys.master_files Where type_desc = 'ROWS' " & _
          " AND DB_NAME(database_id) = " & MyParn(sCatalog)
Set loctable = cmd(cString, con).Execute
If Not loctable.EOF Then
    getFolderBackup = ParentFolder(loctable!DataFilePath & "")
End If
End Function
Public Function ParentFolder(ByVal fullPath As String) As String
    Dim lastBackslashPos As Long
    
    ' Find the position of the last backslash
    lastBackslashPos = InStrRev(fullPath, "\")
    
    ' If found, return the part of the string to the left of the last backslash
    If lastBackslashPos > 0 Then
        ParentFolder = Left(fullPath, lastBackslashPos - 1)
    Else
        ' No parent folder found (e.g., path is just "C:")
        ParentFolder = ""
    End If
End Function
Private Function MyCreateFolder(pDir As String, Optional bMsg As Boolean = False) As Boolean
On Error GoTo myerror
Dim fs As FileSystemObject
Set fs = CreateObject("Scripting.FileSystemObject")
aString = Split(pDir, "\")
cString = aString(0)
For i = 1 To UBound(aString)
    On Error Resume Next
    cString = cString & "\" & aString(i)
    If Not fs.FolderExists(cString) Then fs.CreateFolder (cString)
    Err.Clear
Next
MyCreateFolder = fs.FolderExists(cString)
Exit Function
myerror:
If bMsg Then MsgBox Err.Description
Err.Clear
End Function
Private Function LastDrive(Optional bLetter As Boolean = False)
Dim fs, d, DC, letter
Set fs = CreateObject("Scripting.FileSystemObject")
Set DC = fs.Drives
For Each d In DC
    If d.DriveType = 2 And d.DriveLetter <> "G" Then
        On Error Resume Next
        letter = IIf(bLetter, d.DriveLetter, d.SerialNumber)
    End If
Next
LastDrive = letter
End Function

