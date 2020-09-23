Attribute VB_Name = "FileProcedures"
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Const OFS_MAXPATHNAME = 128
Private Const OF_PARSE = &H100
Private Const OF_SHARE_DENY_NONE = &H40

Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Enum FileAccessEnum
    CreationTime = 0
    LastAccess = 1
    LastWrite = 2
End Enum

Public Enum AccessType
    FileInput = 0
    FileOutPut = 1
End Enum

Public Function GetPath(Address As String) As String
'get the path section of the string passed

Dim Counter As Integer
Dim LastPos As Integer

LastPos = 1
For Counter = 1 To Len(Address)
    If Mid(Address, Counter, 1) = "\" Then
        LastPos = Counter
    End If
Next Counter

If LastPos = 3 Then
    'the path is the drive - include the backslash
    LastPos = 4
End If
GetPath = Left(Address, (LastPos - 1)) 'Mid(Path, 1, (LastPos - 1))
End Function

Public Function GetFileName(Path As String) As String
'get the filename section of the string passed

Dim Counter As Integer
Dim LastPos As Integer

LastPos = 1
For Counter = 1 To Len(Path)
    If Mid(Path, Counter, 1) = "\" Then
        LastPos = Counter
    End If
Next Counter

GetFileName = Mid(Path, (LastPos + 1), Len(Path))

End Function

Public Function GetDateOfFile(Path As String, WhichTime As FileAccessEnum) As String
'returns the date of the file specified (minus one day)

Dim FilePath(OFS_MAXPATHNAME) As Byte
Dim Counter As Byte
Dim TimeInfo(3) As FILETIME
Dim Result As Long
Dim OpenStruc As OFSTRUCT
Dim FileHandle As Long

Dim TheFileTime As SYSTEMTIME
Dim TheTime As String
Dim ThePath As String
Dim TheFile As String

'if the file doesn't exist or the path is invalid, then exit
If Path = "" Then
    Exit Function
Else
    If Dir(Path, (vbArchive + vbDirectory + vbHidden + vbReadOnly + vbSystem)) = "" Then
        Exit Function
    End If
End If

'parse the filename and path into seperate variables
ThePath = GetPath(Path)
TheFile = GetFileName(Path)

'open file and retreive the time in nanoseconds before converting into
'system time structure.
ChDrive (ThePath)
ChDir (ThePath)
FileHandle = OpenFile(TheFile, OpenStruc, OF_SHARE_DENY_NONE)
Result = GetFileTime(FileHandle, TimeInfo(CreationTime), TimeInfo(LastAccess), TimeInfo(LastWrite))
Result = CloseHandle(FileHandle)

'return the proper time
Select Case WhichTime
Case CreationTime
    Result = FileTimeToSystemTime(TimeInfo(CreationTime), TheFileTime)
Case LastAccess
    Result = FileTimeToSystemTime(TimeInfo(LastAccess), TheFileTime)
Case LastWrite
    Result = FileTimeToSystemTime(TimeInfo(LastWrite), TheFileTime)
End Select

'assemble the date and file of the file from the info retrieved
TheTime = TheFileTime.wDay & "/" & TheFileTime.wMonth & "/" & TheFileTime.wYear & "  " & Format(TheFileTime.wHour, "00") & ":" & Format(TheFileTime.wMinute, "00") & ":" & Format(TheFileTime.wSecond, "00")

GetDateOfFile = TheTime
End Function

Public Function DeleteFiles(ByVal Path As String, ByVal DaysOld As Integer) As Long
'This deletes all the files within the specified path that are more than
'X days old.
'This function will delete ANY file it finds - even read-only or
'hidden files/directories. If the file is in use though, it will cause
'an error.
'The function returns the number of files it deleted

Dim GetAllDir() As String
Dim GetAllFiles() As String

Dim MaxFiles As Integer
Dim MaxDir As Integer

Dim Name As String
Dim MyDateTime As SYSTEMTIME
Dim Counter As Long

Dim FileAge As Long

Dim TotalDel As Long

'first check to see if the path is valid
If (Dir(Path, vbDirectory) = "") Or (Path = "") Then
    'path invalid
    Exit Function
End If

'Find all the files in the specified path
MaxFiles = 0
ReDim Preserve GetAllFiles(MaxFiles)

'Find all the directories in the specified path
MaxDir = 0
ReDim Preserve GetAllDir(MaxDir)

Name = Dir(AddFile(Path, "*.*"), (vbDirectory + vbArchive + vbHidden + vbReadOnly + vbSystem))
While Name <> ""
    'check files
    If (Not IsDirectory(GetAttr(AddFile(Path, Name)))) And (Left(Name, 1) <> ".") Then
        'add name of directory to the list
        ReDim Preserve GetAllFiles(MaxFiles)
        GetAllFiles(MaxFiles) = AddFile(Path, Name)
        MaxFiles = MaxFiles + 1
    End If
    
    'check directories
    If IsDirectory(GetAttr(AddFile(Path, Name))) And (Left(Name, 1) <> ".") Then
        'add name of directory to the list
        ReDim Preserve GetAllDir(MaxDir)
        GetAllDir(MaxDir) = AddFile(Path, Name)
        MaxDir = MaxDir + 1
    End If
    Name = Dir
Wend

On Error Resume Next
'delete all files found if older than a specified number of days
If MaxFiles > 0 Then
    For Counter = 0 To MaxFiles - 1
        FileAge = DateDiff("d", GetDateOfFile(GetAllFiles(Counter), LastWrite), Now)
        If FileAge > DaysOld Then
            'change the attributes of the file so we can delete it.
            Call SetAttr(GetAllFiles(Counter), vbNormal)
            Kill (GetAllFiles(Counter))
            TotalDel = TotalDel + 1
        End If
    Next Counter
End If

'delete all files in the sub-directorys found
For Counter = 0 To MaxDir - 1
    Call SetAttr(GetAllDir(Counter), vbNormal)
    TotalDel = TotalDel + DeleteFiles(GetAllDir(Counter), DaysOld)
    If GetAllDir(Counter) <> "" Then
        'remove the current directory of needed
        RmDir (GetAllDir(Counter))
    End If
Next Counter

'return file count
DeleteFiles = TotalDel
End Function

Public Function AddFile(ByVal Path As String, ByVal File As String) As String
'This procedure adds a file name to a path.

If Right(Path, 2) = ":\" Then
    Path = Path & File
Else
    Path = Path & "\" & File
End If

AddFile = Path
End Function

Public Function IsDirectory(Flags As Integer) As Boolean
'this function returns wether or not the attribute number of a file
'contains the directory flag.

Dim Counter As Integer
Dim NextFlag  As Integer
Dim StartNum As Integer

If Flags < 16 Then
    Exit Function
End If

NextFlag = Flags
For Counter = 8 To 4 Step -1
    If NextFlag >= (2 ^ Counter) Then
        If Counter = 4 Then
            'directory found if the flag contains vbDirectory
            '(2 ^ 4 = 16 = vbDirectory)
            IsDirectory = True
            Exit Function
        Else
            NextFlag = NextFlag - (2 ^ Counter)
        End If
    End If
Next Counter
End Function

Public Function CanAccessFile(FileName As String, Access As AccessType) As Boolean
'This function returns whether or not if a file can be accessed

Dim ErrNum As Long
Dim FileNum As Integer

'prevent errors from stopping execution of the following code
On Error Resume Next

'get a free file access number
FileNum = FreeFile

Select Case Access
Case FileInput
    Open FileName For Input As #FileNum
    ErrNum = Err
Case FileOutPut
    Open filenname For Output As #FileNum
End Select

'close file access
Close #FileNum

If ErrNum = 0 Then
    'no error occurred. it's safe to access the file
    CanAccessFile = True
End If
End Function

