Attribute VB_Name = "RegistryStartUp"
'This module makes a registry entry in;
'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run
'to make a program start up when windows loads.
'
'This program was made by me,
'Eric O' Sullivan. CompApp Technologys (tm)
'is my company. If this product is unsatisfactory
'feel free to contact me at
'DiskJunky@hotmail.com
'================================================
'================================================

Option Explicit

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


Public Const REG_SZ = 1 ' Unicode nul terminated String
Public Const REG_DWORD = 4 ' 32-bit number
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&


'Option Explicit
Const UNIT_NAME = "UTILITY"

Public Sub MakeStartUp(FileName As String)
Dim Counter As Integer
Dim MarkPos As Integer
Dim Application As String
    
Application = GetFileName(FileName)
Application = Left(Application, (Len(Application) - 4)) 'Replace(Application, ".exe", "", , , vbTextCompare) & "~@#"
Call SaveString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Application, FileName)
End Sub

Public Sub SaveKey(hKey As Long, strPath As String)
    Dim KeyHand&
    Dim r As Long
    
    r = RegCreateKey(hKey, strPath, KeyHand&)
    r = RegCloseKey(KeyHand&)
End Sub

Public Function GetString(hKey As Long, strPath As String, strValue As String)
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '
    ' R, "Software\VBW\Registry", "String")
    '
    Dim KeyHand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim r As Long
    Dim lValueType As Long
    
    r = RegOpenKey(hKey, strPath, KeyHand)
    lResult = RegQueryValueEx(KeyHand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(KeyHand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                GetString = Left(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '
    ' tware\VBW\Registry", "String", text1.t
    '     ex
    ' t)
    '
    Dim KeyHand As Long
    Dim r As Long
    
    r = RegCreateKey(hKey, strPath, KeyHand)
    r = RegSetValueEx(KeyHand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(KeyHand)
End Sub


Function GetDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    'EXAMPLE:
    '
    'text1.text = getdword(HKEY_CURRENT_USER
    '
    ' , "Software\VBW\Registry", "Dword")
    '
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim KeyHand As Long
    
    r = RegOpenKey(hKey, strPath, KeyHand)
    ' Get length/data type
    lDataBufSize = 4
    lResult = RegQueryValueEx(KeyHand, strValueName, 0&, lValueType, lBuf, lDataBufSize)


    If lResult = ERROR_SUCCESS Then


        If lValueType = REG_DWORD Then
            GetDWord = lBuf
        End If
        'Else
        'Call errlog("GetDWORD-" & strPath, Fals
        '
        ' e)
    End If
    r = RegCloseKey(KeyHand)
End Function


Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    'EXAMPLE"
    '
    'Call SaveDword(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword", text1.text)
    '
    '
    Dim lResult As Long
    Dim KeyHand As Long
    Dim r As Long
    
    r = RegCreateKey(hKey, strPath, KeyHand)
    lResult = RegSetValueEx(KeyHand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then
    ' Call errlog("SetDWORD", False)
    r = RegCloseKey(KeyHand)
End Function


Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '
    ' ware\VBW")
    '
    Dim r As Long
    
    r = RegDeleteKey(hKey, strKey)
End Function


Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '
    ' ftware\VBW\Registry", "Dword")
    '
    Dim KeyHand As Long
    Dim r As Long
    
    r = RegOpenKey(hKey, strPath, KeyHand)
    r = RegDeleteValue(KeyHand, strValue)
    r = RegCloseKey(KeyHand)
End Function

Public Sub DeleteFromStartup(FileName As String)
Dim Counter As Integer
Dim MarkPos As Integer
Dim Application As String
   
Application = GetFileName(FileName)
Application = Left(Application, (Len(Application) - 4)) 'Replace(Application, ".exe", "", , , vbTextCompare) & "~@#"
Call DeleteValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Application)
End Sub

Public Function GetFileName(Path As String) As String
'returnes the filename from a path.

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


'Private Sub Form_Load()
'    cmdSelectExe.Visible = False
'End Sub

'Private Sub txtFilename_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error GoTo errh
'
'    txtFilename.Text = Data.Files(1)
'
'    If InStr(1, txtFilename.Text, ".exe", vbTextCompare) = 0 Then
'        MsgBox "Please Drag and Drop Exe File Only", vbExclamation, UNIT_NAME
'        txtFilename.Text = ""
'        txtFilename.SetFocus
'    End If
'
'    Exit Sub
'errh:
'    MsgBox "Please Drag and Drop Exe File Only", vbExclamation, UNIT_NAME
'    'nothing
'End Sub

