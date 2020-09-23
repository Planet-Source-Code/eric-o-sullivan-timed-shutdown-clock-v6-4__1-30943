Attribute VB_Name = "Encryption"
Option Explicit

'This encryption algorithim first gets the day of
'the current date (including the character "0" if
'applicable in two digits), converts both characters
'into ascii values, adds them to a prime number and
'uses that as the encryption key. The value is
'encrypted under a simple ascii value addition and
'can be deduced from the first few characters of the
'encrypted string. The first character of the
'encrypted data is ALWAYS the ascii value of how many
'characters after it is the decrypt key, ie the
'length of the decrypt key is the ascii value of the
'first character.
'
'==================================================
'I realise that a better method would be to use a
'"rolling key" method, ie, changing or incrementing
'the encryption key as each character is encrypted.
'But. I'll leave that to you.
'DiskJunky
'==================================================

Declare Function GetTickCount Lib "kernel32" () As Long

Const BaseKey = 43  'used to encrypt the main key
Const AddToKey = 17 'added to help form the main key

'the amount of characters encrypted during file operations.
Const FileIntake = 30720

Private Function GenerateKey() As Integer
'generates the main key use for encryption.

Dim MilliSecond As Integer

'I changed the daynum value to hold a second value
'instead of a day value for more variances.
'Changed again to an even shorter time value.
MilliSecond = (GetTickCount Mod 100)  '/ 1000)
GenerateKey = Val(Trim(Str(Format(MilliSecond, "00")))) + AddToKey 'Second(Time)
End Function

Public Function EncryptData(Text As String) As String
Dim Counter As Integer
Dim DayNum As String
Dim DayKey As Integer
Dim RetData As String
Dim Encrypt As String
Dim TextLen As Long

'if text is empty, return empty
If Text = "" Then
    EncryptData = ""
    Exit Function
End If

DayKey = GenerateKey

'store the amount of digits daykey is, in the first
'character.
RetData = Chr(Len(Trim(Str(DayKey))))
RetData = RetData & EncryptKey(Trim(Str(DayKey)))

'encrypt the rest of the data

TextLen = Len(Text)
Counter = 1
While Counter <= TextLen
    DoEvents
    
    'encrypt each character by adding the key value to the ascii value
    'of each character
    Encrypt = (Chr((Asc(Mid(Text, Counter, 1)) + DayKey) Mod 256))
    
    'save the encrypted character
    RetData = RetData & Encrypt
    Counter = Counter + 1
Wend

EncryptData = RetData
End Function

Public Function DecryptData(Text As String) As String
Dim Counter As Integer
Dim DayNum As String
Dim DayKey As Integer
Dim RetData As String
Dim Decrypt As String
Dim DecryptNum As Integer
Dim TextLen As Long

'get the amount of digits the key is and decrypt the
'key
If Text = "" Then
    Exit Function
End If

'get the amount of digits the key is
DayNum = GetKeyLength(Text)

'decrypt the key from the encrypted text
DayKey = Val(DecryptKey(Mid(Text, 2, Val(DayNum))))

'decrypt the rest of the text
TextLen = Len(Text)
Counter = (Val(DayNum) + 2)
While Counter <= TextLen
    DoEvents
    
    'subtract the key value from the ascii value of each character
    'and account for a negative result
    DecryptNum = (Asc(Mid(Text, Counter, 1)) - DayKey) Mod 256
    If DecryptNum < 0 Then
        DecryptNum = 255 + DecryptNum
    Else
        DecryptNum = DecryptNum Mod 256
    End If
    
    'the character has been decrypted, save the result
    Decrypt = Right(Chr(DecryptNum), 1)
    RetData = RetData & Decrypt

    'next character
    Counter = Counter + 1
Wend

'return the decrypted data
DecryptData = RetData
End Function

Public Function GetKeyLength(Text As String) As String
Dim KeyLength As Integer

'get the amount of digits the key is and decrypt the
'key
If Text = "" Then
    Exit Function
End If
    
KeyLength = Asc(Left(Text, 1))

GetKeyLength = KeyLength
End Function

Private Function EncryptKey(Key As String) As String
'adds the encryption key to the ASCII value of each
'character.

Dim Counter As Integer
Dim NewKey As String

On Error Resume Next

For Counter = 1 To Len(Key)
    NewKey = NewKey & Right(Chr(Asc(Mid(Key, Counter, 1)) + BaseKey), 1)
Next Counter

EncryptKey = NewKey
End Function

Private Function DecryptKey(Key As String) As String
'subtracts the encryption key from the ASCII value
'of each character.

Dim Counter As Integer
Dim NewKey As String
Dim test As Variant

On Error Resume Next

For Counter = 1 To Len(Key)
    NewKey = NewKey & Right(Chr(Asc(Mid(Key, Counter, 1)) - BaseKey), 1)
Next Counter

If Key = "" Then NewKey = ""

DecryptKey = NewKey
End Function

Public Sub FileEncrypt(SourcePath As String, DestPath As String)
'This procedure takes two arguments. The file you want encrypted
'and the name and destination of the file you want the encrypted
'data in.

Dim Buffer As String    '30K buffer
Dim ErrorNum As Integer
Dim FileNumOut As Integer
Dim FileNumIn As Integer

'check for errors accessing the files
'-------------------
If LCase(SourcePath) = LCase(DestPath) Then
    Exit Sub
End If

On Error Resume Next
FileNumIn = FreeFile

Open SourcePath For Input As #FileNumIn
    ErrorNum = Err
Close #FileNumIn

If ErrorNum <> 0 Then
    Exit Sub
End If

Open DestPath For Output As #FileNumIn
    ErrorNum = Err
Close #FileNumIn

If ErrorNum <> 0 Then
    Exit Sub
End If
On Error GoTo 0
'-------------------

FileNumIn = FreeFile
Open SourcePath For Binary As #FileNumIn
    FileNumOut = FreeFile
    
    Open DestPath For Binary As #FileNumOut
        While Not EOF(FileNumIn)
            'input a 30K chunk of the file to be encrypted
            Buffer = Input((FileIntake), #FileNumIn) '- (MaxKeyLen)
            
            'encrypt the information
            Buffer = EncryptData(Buffer)
            
            'save the encrypted information
            Put #FileNumOut, , Buffer
        Wend
    Close #FileNumOut
Close #FileNumIn

End Sub

Public Sub FileDecrypt(SourcePath As String, DestPath As String)
'This procedure takes two arguments. The file you want decrypted
'and the name and destination of the file you want the decrypted
'data in.

Const FileIntake = 30720

Dim KeyLenChar As String
Dim KeyLen As Integer
Dim EncryptedData As String
Dim DecryptedData As String
Dim ErrorNum As Integer
Dim FileNumOut As Integer
Dim FileNumIn As Integer

'check for errors accessing the files
'-------------------
If LCase(SourcePath) = LCase(DestPath) Then
    Exit Sub
End If

On Error Resume Next
FileNumIn = FreeFile

Open SourcePath For Input As #FileNumIn
    ErrorNum = Err
Close #FileNumIn

If ErrorNum <> 0 Then
    Exit Sub
End If

Open DestPath For Output As #FileNumIn
    ErrorNum = Err
Close #FileNumIn

If ErrorNum <> 0 Then
    Exit Sub
End If
On Error GoTo 0
'-------------------

'This decryption works by the following steps;
'1) Input the first character of the file. This will contain the length
'    of the decyption key.
'2) Input the decryption Key using the key length
'3) Input the next 30720 characters - this is the encrypted data
'4) Repeat steps 1, 2 and three until the entire file is read.

FileNumIn = FreeFile
Open SourcePath For Binary As #FileNumIn
    'frmFiles.barFile.Max = LOF(FileNumIn)
    
    FileNumOut = FreeFile
    Open DestPath For Binary As #FileNumOut
        While Not EOF(FileNumIn)
            'input the character with the keylength
            KeyLenChar = Input(1, #FileNumIn)
            
            'save the keylength
            KeyLen = GetKeyLength(KeyLenChar)
            
            'get the decryption key and the encrypted data
            EncryptedData = Input(KeyLen + FileIntake, #FileNumIn)
            
            'decrypt the info recovered
            DecryptedData = DecryptData(KeyLenChar & EncryptedData)
            
            'save decrypted data
            Put #FileNumOut, , DecryptedData
        Wend
    Close #FileNumOut
Close #FileNumIn


End Sub

